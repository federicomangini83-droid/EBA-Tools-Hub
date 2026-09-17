import base64
import csv
import io
import json
import re
import zipfile

import openpyxl

FIELDS = [
    "Id_Tab",
    "Row",
    "Column",
    "Cod_Conto",
    "Cod_Dest2",
    "Cod_Dest3",
    "Cod_Dest4",
    "Cod_Dest5",
    "Cod_Categoria",
    "Formula",
    "Calculation_Logic",
    "DB_Storage_Sign",
    "EBA_Sign",
    "Coordinate",
]

# campi che nell'esplosione vengono presi dalla riga di dettaglio (foglia)
DETAIL_FIELDS = [
    "Cod_Conto",
    "Cod_Dest2",
    "Cod_Dest3",
    "Cod_Dest4",
    "Cod_Dest5",
    "Cod_Categoria",
]

MAX_DEPTH = 20


def rgb(cell):
    return getattr(cell.fill.fgColor, "rgb", None)


def between(text, start, end):
    start_pos = text.find(start)
    end_pos = text.find(end, start_pos + len(start))

    if start_pos < 0 or end_pos < 0:
        return None

    return (
        text[start_pos + len(start): end_pos]
        .strip()
        .strip("'")
        .strip()
    )


def last(text, marker):
    start_pos = text.find(marker)

    if start_pos < 0:
        return None

    return (
        text[start_pos + len(marker):]
        .strip()
        .strip("'")
        .strip()
    )


def parse(value):
    if value is None:
        return None

    text = str(value).replace("\n", "")

    account = "Account= "
    dest2 = "Dest 2 = "
    dest3 = "Dest 3 = "
    dest4 = "Dest 4 = "
    dest5 = "Dest 5 = "
    category = "Category= "
    formula = "Formula= "
    storage_sign = "DB Storage Sign= "
    eba_sign = "EBA Sign= "
    calc_logic = "Calculation logic= "

    result = {
        key: None
        for key in FIELDS[3:13]
    }

    if (
        account in text
        and formula not in text
        and dest2 in text
    ):
        result.update(
            Cod_Conto=between(text, account, dest2),
            Cod_Dest2=between(text, dest2, dest3),
            Cod_Dest3=between(text, dest3, dest4),
            Cod_Dest4=between(text, dest4, dest5),
            Cod_Dest5=between(text, dest5, category),
            Cod_Categoria=between(text, category, storage_sign),
            DB_Storage_Sign=between(
                text,
                storage_sign,
                eba_sign
            ),
            EBA_Sign=(
                between(text, eba_sign, calc_logic)
                if calc_logic in text
                else last(text, eba_sign)
            ),
            Calculation_Logic=(
                last(text, calc_logic)
                if calc_logic in text
                else None
            ),
        )

    elif (
        account not in text
        and formula in text
        and calc_logic not in text
    ):
        result.update(
            Formula=between(
                text,
                formula,
                storage_sign
            ),
            DB_Storage_Sign=between(
                text,
                storage_sign,
                eba_sign
            ),
            EBA_Sign=last(text, eba_sign),
        )

    else:
        return None

    return (
        result
        if result["EBA_Sign"] is not None
        else None
    )


# --------------------------------------------------------------------------- #
#  ESPLOSIONE
# --------------------------------------------------------------------------- #

def split_dest(value):
    """'FR_IFRS,FR_NGAAP' -> ['FR_IFRS', 'FR_NGAAP']."""
    if value is None:
        return [None]

    parts = [
        part.strip()
        for part in str(value).replace(";", ",").split(",")
        if part.strip()
    ]

    return parts or [None]


def explode_dest3(records):
    """Duplica la riga per ogni valore presente in Cod_Dest3."""
    exploded = []

    for record in records:
        for value in split_dest(record.get("Cod_Dest3")):
            copy = dict(record)
            copy["Cod_Dest3"] = value
            exploded.append(copy)

    return exploded


def formula_refs(formula):
    """SUM(R0020,R0030) -> ['r0020', 'r0030']. None se non e' una SUM."""
    if not formula:
        return []

    matches = re.findall(r"R\s*0*\d+", str(formula), flags=re.IGNORECASE)

    return [
        "r" + re.sub(r"\D", "", match).zfill(4)
        for match in matches
    ]


def row_key(record):
    return (
        record.get("Id_Tab"),
        record.get("Column"),
        str(record.get("Row") or "").strip().lower(),
    )


def has_formula(record):
    return bool(str(record.get("Formula") or "").strip())


def build_index(records):
    """Mappa (Id_Tab, Column, row) -> {'details': [...], 'refs': [...]}."""
    index = {}

    for record in records:
        key = row_key(record)
        node = index.setdefault(key, {"details": [], "refs": []})

        if has_formula(record):
            for ref in formula_refs(record["Formula"]):
                ref_key = (key[0], key[1], ref)

                if ref_key not in node["refs"]:
                    node["refs"].append(ref_key)
        else:
            node["details"].append(record)

    return index


def resolve_leaves(key, index, cache, visiting, depth=0):
    """Scende ricorsivamente fino alle righe senza formula."""
    if key in cache:
        return cache[key]

    node = index.get(key)

    if node is None or not node["refs"] or depth >= MAX_DEPTH:
        return [key] if node and node["details"] else []

    if key in visiting:          # riferimento circolare: taglio il ramo
        return []

    visiting.add(key)

    leaves = []

    for ref_key in node["refs"]:
        ref_node = index.get(ref_key)

        if ref_node is None:
            continue

        if ref_node["refs"]:
            children = resolve_leaves(
                ref_key,
                index,
                cache,
                visiting,
                depth + 1
            )
        else:
            children = [ref_key] if ref_node["details"] else []

        for child in children:
            if child not in leaves:      # dedup sui rami convergenti
                leaves.append(child)

    visiting.discard(key)
    cache[key] = leaves

    return leaves


def explode_records(records):
    """Dest3 esploso + SUM risolte fino alle foglie, senza duplicati."""
    base = explode_dest3(records)
    index = build_index(base)

    cache = {}
    output = []
    seen = set()

    for record in base:

        if not has_formula(record):
            output.append(record)
            continue

        leaves = resolve_leaves(row_key(record), index, cache, set())

        for leaf_key in leaves:
            for detail in index[leaf_key]["details"]:

                row = dict(record)

                for field in DETAIL_FIELDS:
                    row[field] = detail.get(field)

                signature = tuple(row.get(field) for field in FIELDS[:-1])

                if signature in seen:
                    continue

                seen.add(signature)
                output.append(row)

    return output


# --------------------------------------------------------------------------- #
#  OUTPUT
# --------------------------------------------------------------------------- #

def to_csv(records):
    buffer = io.StringIO(newline="")

    writer = csv.DictWriter(
        buffer,
        fieldnames=FIELDS,
        lineterminator="\n",
        extrasaction="ignore",
    )

    writer.writeheader()
    writer.writerows(records)

    return buffer.getvalue()


def to_zip(raw_csv, exploded_csv):
    buffer = io.BytesIO()

    with zipfile.ZipFile(
        buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as archive:
        archive.writestr("mapping_raw.csv", raw_csv)
        archive.writestr("mapping_exploded.csv", exploded_csv)

    return buffer.getvalue()


def extract_records(path):
    workbook = openpyxl.load_workbook(
        path,
        data_only=False
    )

    sheet_names = workbook.sheetnames

    if not sheet_names:
        raise ValueError(
            "The Excel file does not contain valid worksheets."
        )

    reference_sheet = workbook[sheet_names[0]]

    color = rgb(reference_sheet["A1"])
    row_style = reference_sheet["A2"]._style
    column_style = reference_sheet["B1"]._style

    columns = {}
    rows = {}

    max_columns = {}
    max_rows = {}

    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]

        current_column = 1
        current_row = 1

        for row in sheet.iter_rows():
            for cell in row:

                if cell.column == 1:
                    continue

                if (
                    cell.value is not None
                    and (
                        cell._style == column_style
                        or rgb(cell) == color
                    )
                ):
                    current_column += 1

                    columns[(sheet_name, current_column)] = (
                        "c" + str(cell.value).zfill(4)
                    )

                elif cell.value is None:
                    break

            break

        for column in sheet.iter_cols():
            for cell in column:

                if cell.row == 1:
                    continue

                if (
                    cell.value is not None
                    and (
                        cell._style == row_style
                        or rgb(cell) == color
                    )
                ):
                    current_row += 1

                    rows[(sheet_name, current_row)] = (
                        "r" + str(cell.value).zfill(4)
                    )

                elif cell.value is None:
                    break

            break

        max_columns[sheet_name] = current_column
        max_rows[sheet_name] = current_row

    records = []

    for sheet_name in sheet_names:

        if sheet_name in {
            "Test_Formula",
            "Macro"
        }:
            continue

        sheet = workbook[sheet_name]

        for row_range in sheet.iter_rows(
            min_row=2,
            max_row=max_rows[sheet_name],
            min_col=2,
            max_col=max_columns[sheet_name]
        ):
            for cell in row_range:

                if rgb(cell) == color:
                    continue

                parsed = parse(cell.value)

                if parsed:
                    records.append(
                        {
                            "Id_Tab": sheet_name,
                            "Row": rows.get(
                                (sheet_name, cell.row),
                                cell.row
                            ),
                            "Column": columns.get(
                                (sheet_name, cell.column),
                                cell.column
                            ),
                            **parsed,
                            "Coordinate": cell.coordinate,
                        }
                    )

    return records


def process_workbook(path, zip_path=None):
    records = extract_records(path)
    exploded = explode_records(records)

    raw_csv = to_csv(records)
    exploded_csv = to_csv(exploded)

    archive = to_zip(raw_csv, exploded_csv)

    if zip_path:
        with open(zip_path, "wb") as handle:
            handle.write(archive)

    return json.dumps(
        {
            "zip_base64": base64.b64encode(archive).decode("ascii"),
            "zip_name": "mapping.zip",
            "files": ["mapping_raw.csv", "mapping_exploded.csv"],
            "count": len(records),
            "count_exploded": len(exploded),
        },
        ensure_ascii=False,
    )


if __name__ == "__main__":
    import sys

    print(
        process_workbook(
            sys.argv[1],
            sys.argv[2] if len(sys.argv) > 2 else "mapping.zip",
        )[:400]
    )
