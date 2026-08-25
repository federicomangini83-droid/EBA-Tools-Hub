const csvFiles = document.getElementById("csvFiles");
const outputName = document.getElementById("outputName");
const processBtn = document.getElementById("processBtn");
const spinner = document.getElementById("spinner");
const processingStatus = document.getElementById("processingStatus");
const statusDot = document.getElementById("statusDot");
const runtimeStatus = document.getElementById("runtimeStatus");
const selectionSummary = document.getElementById("selectionSummary");
const selectedFiles = document.getElementById("selectedFiles");
const errorBox = document.getElementById("errorBox");
const resultBox = document.getElementById("resultBox");
const resultSummary = document.getElementById("resultSummary");
const stats = document.getElementById("stats");
const previewTable = document.getElementById("previewTable");
const downloadBtn = document.getElementById("downloadBtn");
const clearLogBtn = document.getElementById("clearLogBtn");
const logOutput = document.getElementById("logOutput");

const MAPPING_URL = "./data/EBA_reporting_framework.xlsx";
const ENGINE_URL = "./python/csv_merger.py";
const MAPPING_PATH = "/tmp/EBA_reporting_framework.xlsx";

let pyodide = null;
let mergedCsv = "";
let isProcessing = false;

function showError(message) {
  errorBox.textContent = message;
  errorBox.classList.remove("hidden");
}

function clearError() {
  errorBox.textContent = "";
  errorBox.classList.add("hidden");
}

function setRuntimeState(state, message) {
  statusDot.classList.toggle("ready", state === "ready");
  statusDot.classList.toggle("error-state", state === "error");
  runtimeStatus.textContent = message;
  updateProcessButton();
}

function updateProcessButton() {
  processBtn.disabled = !pyodide || csvFiles.files.length === 0 || isProcessing;
}

function setProcessing(value, message = "") {
  isProcessing = value;
  spinner.classList.toggle("hidden", !value);
  spinner.style.display = value ? "inline-block" : "none";
  processingStatus.textContent = message;
  processBtn.textContent = value ? "Processing..." : "Process files";
  updateProcessButton();
}

function createVirtualInput(index, originalName) {
  const cleaned = originalName.replace(/[^a-zA-Z0-9._-]/g, "_");
  const directory = `/tmp/eba_csv_input_${index}`;
  try { pyodide.FS.mkdir(directory); } catch (_) {}
  return { directory, path: `${directory}/${cleaned}` };
}

function normaliseOutputName() {
  let value = outputName.value.trim() || "EBA_Merged_Output.csv";
  if (!value.toLowerCase().endsWith(".csv")) value += ".csv";
  outputName.value = value;
  return value;
}

function updateSelectedFiles() {
  const files = Array.from(csvFiles.files);
  selectionSummary.textContent = files.length
    ? `${files.length} CSV file${files.length === 1 ? "" : "s"} selected.`
    : "No files selected.";

  selectedFiles.innerHTML = "";
  selectedFiles.classList.toggle("hidden", files.length === 0);
  if (files.length) {
    const list = document.createElement("ul");
    files.forEach((file) => {
      const item = document.createElement("li");
      item.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
      list.appendChild(item);
    });
    selectedFiles.appendChild(list);
  }
  updateProcessButton();
}

function renderLog(lines) {
  logOutput.textContent = lines.join("\n");
  logOutput.scrollTop = logOutput.scrollHeight;
}

function renderStats(result) {
  const values = [
    [result.row_count.toLocaleString(), "Output rows"],
    [result.mapped_rows.toLocaleString(), "Mapped rows"],
    [result.processed_files.toLocaleString(), "Files processed"]
  ];
  stats.innerHTML = "";
  values.forEach(([value, label]) => {
    const box = document.createElement("div");
    box.className = "stat";
    const strong = document.createElement("strong");
    strong.textContent = value;
    const caption = document.createElement("span");
    caption.textContent = label;
    box.append(strong, caption);
    stats.appendChild(box);
  });
}

function renderPreview(records, columns) {
  previewTable.innerHTML = "";
  if (!records.length) return;

  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  columns.forEach((column) => {
    const cell = document.createElement("th");
    cell.textContent = column;
    headRow.appendChild(cell);
  });
  head.appendChild(headRow);
  previewTable.appendChild(head);

  const body = document.createElement("tbody");
  records.forEach((record) => {
    const row = document.createElement("tr");
    columns.forEach((column) => {
      const cell = document.createElement("td");
      cell.textContent = record[column] ?? "";
      row.appendChild(cell);
    });
    body.appendChild(row);
  });
  previewTable.appendChild(body);
}

const runtimeReady = (async () => {
  try {
    setRuntimeState("loading", "Loading Pyodide...");
    pyodide = await loadPyodide({
      indexURL: "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/"
    });

    setRuntimeState("loading", "Loading pandas...");
    await pyodide.loadPackage(["pandas", "micropip"]);

    setRuntimeState("loading", "Installing Excel support...");
    const micropip = pyodide.pyimport("micropip");
    try {
      await micropip.install("openpyxl");
    } finally {
      micropip.destroy();
    }

    setRuntimeState("loading", "Loading reporting framework mapping...");
    const mappingResponse = await fetch(MAPPING_URL, { cache: "no-cache" });
    if (!mappingResponse.ok) {
      throw new Error(`Mapping workbook not found: HTTP ${mappingResponse.status}`);
    }
    pyodide.FS.writeFile(
      MAPPING_PATH,
      new Uint8Array(await mappingResponse.arrayBuffer())
    );

    const engineResponse = await fetch(ENGINE_URL, { cache: "no-cache" });
    if (!engineResponse.ok) {
      throw new Error(`Python engine not found: HTTP ${engineResponse.status}`);
    }
    await pyodide.runPythonAsync(await engineResponse.text());

    setRuntimeState("ready", "Engine ready");
  } catch (error) {
    console.error(error);
    pyodide = null;
    setRuntimeState("error", "Engine unavailable");
    showError(`Initialisation error: ${error.message}`);
    throw error;
  }
})();

runtimeReady.catch(() => {});

async function processFiles() {
  if (isProcessing || !pyodide) return;
  const files = Array.from(csvFiles.files);
  if (!files.length) {
    showError("Select at least one CSV file.");
    return;
  }

  clearError();
  resultBox.classList.add("hidden");
  mergedCsv = "";
  setProcessing(true, `Copying ${files.length} files to the Python environment...`);
  logOutput.textContent = "Preparing input files...";

  const virtualPaths = [];
  const virtualDirectories = [];
  try {
    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      const virtualInput = createVirtualInput(index, file.name);
      pyodide.FS.writeFile(
        virtualInput.path,
        new Uint8Array(await file.arrayBuffer())
      );
      virtualPaths.push(virtualInput.path);
      virtualDirectories.push(virtualInput.directory);
    }

    setProcessing(true, "Merging CSV files and applying the mapping...");
    pyodide.globals.set("input_paths_json", JSON.stringify(virtualPaths));
    pyodide.globals.set("mapping_path_from_js", MAPPING_PATH);

    const resultJson = await pyodide.runPythonAsync(
      "process_csv_files(input_paths_json, mapping_path_from_js)"
    );
    const result = JSON.parse(resultJson);

    mergedCsv = result.csv;
    renderLog(result.logs || []);
    renderStats(result);
    renderPreview(result.preview || [], result.columns || []);
    resultSummary.textContent =
      `${result.row_count.toLocaleString()} rows generated. ` +
      `${result.unmapped_rows.toLocaleString()} rows have no mapping coordinates.`;
    resultBox.classList.remove("hidden");
    setProcessing(false, "Processing completed.");
  } catch (error) {
    console.error(error);
    showError(`Processing error: ${error.message}`);
    setProcessing(false, "Processing failed.");
  } finally {
    virtualPaths.forEach((path) => {
      try { pyodide.FS.unlink(path); } catch (_) {}
    });
    virtualDirectories.forEach((directory) => {
      try { pyodide.FS.rmdir(directory); } catch (_) {}
    });
    try { pyodide.globals.delete("input_paths_json"); } catch (_) {}
    try { pyodide.globals.delete("mapping_path_from_js"); } catch (_) {}
  }
}

function downloadResult() {
  if (!mergedCsv) return;
  const url = URL.createObjectURL(
    new Blob(["\uFEFF", mergedCsv], { type: "text/csv;charset=utf-8" })
  );
  const link = document.createElement("a");
  link.href = url;
  link.download = normaliseOutputName();
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

csvFiles.addEventListener("change", updateSelectedFiles);
processBtn.addEventListener("click", processFiles);
downloadBtn.addEventListener("click", downloadResult);
clearLogBtn.addEventListener("click", () => {
  logOutput.textContent = "Waiting for input files...";
});
outputName.addEventListener("blur", normaliseOutputName);

updateSelectedFiles();
