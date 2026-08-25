const $ = (id) => document.getElementById(id);
const csvFiles = $("csvFiles");
const csvFolder = $("csvFolder");
const outputName = $("outputName");
const processBtn = $("processBtn");
const spinner = $("spinner");
const processingStatus = $("processingStatus");
const statusDot = $("statusDot");
const runtimeStatus = $("runtimeStatus");
const selectionSummary = $("selectionSummary");
const selectedFilesBox = $("selectedFiles");
const errorBox = $("errorBox");
const resultBox = $("resultBox");
const resultSummary = $("resultSummary");
const stats = $("stats");
const previewTable = $("previewTable");
const downloadBtn = $("downloadBtn");
const clearLogBtn = $("clearLogBtn");
const logOutput = $("logOutput");

const MAPPING_URL = "./data/EBA_reporting_framework.xlsx";
const ENGINE_URL = "./python/csv_merger.py";
const MAPPING_PATH = "/tmp/EBA_reporting_framework.xlsx";
let pyodide = null;
let mergedCsv = "";
let isProcessing = false;
let selectedFileList = [];

const requiredElements = {
    csvFiles,
    csvFolder,
    outputName,
    processBtn,
    spinner,
    processingStatus,
    statusDot,
    runtimeStatus,
    selectionSummary,
    selectedFilesBox,
    errorBox,
    resultBox,
    resultSummary,
    stats,
    previewTable,
    downloadBtn,
    clearLogBtn,
    logOutput
};

const missingElements = Object.entries(requiredElements)
    .filter(([, element]) => !element)
    .map(([name]) => name);

if (missingElements.length) {
    throw new Error(`HTML/JavaScript mismatch. Missing elements: ${missingElements.join(", ")}`);
}

function isCsv(file) {
    return file.name.trim().toLowerCase().endsWith(".csv");
}
function showError(message) {
    if (!errorBox) { console.error(message); return; }
    errorBox.textContent = message;
    errorBox.classList.remove("hidden");
}
function clearError() {
    if (!errorBox) return;
    errorBox.textContent = "";
    errorBox.classList.add("hidden");
}
function updateProcessButton() { processBtn.disabled = !pyodide || !selectedFileList.length || isProcessing; }
function setRuntimeState(state, message) {
    statusDot.classList.toggle("ready", state === "ready");
    statusDot.classList.toggle("error-state", state === "error");
    runtimeStatus.textContent = message;
    updateProcessButton();
}
function setProcessing(value, message = "") {
    isProcessing = value;
    spinner.classList.toggle("hidden", !value);
    processingStatus.textContent = message;
    processBtn.textContent = value ? "Processing..." : "Process files";
    updateProcessButton();
}
function normaliseOutputName() {
    let value = outputName.value.trim() || "EBA_Merged_Output.csv";
    if (!value.toLowerCase().endsWith(".csv")) value += ".csv";
    outputName.value = value;
    return value;
}
function refreshSelection(fileList, sourceLabel) {
    const received = Array.from(fileList || []);
    selectedFileList = received.filter(isCsv);
    selectionSummary.textContent = selectedFileList.length
        ? `${selectedFileList.length} CSV file${selectedFileList.length === 1 ? "" : "s"} selected from ${sourceLabel}.`
        : "No files selected.";
    selectedFilesBox.innerHTML = "";
    selectedFilesBox.classList.toggle("hidden", !selectedFileList.length);
    if (selectedFileList.length) {
        const list = document.createElement("ul");
        selectedFileList.forEach((file) => {
            const item = document.createElement("li");
            item.textContent = `${file.webkitRelativePath || file.name} (${(file.size / 1024).toFixed(1)} KB)`;
            list.appendChild(item);
        });
        selectedFilesBox.appendChild(list);
        clearError();
    } else {
        showError(`The selected ${sourceLabel} contains ${received.length} file(s), but no files ending in .csv were found.`);
    }
    updateProcessButton();
}
function renderLog(lines) { logOutput.textContent = lines.join("\n"); }
function renderStats(result) {
    stats.innerHTML = "";
    [[result.row_count, "Output rows"], [result.mapped_rows, "Mapped rows"], [result.processed_files, "Files processed"]].forEach(([value, label]) => {
        const box = document.createElement("div"); box.className = "stat";
        const strong = document.createElement("strong"); strong.textContent = Number(value).toLocaleString();
        const span = document.createElement("span"); span.textContent = label;
        box.append(strong, span); stats.appendChild(box);
    });
}
function renderPreview(records, columns) {
    previewTable.innerHTML = "";
    if (!records.length) return;
    const thead = document.createElement("thead"); const hr = document.createElement("tr");
    columns.forEach((column) => { const th = document.createElement("th"); th.textContent = column; hr.appendChild(th); });
    thead.appendChild(hr); previewTable.appendChild(thead);
    const tbody = document.createElement("tbody");
    records.forEach((record) => { const row = document.createElement("tr"); columns.forEach((column) => { const td = document.createElement("td"); td.textContent = record[column] ?? ""; row.appendChild(td); }); tbody.appendChild(row); });
    previewTable.appendChild(tbody);
}

(async () => {
    try {
        setRuntimeState("loading", "Loading Pyodide...");
        pyodide = await loadPyodide({ indexURL: "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/" });
        setRuntimeState("loading", "Loading pandas...");
        await pyodide.loadPackage(["pandas", "micropip"]);
        const micropip = pyodide.pyimport("micropip");
        try { await micropip.install("openpyxl"); } finally { micropip.destroy(); }
        setRuntimeState("loading", "Loading reporting framework mapping...");
        const mappingResponse = await fetch(MAPPING_URL, { cache: "no-cache" });
        if (!mappingResponse.ok) throw new Error(`Mapping workbook not found: HTTP ${mappingResponse.status}`);
        pyodide.FS.writeFile(MAPPING_PATH, new Uint8Array(await mappingResponse.arrayBuffer()));
        const engineResponse = await fetch(ENGINE_URL, { cache: "no-cache" });
        if (!engineResponse.ok) throw new Error(`Python engine not found: HTTP ${engineResponse.status}`);
        await pyodide.runPythonAsync(await engineResponse.text());
        setRuntimeState("ready", "Engine ready");
    } catch (error) {
        pyodide = null; setRuntimeState("error", "Engine unavailable"); showError(`Initialisation error: ${error.message}`); console.error(error);
    }
})();

async function processFiles() {
    if (!pyodide || isProcessing || !selectedFileList.length) return;
    clearError(); resultBox.classList.add("hidden"); mergedCsv = "";
    setProcessing(true, `Copying ${selectedFileList.length} files...`);
    const paths = []; const directories = [];
    try {
        for (let index = 0; index < selectedFileList.length; index += 1) {
            const file = selectedFileList[index];
            const directory = `/tmp/eba_csv_input_${index}`;
            try { pyodide.FS.mkdir(directory); } catch (_) {}
            const cleanName = file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
            const path = `${directory}/${cleanName}`;
            pyodide.FS.writeFile(path, new Uint8Array(await file.arrayBuffer()));
            paths.push(path); directories.push(directory);
        }
        setProcessing(true, "Merging CSV files and applying the mapping...");
        pyodide.globals.set("input_paths_json", JSON.stringify(paths));
        pyodide.globals.set("mapping_path_from_js", MAPPING_PATH);
        const result = JSON.parse(await pyodide.runPythonAsync("process_csv_files(input_paths_json, mapping_path_from_js)"));
        mergedCsv = result.csv;
        renderLog(result.logs || []); renderStats(result); renderPreview(result.preview || [], result.columns || []);
        if (resultSummary) {
            resultSummary.textContent = `${Number(result.row_count).toLocaleString()} rows generated. ${Number(result.unmapped_rows).toLocaleString()} rows have no mapping coordinates.`;
        }
        resultBox.classList.remove("hidden"); setProcessing(false, "Processing completed.");
    } catch (error) {
        showError(`Processing error: ${error.message}`); setProcessing(false, "Processing failed."); console.error(error);
    } finally {
        paths.forEach((path) => { try { pyodide.FS.unlink(path); } catch (_) {} });
        directories.forEach((directory) => { try { pyodide.FS.rmdir(directory); } catch (_) {} });
        try { pyodide.globals.delete("input_paths_json"); } catch (_) {}
        try { pyodide.globals.delete("mapping_path_from_js"); } catch (_) {}
    }
}
function downloadResult() {
    if (!mergedCsv) return;
    const url = URL.createObjectURL(new Blob(["\uFEFF", mergedCsv], { type: "text/csv;charset=utf-8" }));
    const link = document.createElement("a"); link.href = url; link.download = normaliseOutputName();
    document.body.appendChild(link); link.click(); link.remove(); URL.revokeObjectURL(url);
}

if ("webkitdirectory" in csvFolder) csvFolder.webkitdirectory = true;
csvFiles.addEventListener("change", (event) => { csvFolder.value = ""; refreshSelection(event.target.files, "file selection"); });
csvFolder.addEventListener("change", (event) => { csvFiles.value = ""; refreshSelection(event.target.files, "folder"); });
processBtn.addEventListener("click", processFiles);
downloadBtn.addEventListener("click", downloadResult);
clearLogBtn.addEventListener("click", () => { logOutput.textContent = "Waiting for input files..."; });
outputName.addEventListener("blur", normaliseOutputName);
updateProcessButton();

const processingLogCard =
    document.getElementById("processingLogCard");

if (processingLogCard) {
    processingLogCard.style.display = "none";
}
