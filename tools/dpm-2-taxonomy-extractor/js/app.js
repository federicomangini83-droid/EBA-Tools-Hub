/* ==========================================================================
   DOM references
   ========================================================================== */

const $ = (id) => document.getElementById(id);

const jsonFolder = $("jsonFolder");
const excludeList = $("excludeList");
const specificList = $("specificList");
const filePattern = $("filePattern");
const outputName = $("outputName");
const matchCount = $("matchCount");

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
const sheetsTable = $("sheetsTable");
const previewTable = $("previewTable");
const downloadBtn = $("downloadBtn");
const saveGithubBtn = $("saveGithubBtn");
const archiveStatus = $("archiveStatus");

const refreshHistoryBtn = $("refreshHistoryBtn");
const historyTable = $("historyTable");
const historyError = $("historyError");
const ghOwner = $("ghOwner");
const ghRepo = $("ghRepo");
const ghBranch = $("ghBranch");
const ghFolder = $("ghFolder");
const ghToken = $("ghToken");
const ghRemember = $("ghRemember");

const clearLogBtn = $("clearLogBtn");
const logOutput = $("logOutput");
const processingLogCard = $("processingLogCard");

// Hide the processing log immediately, independently of the stylesheet cache.
if (processingLogCard) {
    processingLogCard.style.display = "none";
}

/* ==========================================================================
   Constants and state
   ========================================================================== */

const ENGINE_URL = "./python/taxonomy_extractor.py";
const WORK_ROOT = "/tmp/dpm_taxonomy";
const OUTPUT_PATH = "/tmp/dpm_output.xlsx";

const TOKEN_KEY = "eba_tools_hub_github_token";
const POLL_INTERVAL = 3000;
const DEPLOY_TIMEOUT = 180000;

const XLSX_MIME =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

let pyodide = null;
let selectedFileList = [];
let workbookBytes = null;
let isProcessing = false;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/* ==========================================================================
   Helpers
   ========================================================================== */

function isJson(file) {
    return file.name.trim().toLowerCase().endsWith(".json");
}

function relativePath(file) {
    return file.webkitRelativePath || file.name;
}

function showError(message) {
    if (!errorBox) {
        console.error(message);
        return;
    }
    errorBox.textContent = message;
    errorBox.classList.remove("hidden");
}

function clearError() {
    if (!errorBox) return;
    errorBox.textContent = "";
    errorBox.classList.add("hidden");
}

function sizeLabel(size) {
    if (size < 1024) return `${size} B`;
    if (size < 1048576) return `${(size / 1024).toFixed(1)} KB`;
    return `${(size / 1048576).toFixed(1)} MB`;
}

function parseList(value) {
    return value
        .split(",")
        .map((item) => item.trim())
        .filter(Boolean);
}

function normaliseOutputName() {
    let value = outputName.value.trim() || "EBA_reporting_framework.xlsx";
    if (!value.toLowerCase().endsWith(".xlsx")) value += ".xlsx";
    outputName.value = value;
    return value;
}

function updateProcessButton() {
    processBtn.disabled = !pyodide || !selectedFileList.length || isProcessing;
}

function setRuntimeState(state, message) {
    statusDot.classList.toggle("ready", state === "ready");
    statusDot.classList.toggle("error-state", state === "error");
    runtimeStatus.textContent = message;
    updateProcessButton();
}

function setProcessing(value, message = "") {
    isProcessing = value;
    spinner.classList.toggle("hidden", !value);
    spinner.style.display = value ? "inline-block" : "none";
    processingStatus.textContent = message;
    processBtn.textContent = value ? "Processing..." : "Process taxonomy";
    updateProcessButton();
}

/* ==========================================================================
   Folder selection and client-side match preview
   ========================================================================== */

/*
 * The definitive filtering is performed by the Python engine.
 * This preview mirrors the same rules so the user sees the match count
 * before starting a potentially long extraction.
 */
function countMatches() {
    const exclude = parseList(excludeList.value).map((v) => v.toLowerCase());
    const specific = parseList(specificList.value).map((v) => v.toLowerCase());

    let pattern = null;
    try {
        // Translate the Python inline flag (?i) into a JavaScript flag.
        const raw = filePattern.value.trim();
        const caseInsensitive = raw.startsWith("(?i)");
        const body = caseInsensitive ? raw.slice(4) : raw;
        pattern = new RegExp(body, caseInsensitive ? "i" : "");
    } catch (_) {
        return null;
    }

    return selectedFileList.filter((file) => {
        const path = relativePath(file);
        const lower = path.toLowerCase();

        const matchesPattern = pattern.test(path);
        const notExcluded = !exclude.some((value) => lower.includes(value));
        const isSpecific =
            !specific.length || specific.some((value) => lower.includes(value));

        return matchesPattern && notExcluded && isSpecific;
    }).length;
}

function updateMatchCount() {
    const matches = countMatches();
    matchCount.textContent = matches === null ? "invalid pattern" : matches.toLocaleString();
}

function refreshSelection(fileList) {
    const received = Array.from(fileList || []);
    selectedFileList = received.filter(isJson);

    selectionSummary.textContent = selectedFileList.length
        ? `${selectedFileList.length.toLocaleString()} JSON file${selectedFileList.length === 1 ? "" : "s"} found in the selected folder.`
        : "No folder selected.";

    selectedFilesBox.innerHTML = "";
    selectedFilesBox.classList.toggle("hidden", !selectedFileList.length);

    if (selectedFileList.length) {
        const list = document.createElement("ul");
        selectedFileList.slice(0, 50).forEach((file) => {
            const item = document.createElement("li");
            item.textContent = `${relativePath(file)} (${sizeLabel(file.size)})`;
            list.appendChild(item);
        });

        if (selectedFileList.length > 50) {
            const more = document.createElement("li");
            more.textContent = `... and ${(selectedFileList.length - 50).toLocaleString()} more files`;
            list.appendChild(more);
        }

        selectedFilesBox.appendChild(list);
        clearError();
    } else {
        showError(
            `The selected folder contains ${received.length.toLocaleString()} file(s), but no files ending in .json were found.`
        );
    }

    updateMatchCount();
    updateProcessButton();
}

/* ==========================================================================
   Result rendering
   ========================================================================== */

function renderLog(lines) {
    if (logOutput) logOutput.textContent = lines.join("\n");
}

function renderStats(result) {
    stats.innerHTML = "";
    [
        [result.processed_files, "Files processed"],
        [result.mapping_rows, "Mapping rows"],
        [result.sheets.length, "Sheets generated"]
    ].forEach(([value, label]) => {
        const box = document.createElement("div");
        box.className = "stat";

        const strong = document.createElement("strong");
        strong.textContent = Number(value).toLocaleString();

        const span = document.createElement("span");
        span.textContent = label;

        box.append(strong, span);
        stats.appendChild(box);
    });
}

function renderSheets(sheets) {
    sheetsTable.innerHTML = "";

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    ["Sheet", "Rows", "Columns"].forEach((label) => {
        const th = document.createElement("th");
        th.textContent = label;
        headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    sheetsTable.appendChild(thead);

    const tbody = document.createElement("tbody");
    sheets.forEach((sheet) => {
        const row = document.createElement("tr");

        const name = document.createElement("td");
        name.textContent = sheet.sheet;

        const rows = document.createElement("td");
        rows.textContent = Number(sheet.rows).toLocaleString();

        const columns = document.createElement("td");
        columns.textContent = Number(sheet.columns).toLocaleString();

        row.append(name, rows, columns);
        tbody.appendChild(row);
    });
    sheetsTable.appendChild(tbody);
}

function renderPreview(records, columns) {
    previewTable.innerHTML = "";
    if (!records.length) return;

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    columns.forEach((column) => {
        const th = document.createElement("th");
        th.textContent = column;
        headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    previewTable.appendChild(thead);

    const tbody = document.createElement("tbody");
    records.forEach((record) => {
        const row = document.createElement("tr");
        columns.forEach((column) => {
            const cell = document.createElement("td");
            cell.textContent = record[column] ?? "";
            row.appendChild(cell);
        });
        tbody.appendChild(row);
    });
    previewTable.appendChild(tbody);
}

/* ==========================================================================
   Pyodide virtual filesystem helpers
   ========================================================================== */

function ensureDirectory(path) {
    const parts = path.split("/").filter(Boolean);
    let current = "";
    parts.forEach((part) => {
        current += `/${part}`;
        try { pyodide.FS.mkdir(current); } catch (_) {}
    });
}

function removeDirectoryTree(path) {
    let entries = [];
    try { entries = pyodide.FS.readdir(path); } catch (_) { return; }

    entries.forEach((entry) => {
        if (entry === "." || entry === "..") return;
        const child = `${path}/${entry}`;
        const isDirectory = pyodide.FS.isDir(pyodide.FS.stat(child).mode);
        if (isDirectory) {
            removeDirectoryTree(child);
        } else {
            try { pyodide.FS.unlink(child); } catch (_) {}
        }
    });

    try { pyodide.FS.rmdir(path); } catch (_) {}
}

/* ==========================================================================
   GitHub Contents API
   ========================================================================== */

function getConfig() {
    return {
        owner: ghOwner.value.trim(),
        repo: ghRepo.value.trim(),
        branch: ghBranch.value.trim() || "main",
        folder: ghFolder.value.trim().replace(/^\/+|\/+$/g, ""),
        token: ghToken.value.trim()
    };
}

function apiBase(config) {
    return `https://api.github.com/repos/${config.owner}/${config.repo}/contents/${config.folder}`;
}

function apiHeaders(config, authenticated = false) {
    const headers = {
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    };
    if (authenticated && config.token) {
        headers.Authorization = `Bearer ${config.token}`;
    }
    return headers;
}

function bytesToBase64(bytes) {
    let binary = "";
    const chunkSize = 8192;
    for (let index = 0; index < bytes.length; index += chunkSize) {
        binary += String.fromCharCode.apply(
            null,
            bytes.subarray(index, index + chunkSize)
        );
    }
    return btoa(binary);
}

function publishedUrl(config, fileName) {
    return `${location.origin}/${config.repo}/${config.folder}/${encodeURIComponent(fileName)}`;
}

async function getStoredFile(config, fileName) {
    const url = `${apiBase(config)}/${encodeURIComponent(fileName)}?ref=${encodeURIComponent(config.branch)}&v=${Date.now()}`;
    const response = await fetch(url, {
        headers: apiHeaders(config, true),
        cache: "no-store"
    });

    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`File lookup failed: HTTP ${response.status}`);
    return response.json();
}

async function saveFile(config, fileName, bytes) {
    const previous = await getStoredFile(config, fileName);

    const body = {
        message: `${previous ? "Update" : "Add"} taxonomy workbook: ${fileName}`,
        content: bytesToBase64(bytes),
        branch: config.branch
    };

    if (previous?.sha) body.sha = previous.sha;

    const response = await fetch(`${apiBase(config)}/${encodeURIComponent(fileName)}`, {
        method: "PUT",
        headers: { ...apiHeaders(config, true), "Content-Type": "application/json" },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        throw new Error(`Save failed: HTTP ${response.status} ${await response.text()}`);
    }

    return { wasUpdate: Boolean(previous) };
}

async function deleteFile(config, item) {
    const response = await fetch(`${apiBase(config)}/${encodeURIComponent(item.name)}`, {
        method: "DELETE",
        headers: { ...apiHeaders(config, true), "Content-Type": "application/json" },
        body: JSON.stringify({
            message: `Delete taxonomy workbook: ${item.name}`,
            sha: item.sha,
            branch: config.branch
        })
    });

    if (!response.ok) {
        throw new Error(`Delete failed: HTTP ${response.status} ${await response.text()}`);
    }
}

/*
 * Binary files cannot be compared byte by byte reliably through the CDN,
 * so the deployment check verifies availability and reported size instead.
 */
async function waitForDeployment(config, fileName, shouldExist, expectedSize = 0) {
    const started = Date.now();

    while (Date.now() - started < DEPLOY_TIMEOUT) {
        const response = await fetch(`${publishedUrl(config, fileName)}?v=${Date.now()}`, {
            cache: "no-store"
        });

        if (!shouldExist && response.status === 404) return;

        if (shouldExist && response.ok) {
            const blob = await response.blob();
            if (!expectedSize || blob.size === expectedSize) return;
        }

        await sleep(POLL_INTERVAL);
    }

    throw new Error("GitHub Pages did not complete the deployment within 3 minutes.");
}

async function downloadStoredFile(item) {
    const config = getConfig();
    const response = await fetch(`${publishedUrl(config, item.name)}?v=${Date.now()}`, {
        cache: "no-store"
    });

    if (!response.ok) throw new Error(`Download failed: HTTP ${response.status}`);

    const url = URL.createObjectURL(await response.blob());
    const link = document.createElement("a");
    link.href = url;
    link.download = item.name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

/* ==========================================================================
   History
   ========================================================================== */

function createHistoryRow(item) {
    const row = document.createElement("tr");

    const nameCell = document.createElement("td");
    nameCell.textContent = item.name;

    const sizeCell = document.createElement("td");
    sizeCell.textContent = sizeLabel(item.size || 0);

    const actionCell = document.createElement("td");
    const actions = document.createElement("div");
    actions.className = "row-actions";

    const downloadButton = document.createElement("button");
    downloadButton.textContent = "Download";
    downloadButton.className = "secondary";
    downloadButton.onclick = async () => {
        downloadButton.disabled = true;
        downloadButton.textContent = "Downloading...";
        try {
            await downloadStoredFile(item);
        } catch (error) {
            alert(error.message);
        } finally {
            downloadButton.disabled = false;
            downloadButton.textContent = "Download";
        }
    };

    const deleteButton = document.createElement("button");
    deleteButton.textContent = "Delete";
    deleteButton.className = "danger";
    deleteButton.onclick = async () => {
        const config = getConfig();

        if (!config.token) {
            alert("Enter the GitHub token first.");
            return;
        }

        if (!confirm(`Delete ${item.name}?`)) return;

        downloadButton.disabled = true;
        deleteButton.disabled = true;
        deleteButton.textContent = "Deleting...";
        row.style.opacity = ".65";
        archiveStatus.textContent = `Deleting ${item.name}: creating the GitHub commit...`;

        try {
            await deleteFile(config, item);
            deleteButton.textContent = "Waiting for deployment...";
            archiveStatus.textContent = `Commit created. Waiting for ${item.name} to be removed...`;
            await waitForDeployment(config, item.name, false);
            location.reload();
        } catch (error) {
            alert(error.message);
            downloadButton.disabled = false;
            deleteButton.disabled = false;
            deleteButton.textContent = "Delete";
            row.style.opacity = "";
            archiveStatus.textContent = `Deletion not completed: ${error.message}`;
        }
    };

    actions.append(downloadButton, deleteButton);
    actionCell.appendChild(actions);
    row.append(nameCell, sizeCell, actionCell);
    return row;
}

async function loadHistory() {
    const body = historyTable.querySelector("tbody");
    body.innerHTML = "";
    historyError.classList.add("hidden");

    const config = getConfig();
    if (!config.owner || !config.repo) return;

    refreshHistoryBtn.disabled = true;
    refreshHistoryBtn.textContent = "Refreshing...";

    try {
        const url = `${apiBase(config)}?ref=${encodeURIComponent(config.branch)}&v=${Date.now()}`;
        let response = await fetch(url, {
            headers: apiHeaders(config, Boolean(config.token)),
            cache: "no-store"
        });

        // A stale or revoked token must not block the public listing.
        if (response.status === 401) {
            response = await fetch(url, {
                headers: apiHeaders(config, false),
                cache: "no-store"
            });
        }

        const items = response.status === 404 ? [] : await response.json();

        if (!response.ok && response.status !== 404) {
            throw new Error(`History unavailable: HTTP ${response.status}`);
        }

        const workbooks = (Array.isArray(items) ? items : [])
            .filter((item) => item.type === "file" && item.name.toLowerCase().endsWith(".xlsx"))
            .sort((a, b) => a.name.localeCompare(b.name));

        workbooks.forEach((item) => body.appendChild(createHistoryRow(item)));

        if (!workbooks.length) {
            const row = document.createElement("tr");
            const cell = document.createElement("td");
            cell.colSpan = 3;
            cell.textContent = "No workbooks are currently stored.";
            row.appendChild(cell);
            body.appendChild(row);
        }
    } catch (error) {
        historyError.textContent = error.message;
        historyError.classList.remove("hidden");
    } finally {
        refreshHistoryBtn.disabled = false;
        refreshHistoryBtn.textContent = "Refresh list";
    }
}

async function saveCurrentToGithub() {
    if (!workbookBytes) {
        archiveStatus.textContent = "Generate a workbook first.";
        return;
    }

    const config = getConfig();

    if (!config.owner || !config.repo) {
        archiveStatus.textContent = "Configure the GitHub owner and repository.";
        return;
    }

    if (!config.token) {
        archiveStatus.textContent = "Enter the GitHub token to save to storage.";
        return;
    }

    const fileName = normaliseOutputName();
    const originalText = saveGithubBtn.textContent;

    saveGithubBtn.disabled = true;
    saveGithubBtn.textContent = "Saving...";
    archiveStatus.textContent = "Creating the GitHub commit...";

    try {
        const result = await saveFile(config, fileName, workbookBytes);

        saveGithubBtn.textContent = "Waiting for deployment...";
        archiveStatus.textContent = `Commit created. Waiting for ${fileName} to be published...`;

        await waitForDeployment(config, fileName, true, workbookBytes.length);

        archiveStatus.textContent = result.wasUpdate
            ? `Workbook overwritten and published: ${fileName}. Reloading page...`
            : `Workbook saved and published: ${fileName}. Reloading page...`;

        location.reload();
    } catch (error) {
        console.error(error);
        archiveStatus.textContent = `Save not completed: ${error.message}`;
        saveGithubBtn.disabled = false;
        saveGithubBtn.textContent = originalText;
    }
}

/* ==========================================================================
   Token persistence
   ========================================================================== */

function loadSavedToken() {
    try {
        const token = localStorage.getItem(TOKEN_KEY);
        if (token) {
            ghToken.value = token;
            ghRemember.checked = true;
        }
    } catch (_) {}
}

/* ==========================================================================
   Python runtime
   ========================================================================== */

(async () => {
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
            // xlsxwriter reproduces the original formatting; openpyxl is the fallback.
            try {
                await micropip.install("xlsxwriter");
            } catch (_) {
                console.warn("xlsxwriter unavailable, falling back to openpyxl.");
            }
        } finally {
            micropip.destroy();
        }

        setRuntimeState("loading", "Loading the extraction engine...");
        const engineResponse = await fetch(ENGINE_URL, { cache: "no-cache" });
        if (!engineResponse.ok) {
            throw new Error(`Python engine not found: HTTP ${engineResponse.status}`);
        }
        await pyodide.runPythonAsync(await engineResponse.text());

        setRuntimeState("ready", "Engine ready");
    } catch (error) {
        pyodide = null;
        setRuntimeState("error", "Engine unavailable");
        showError(`Initialisation error: ${error.message}`);
        console.error(error);
    }
})();

/* ==========================================================================
   Processing
   ========================================================================== */

async function processTaxonomy() {
    if (!pyodide || isProcessing || !selectedFileList.length) return;

    clearError();
    resultBox.classList.add("hidden");
    workbookBytes = null;
    archiveStatus.textContent = "";

    setProcessing(true, `Copying ${selectedFileList.length.toLocaleString()} files...`);

    const virtualPaths = [];

    try {
        removeDirectoryTree(WORK_ROOT);
        ensureDirectory(WORK_ROOT);

        for (let index = 0; index < selectedFileList.length; index += 1) {
            const file = selectedFileList[index];
            const relative = relativePath(file).replace(/[^a-zA-Z0-9._/-]/g, "_");
            const fullPath = `${WORK_ROOT}/${relative}`;
            const directory = fullPath.slice(0, fullPath.lastIndexOf("/"));

            ensureDirectory(directory);
            pyodide.FS.writeFile(fullPath, new Uint8Array(await file.arrayBuffer()));
            virtualPaths.push(fullPath);

            if (index % 200 === 0) {
                setProcessing(
                    true,
                    `Copying files: ${index.toLocaleString()} / ${selectedFileList.length.toLocaleString()}`
                );
                await sleep(0);
            }
        }

        setProcessing(true, "Extracting taxonomy tables...");

        pyodide.globals.set("file_paths_json", JSON.stringify(virtualPaths));
        pyodide.globals.set("exclude_json", JSON.stringify(parseList(excludeList.value)));
        pyodide.globals.set("specific_json", JSON.stringify(parseList(specificList.value)));
        pyodide.globals.set("output_path_from_js", OUTPUT_PATH);
        pyodide.globals.set("pattern_from_js", filePattern.value.trim());

        const result = JSON.parse(
            await pyodide.runPythonAsync(
                "process_taxonomy(file_paths_json, exclude_json, specific_json, output_path_from_js, pattern_from_js)"
            )
        );

        workbookBytes = pyodide.FS.readFile(OUTPUT_PATH);

        renderLog(result.logs || []);
        renderStats(result);
        renderSheets(result.sheets || []);
        renderPreview(result.preview || [], result.columns || []);

        resultSummary.textContent =
            `${Number(result.processed_files).toLocaleString()} of ${Number(result.matched_files).toLocaleString()} matching files processed. ` +
            `Workbook size: ${sizeLabel(workbookBytes.length)}.`;

        resultBox.classList.remove("hidden");
        archiveStatus.textContent =
            "Workbook ready. Download it or save it to GitHub storage.";
        setProcessing(false, "Processing completed.");
    } catch (error) {
        showError(`Processing error: ${error.message}`);
        setProcessing(false, "Processing failed.");
        console.error(error);
    } finally {
        try { removeDirectoryTree(WORK_ROOT); } catch (_) {}
        try { pyodide.FS.unlink(OUTPUT_PATH); } catch (_) {}
        ["file_paths_json", "exclude_json", "specific_json", "output_path_from_js", "pattern_from_js"].forEach(
            (name) => {
                try { pyodide.globals.delete(name); } catch (_) {}
            }
        );
    }
}

function downloadResult() {
    if (!workbookBytes) return;

    const url = URL.createObjectURL(
        new Blob([workbookBytes], { type: XLSX_MIME })
    );

    const link = document.createElement("a");
    link.href = url;
    link.download = normaliseOutputName();
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

/* ==========================================================================
   Events
   ========================================================================== */

if ("webkitdirectory" in jsonFolder) {
    jsonFolder.webkitdirectory = true;
}

jsonFolder.addEventListener("change", (event) => {
    refreshSelection(event.target.files);
});

[excludeList, specificList, filePattern].forEach((input) => {
    input.addEventListener("input", updateMatchCount);
});

processBtn.addEventListener("click", processTaxonomy);
downloadBtn.addEventListener("click", downloadResult);
saveGithubBtn.addEventListener("click", saveCurrentToGithub);
refreshHistoryBtn.addEventListener("click", loadHistory);
outputName.addEventListener("blur", normaliseOutputName);

if (clearLogBtn && logOutput) {
    clearLogBtn.addEventListener("click", () => {
        logOutput.textContent = "Waiting for input files...";
    });
}

ghRemember.addEventListener("change", () => {
    try {
        if (ghRemember.checked && ghToken.value.trim()) {
            localStorage.setItem(TOKEN_KEY, ghToken.value.trim());
        } else {
            localStorage.removeItem(TOKEN_KEY);
        }
    } catch (_) {}
});

ghToken.addEventListener("input", () => {
    if (ghRemember.checked) {
        try { localStorage.setItem(TOKEN_KEY, ghToken.value.trim()); } catch (_) {}
    }
});

/* ==========================================================================
   Startup
   ========================================================================== */

updateProcessButton();
loadSavedToken();
loadHistory();
