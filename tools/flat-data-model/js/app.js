const $ = (id) => document.getElementById(id);

const fileInput = $("fileInput");
const processBtn = $("processBtn");
const spinner = $("spinner");
const errorBox = $("errorBox");
const resultBox = $("resultBox");
const recordCount = $("recordCount");
const resultTable = $("resultTable");
const downloadBtn = $("downloadBtn");
const saveGithubBtn = $("saveGithubBtn");
const startNameInput = $("startName");
const prefixInput = $("fileNamePrefix");
const endNameInput = $("endName");
const fileNamePreview = $("fileNamePreview");
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
const tabRaw = $("tabRaw");
const tabExploded = $("tabExploded");

let pyodide = null;
let lastZipBytes = null;
let lastFileName = "output.zip";
let lastPreview = { raw: [], exploded: [] };
let activeTab = "raw";
let processing = false;

const TOKEN_KEY = "eba_tools_hub_github_token";
const POLL_INTERVAL = 3000;
const DEPLOY_TIMEOUT = 180000;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

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
  const result = { Accept: "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28" };
  if (authenticated && config.token) result.Authorization = `Bearer ${config.token}`;
  return result;
}

/* nome senza estensione: serve sia allo ZIP sia ai CSV interni */
function baseName() {
  return `${startNameInput.value || ""}${prefixInput.value || ""}${endNameInput.value || ""}`;
}

function outputName() {
  return `${baseName()}.zip`;
}

function setBusy(value, label = "Process file") {
  processing = value;
  spinner.hidden = !value;
  spinner.classList.toggle("hidden", !value);
  spinner.style.display = value ? "inline-block" : "none";
  processBtn.disabled = value;
  processBtn.textContent = value ? label : "Process file";
}

/* base64 -> Uint8Array (lo ZIP arriva da Python codificato) */
function fromBase64(text) {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/* Uint8Array -> base64, a blocchi per non saturare lo stack su file grandi */
function toBase64(bytes) {
  let binary = "";
  const chunk = 32768;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function publishedUrl(config, fileName) {
  return `${location.origin}/${config.repo}/${config.folder}/${encodeURIComponent(fileName)}`;
}

function sizeLabel(size) {
  if (size < 1024) return `${size} B`;
  if (size < 1048576) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / 1048576).toFixed(1)} MB`;
}

function renderResults(records) {
  resultTable.innerHTML = "";
  if (!records.length) return;
  const columns = Object.keys(records[0]);
  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  columns.forEach((column) => { const th = document.createElement("th"); th.textContent = column; headRow.append(th); });
  head.append(headRow);
  resultTable.append(head);
  const body = document.createElement("tbody");
  records.forEach((record) => {
    const row = document.createElement("tr");
    columns.forEach((column) => { const cell = document.createElement("td"); cell.textContent = record[column] ?? ""; row.append(cell); });
    body.append(row);
  });
  resultTable.append(body);
}

function showTab(name) {
  activeTab = name;
  tabRaw.classList.toggle("active", name === "raw");
  tabExploded.classList.toggle("active", name === "exploded");
  renderResults(name === "raw" ? lastPreview.raw : lastPreview.exploded);
}

const pythonReady = (async () => {
  pyodide = await loadPyodide({ indexURL: "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/" });
  await pyodide.loadPackage("micropip");
  const micropip = pyodide.pyimport("micropip");
  try { await micropip.install("openpyxl"); } finally { micropip.destroy(); }
  const response = await fetch("./python/flat_data_model_processor.py", { cache: "no-cache" });
  if (!response.ok) throw new Error(`Python engine not found: HTTP ${response.status}`);
  await pyodide.runPythonAsync(await response.text());
  processBtn.disabled = false;
})();

pythonReady.catch((error) => { console.error(error); processBtn.title = "Python engine unavailable"; });

async function getStoredFile(config, fileName) {
  const response = await fetch(`${apiBase(config)}/${encodeURIComponent(fileName)}?ref=${encodeURIComponent(config.branch)}&v=${Date.now()}`, { headers: apiHeaders(config, true), cache: "no-store" });
  if (response.status === 404) return null;
  if (!response.ok) throw new Error(`File lookup failed: HTTP ${response.status}`);
  return response.json();
}

/* lo ZIP e' binario: niente BOM, niente conversione da testo */
async function saveFile(config, fileName, bytes) {
  const previous = await getStoredFile(config, fileName);
  const body = { message: `${previous ? "Update" : "Add"} ZIP: ${fileName}`, content: toBase64(bytes), branch: config.branch };
  if (previous?.sha) body.sha = previous.sha;
  const response = await fetch(`${apiBase(config)}/${encodeURIComponent(fileName)}`, { method: "PUT", headers: { ...apiHeaders(config, true), "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (!response.ok) throw new Error(`Save failed: HTTP ${response.status} ${await response.text()}`);
}

async function deleteFile(config, item) {
  const response = await fetch(`${apiBase(config)}/${encodeURIComponent(item.name)}`, { method: "DELETE", headers: { ...apiHeaders(config, true), "Content-Type": "application/json" }, body: JSON.stringify({ message: `Delete file: ${item.name}`, sha: item.sha, branch: config.branch }) });
  if (!response.ok) throw new Error(`Delete failed: HTTP ${response.status} ${await response.text()}`);
}

/* confronto per dimensione: il testo non e' applicabile a un binario */
async function waitForDeployment(config, fileName, shouldExist, expectedSize = 0) {
  const started = Date.now();
  while (Date.now() - started < DEPLOY_TIMEOUT) {
    const response = await fetch(`${publishedUrl(config, fileName)}?v=${Date.now()}`, { cache: "no-store" });
    if (!shouldExist && response.status === 404) return;
    if (shouldExist && response.ok) {
      const buffer = await response.arrayBuffer();
      if (buffer.byteLength === expectedSize) return;
    }
    await sleep(POLL_INTERVAL);
  }
  throw new Error("GitHub Pages did not complete the deployment within 3 minutes.");
}

async function downloadStoredFile(item) {
  const response = await fetch(`${publishedUrl(getConfig(), item.name)}?v=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Download failed: HTTP ${response.status}`);
  const url = URL.createObjectURL(await response.blob());
  const link = document.createElement("a");
  link.href = url;
  link.download = item.name;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function createHistoryRow(item) {
  const row = document.createElement("tr");
  const nameCell = document.createElement("td");
  const sizeCell = document.createElement("td");
  const actionCell = document.createElement("td");
  const actions = document.createElement("div");
  const downloadButton = document.createElement("button");
  const deleteButton = document.createElement("button");
  nameCell.textContent = item.name;
  sizeCell.textContent = sizeLabel(item.size || 0);
  actions.className = "row-actions";
  downloadButton.textContent = "Download";
  downloadButton.className = "secondary";
  downloadButton.onclick = async () => {
    downloadButton.disabled = true;
    downloadButton.textContent = "Downloading...";
    try { await downloadStoredFile(item); } catch (error) { alert(error.message); }
    finally { downloadButton.disabled = false; downloadButton.textContent = "Download"; }
  };
  deleteButton.textContent = "Delete";
  deleteButton.className = "danger";
  deleteButton.onclick = async () => {
    const config = getConfig();
    if (!config.token) return alert("Enter the GitHub token first.");
    if (!confirm(`Delete ${item.name}?`)) return;
    downloadButton.disabled = true;
    deleteButton.disabled = true;
    deleteButton.textContent = "Deleting...";
    row.style.opacity = ".65";
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
    }
  };
  actions.append(downloadButton, deleteButton);
  actionCell.append(actions);
  row.append(nameCell, sizeCell, actionCell);
  return row;
}

async function loadHistory() {
  const body = historyTable.querySelector("tbody");
  body.innerHTML = "";
  historyError.classList.add("hidden");
  refreshHistoryBtn.disabled = true;
  refreshHistoryBtn.textContent = "Refreshing...";
  try {
    const config = getConfig();
    const response = await fetch(`${apiBase(config)}?ref=${encodeURIComponent(config.branch)}&v=${Date.now()}`, { headers: apiHeaders(config, Boolean(config.token)), cache: "no-store" });
    const items = response.status === 404 ? [] : await response.json();
    if (!response.ok && response.status !== 404) throw new Error(`History unavailable: HTTP ${response.status}`);
    items
      .filter((item) => item.type === "file" && /\.(zip|csv)$/i.test(item.name))
      .sort((a, b) => a.name.localeCompare(b.name))
      .forEach((item) => body.append(createHistoryRow(item)));
    if (!body.children.length) { const row = document.createElement("tr"); const cell = document.createElement("td"); cell.colSpan = 3; cell.textContent = "No files are currently stored."; row.append(cell); body.append(row); }
  } catch (error) {
    historyError.textContent = error.message;
    historyError.classList.remove("hidden");
  } finally {
    refreshHistoryBtn.disabled = false;
    refreshHistoryBtn.textContent = "Refresh list";
  }
}

processBtn.onclick = async () => {
  if (processing) return;
  errorBox.classList.add("hidden");
  resultBox.classList.add("hidden");
  const file = fileInput.files[0];
  if (!file) { errorBox.textContent = "Select an Excel file first."; errorBox.classList.remove("hidden"); return; }
  const path = "/tmp/input.xlsx";
  setBusy(true, pyodide ? "Processing..." : "Preparing...");
  try {
    await pythonReady;
    pyodide.FS.writeFile(path, new Uint8Array(await file.arrayBuffer()));
    pyodide.globals.set("workbook_path_from_js", path);
    pyodide.globals.set("base_name_from_js", baseName() || "mapping");
    const result = JSON.parse(await pyodide.runPythonAsync(
      "process_workbook(workbook_path_from_js, base_name_from_js)"
    ));
    if (!result.count) throw new Error("No valid records were found.");

    lastZipBytes = fromBase64(result.zip_base64);
    lastFileName = outputName();
    lastPreview = { raw: result.preview || [], exploded: result.preview_exploded || [] };

    showTab("raw");
    recordCount.textContent =
      `${result.files[0]}: ${result.count} rows — ${result.files[1]}: ${result.count_exploded} rows ` +
      `(preview limited to the first 200 rows of each file)`;
    resultBox.classList.remove("hidden");
    archiveStatus.textContent = `ZIP ready (${sizeLabel(lastZipBytes.length)}). Download it or save it to GitHub storage.`;
  } catch (error) {
    errorBox.textContent = `Processing error: ${error.message}`;
    errorBox.classList.remove("hidden");
  } finally {
    setBusy(false);
    try { pyodide.FS.unlink(path); } catch (_) {}
    try { pyodide.globals.delete("workbook_path_from_js"); } catch (_) {}
    try { pyodide.globals.delete("base_name_from_js"); } catch (_) {}
  }
};

saveGithubBtn.onclick = async () => {
  const config = getConfig();
  if (!lastZipBytes) return archiveStatus.textContent = "Generate a ZIP first.";
  if (!config.token) return archiveStatus.textContent = "Enter the GitHub token first.";
  saveGithubBtn.disabled = true;
  saveGithubBtn.textContent = "Saving...";
  try {
    await saveFile(config, lastFileName, lastZipBytes);
    saveGithubBtn.textContent = "Waiting for deployment...";
    archiveStatus.textContent = `Commit created. Waiting for ${lastFileName} to be published...`;
    await waitForDeployment(config, lastFileName, true, lastZipBytes.length);
    location.reload();
  } catch (error) {
    archiveStatus.textContent = error.message;
    saveGithubBtn.disabled = false;
    saveGithubBtn.textContent = "Save to GitHub storage";
  }
};

downloadBtn.onclick = () => {
  if (!lastZipBytes) return;
  const url = URL.createObjectURL(new Blob([lastZipBytes], { type: "application/zip" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = lastFileName;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

tabRaw.onclick = () => showTab("raw");
tabExploded.onclick = () => showTab("exploded");

[startNameInput, prefixInput, endNameInput].forEach((input) => { input.oninput = () => fileNamePreview.textContent = outputName(); });
refreshHistoryBtn.onclick = loadHistory;
ghRemember.onchange = () => ghRemember.checked ? localStorage.setItem(TOKEN_KEY, ghToken.value) : localStorage.removeItem(TOKEN_KEY);
ghToken.oninput = () => { if (ghRemember.checked) localStorage.setItem(TOKEN_KEY, ghToken.value); };
const savedToken = localStorage.getItem(TOKEN_KEY);
if (savedToken) { ghToken.value = savedToken; ghRemember.checked = true; }
setBusy(false);
fileNamePreview.textContent = outputName();
loadHistory();
