const $ = id => document.getElementById(id);

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

let pyodide = null;
let lastCSV = "";
let lastFileName = "output.csv";
let processing = false;

const TOKEN_KEY = "eba_tools_hub_github_token";
const POLL = 3000;
const TIMEOUT = 180000;

const sleep = ms => new Promise(r => setTimeout(r, ms));

function cfg() {
    return {
        owner: ghOwner.value.trim(),
        repo: ghRepo.value.trim(),
        branch: ghBranch.value.trim() || "main",
        folder: ghFolder.value
            .trim()
            .replace(/^\/+|\/+$/g, ""),
        token: ghToken.value.trim()
    };
}

function base(c) {
    return `https://api.github.com/repos/${c.owner}/${c.repo}/contents/${c.folder}`;
}

function headers(c, auth = false) {
    const h = {
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    };

    if (auth && c.token) {
        h.Authorization = `Bearer ${c.token}`;
    }

    return h;
}

function name() {
    return (
        `${startNameInput.value || ""}` +
        `${prefixInput.value || ""}` +
        `${endNameInput.value || ""}.csv`
    );
}

function busy(v, t = "Process file") {
    processing = v;

    spinner.hidden = !v;
    spinner.classList.toggle("hidden", !v);
    spinner.style.display = v ? "inline-block" : "none";

    processBtn.disabled = v;
    processBtn.textContent = v ? t : "Process file";
}

function b64(s) {
    const a = new TextEncoder().encode(s);

    let x = "";

    a.forEach(b => {
        x += String.fromCharCode(b);
    });

    return btoa(x);
}

function pageUrl(c, n) {
    return (
        `${location.origin}/` +
        `${c.repo}/` +
        `${c.folder}/` +
        `${encodeURIComponent(n)}`
    );
}

function bytes(n) {
    if (n < 1024) {
        return `${n} B`;
    }

    if (n < 1048576) {
        return `${(n / 1024).toFixed(1)} KB`;
    }

    return `${(n / 1048576).toFixed(1)} MB`;
}

function render(records) {
    resultTable.innerHTML = "";

    const hs = Object.keys(records[0] || {});

    const head = document.createElement("thead");
    const hr = document.createElement("tr");

    hs.forEach(h => {
        const th = document.createElement("th");
        th.textContent = h;
        hr.append(th);
    });

    head.append(hr);
    resultTable.append(head);

    const body = document.createElement("tbody");

    records.slice(0, 200).forEach(r => {
        const tr = document.createElement("tr");

        hs.forEach(h => {
            const td = document.createElement("td");
            td.textContent = r[h] ?? "";
            tr.append(td);
        });

        body.append(tr);
    });

    resultTable.append(body);
}
