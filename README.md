# EBA Tools Hub

A single GitHub Pages repository for EBA, regulatory reporting and data management tools.

Link app https://federicomangini83-droid.github.io/EBA-Tools-Hub/

## Available tool

### Flat Data Model Creator

The tool converts a standard Excel workbook into a flat CSV data model directly in the browser. It also supports:

- local processing with Python and Pyodide;
- CSV preview and immediate download;
- GitHub-based CSV storage;
- replacement of an existing file when the same name is used;
- download and deletion of stored CSV files;
- automatic waiting for the GitHub Pages deployment before refreshing the page.

## Published URLs

After enabling GitHub Pages:

```text
Hub:
https://federicomangini83-droid.github.io/EBA-Tools-Hub/

Flat Data Model Creator:
https://federicomangini83-droid.github.io/EBA-Tools-Hub/tools/flat-data-model/
```

## Repository structure

```text
EBA-Tools-Hub/
├── index.html
├── README.md
├── .nojekyll
├── assets/
│   └── css/
│       └── hub.css
├── tools/
│   └── flat-data-model/
│       ├── index.html
│       ├── css/
│       │   └── style.css
│       ├── js/
│       │   └── app.js
│       └── python/
│           └── flat_data_model_processor.py
└── storage/
    └── flat-data-model/
        └── .gitkeep
```

## GitHub token

Saving and deleting files requires a fine-grained personal access token configured with:

```text
Repository: EBA-Tools-Hub
Repository permission: Contents - Read and write
```

The token is entered in the tool under **GitHub configuration**. If **Remember the token in this browser** is selected, the token is stored only in that browser's local storage.

## CSV API access

List the stored files through the GitHub Contents API:

```text
https://api.github.com/repos/federicomangini83-droid/EBA-Tools-Hub/contents/storage/flat-data-model
```

Retrieve a CSV directly:

```text
https://raw.githubusercontent.com/federicomangini83-droid/EBA-Tools-Hub/main/storage/flat-data-model/FILE_NAME.csv
```

## Adding another tool

Use the same structure for every new tool:

```text
tools/new-tool/
storage/new-tool/
```

Then add a new card to the root `index.html` catalogue.

## GitHub Pages publication

1. Upload the contents of this package to the root of the `EBA-Tools-Hub` repository.
2. Open `Settings > Pages`.
3. Select `Deploy from a branch`.
4. Select branch `main` and folder `/ (root)`.
5. Wait for the deployment to complete.
