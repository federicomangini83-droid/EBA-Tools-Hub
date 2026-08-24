# EBA Tools Hub

Repository unica per strumenti EBA, FinRep e Regulatory Data.

Link app https://federicomangini83-droid.github.io/EBA-Tools-Hub/

## Applicazione

Dopo la pubblicazione GitHub Pages:

- Hub: `https://federicomangini83-droid.github.io/EBAToolsHub/`
- Flat Data Model: `https://federicomangini83-droid.github.io/EBAToolsHub/tools/flat-data-model/`

## Struttura

```text
EBAToolsHub/
├── index.html
├── assets/
│   └── css/hub.css
├── tools/
│   └── flat-data-model/
│       ├── index.html
│       ├── css/style.css
│       ├── js/app.js
│       └── python/finrep_processor.py
└── storage/
    └── flat-data-model/
        └── .gitkeep
```

## Pubblicazione

1. Crea la repository pubblica `EBAToolsHub`.
2. Carica nella root tutto il contenuto dello ZIP.
3. Vai in `Settings > Pages`.
4. Seleziona `Deploy from a branch`, branch `main`, cartella `/ (root)`.
5. Attendi il deployment.

## Token GitHub

Il token usato nella vecchia repository non funzionerà automaticamente se era limitato solo a `FlatDataModel`.
Crea o modifica un fine-grained personal access token con:

- repository selezionata: `EBAToolsHub`;
- permesso repository `Contents: Read and write`.

Inseriscilo nel tool in `Configurazione GitHub`.

## Storico e accesso esterno

I CSV vengono archiviati in:

```text
storage/flat-data-model/
```

Elenco via GitHub API:

```text
https://api.github.com/repos/federicomangini83-droid/EBAToolsHub/contents/storage/flat-data-model
```

File raw:

```text
https://raw.githubusercontent.com/federicomangini83-droid/EBAToolsHub/main/storage/flat-data-model/NOME_FILE.csv
```

## Aggiunta di un nuovo tool

Per ogni nuovo strumento:

```text
tools/nome-tool/
storage/nome-tool/
```

Poi aggiungi una card nella homepage `index.html`.
