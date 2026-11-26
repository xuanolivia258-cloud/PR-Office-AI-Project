# PR Office AI Project

PR / PO tracking system with a persistent Node.js backend.

## Run with Docker

```bash
docker compose up -d --build
```

Open `http://localhost:8090`.

Application data is stored in the named Docker volume `pr-office-data`. Rebuilding or restarting the container does not delete the records.

## Run locally

Node.js 20 or newer is required.

```bash
npm start
```

The local data file is created at `data/store.json` and is excluded from Git.

## Stored data

The backend stores:

- PR / PO tracking records
- Remarks, keyed by PO number when available
- Manually dismissed PO highlights

Each Excel import replaces the previous tracking records, clears old highlight dismissals so the 48-hour rule is recalculated, and keeps remarks for matching PO numbers.
