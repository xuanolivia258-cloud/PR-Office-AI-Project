# PR Office AI Project

PR / PO tracking system with separate frontend and backend containers.

## Run with Docker

```bash
docker compose up -d --build
```

Open `http://localhost:8090`.

The deployment contains:

- `frontend`: Nginx serves the web interface and proxies `/api` requests.
- `backend`: Node.js provides the persistent data API.
- `pr-office-data`: named Docker volume for stored records.

Rebuilding or restarting either container does not delete the records.

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
