# PR Office AI Project

PR / PO tracking system for an 8–10 person team, with separate frontend, backend, and database containers.

## Run with Docker

```bash
docker compose up -d --build
```

Open `http://localhost:8090`.

The deployment contains:

- `frontend`: Nginx serves the web interface and proxies `/api` requests.
- `backend`: Node.js provides the shared API and transaction-safe updates.
- `database`: PostgreSQL stores records for all users.
- `pr-office-database`: named Docker volume for persistent database storage.

Rebuilding or restarting either container does not delete the records.

## Configuration

For a shared or production deployment, set a strong database password before starting:

```bash
POSTGRES_PASSWORD=replace-with-a-strong-password docker compose up -d --build
```

The default password is intended only for local testing. The web port is `8090`; the database is not exposed outside Docker.

## Stored data

PostgreSQL stores:

- PR / PO tracking records
- Remarks, keyed by PO number when available
- Manually dismissed PO highlights

Each Excel import replaces the previous tracking records, clears old highlight dismissals so the 48-hour rule is recalculated, and keeps remarks for matching PO numbers.
