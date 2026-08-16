# Cash Tray Manager — Apps Script API Contract

This document defines the interface between `index.html` (client) and the Google
Apps Script backend (`Code.gs`). Keep this in sync whenever either side changes.

## Security model

- **Important:** the deployed Apps Script in the wild accepts the **legacy** contract
  below (token in URL query string for GET, JSON body for POST). The current client uses
  exactly that contract so it works with both the deployed backend and the new `Code.gs`.
- The token is embedded in client-side JS, so it is **visible to any user**. Treat it
  as a shared secret that limits casual tampering only. For real security, redeploy the
  Apps Script with **Execute as: Me / Who has access: only your domain** and check
  `Session.getEffectiveUser()` inside the script. This removes the need for a client token.
- Backend must verify the token on every request and return `{status:'error', message}` on
  mismatch or malformed input.

## Transport

- **Read / delete:** `GET <GAS_URL>?action=...&token=...&...` (query params).
- **Create / update:** `POST <GAS_URL>` with `Content-Type: text/plain` (avoiding CORS
  preflight) and a JSON body that includes `_token`.
- A shared `SECRET_TOKEN` must match `EXPECTED_TOKEN` in `Code.gs`.

> Do **not** send `action:"getAll"` via POST to the legacy backend: it does not understand
> it and appends an empty junk row. The new `Code.gs` handles `getAll`/`delete` in both GET
> and POST, so upgrading is safe.

## Endpoints

### 1. `getAll` — GET

```
GET <GAS_URL>?action=getAll&token=<TOKEN>&sheetId=<optional-override>
```

Response:

```json
{
  "status": "ok",
  "data": [
    {
      "id": "ID_...",
      "date": "2026-08-15",
      "bankIndentNo": "IND-001",
      "bankIndentVal": 100000,
      "atmIds": ["ATM001", "ATM002"],
      "totalNotesAll": 500,
      "totalValueAll": 200000,
      "totalLoadedVal": 100000,
      "trays": {
        "t1": { "opening": 0, "added": 100, "totalNotes": 100, "totalValue": 20000 },
        "t2": { "opening": 0, "added": 100, "totalNotes": 100, "totalValue": 10000 },
        "t3": { "opening": 0, "added": 100, "totalNotes": 100, "totalValue": 50000 }
      },
      "summary": { "totalNotesAll": 500, "totalValueAll": 200000, "totalLoadedVal": 100000 }
    }
  ]
}
```

### 2. `save` (create)

```json
{
  "action": "save", "_token": "...", "id": "ID_<timestamp>",
  "date": "2026-08-15", "bankIndentNo": "IND-001", "bankIndentVal": 100000,
  "atmIds": ["ATM001"],
  "totalNotesAll": 500, "totalValueAll": 200000, "totalLoadedVal": 100000,
  "trays": { "t1": {...}, "t2": {...}, "t3": {...} },
  "summary": { "totalNotesAll": 500, "totalValueAll": 200000, "totalLoadedVal": 100000 }
}
```

Response: `{ "status": "ok" }`.

### 3. `update`

Same body as `save`, with `action: "update"` and the **existing** `id`. Replaces the row
with that id (or upserts if not found).

### 4. `delete`

- Legacy (deployed backend): `GET <GAS_URL>?action=delete&token=<TOKEN>&id=<ID>`
- New `Code.gs`: also accepts `POST` with `{ "action": "delete", "_token": "...", "id": "ID_..." }`

Response: `{ "status": "ok" }`. Hard-deletes the row.

## Spreadsheet layout (sheet tab: `Data`)

| Col | Field            | Type         |
|----:|------------------|--------------|
| A   | id               | text         |
| B   | date             | text (YYYY-MM-DD) |
| C   | bankIndentNo     | text         |
| D   | bankIndentVal    | number       |
| E   | atmIds           | JSON array   |
| F   | totalNotesAll    | number       |
| G   | totalValueAll    | number       |
| H   | totalLoadedVal   | number       |
| I   | trays            | JSON object  |
| J   | summary          | JSON object  |
| K   | updatedAt        | timestamp    |

Trays and summary are stored as JSON strings for flexibility; parse on read.

## Deploying

1. Open the script bound to your spreadsheet (Extensions → Apps Script).
2. Paste `Code.gs`.
3. Set `EXPECTED_TOKEN` and `CONFIG_SHEET_ID` at the top.
4. Deploy → New deployment → **Web app**.
   - Execute as: **Me**
   - Who has access: as appropriate for your users
5. Copy the `/exec` URL into `GAS_URL` in `index.html`.
6. Re-deploy the web app whenever you change `Code.gs` (you get a new URL each time;
   keep the `exec` URL stable by choosing "Update" on an existing deployment).
