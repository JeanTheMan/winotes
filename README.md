# winotes

An unofficial Python wrapper for **Windows Sticky Notes**, built by reverse-engineering the app's local SQLite database (`plum.sqlite`). No official API exists, this talks directly to the database and automatically restarts the app after every write so changes appear instantly.

---

## Requirements

- Windows 10 / 11 with the Sticky Notes app installed
- Python 3.10+
- No third-party packages — stdlib only (`sqlite3`, `subprocess`, `uuid`, etc.)

---

## Installation

```
pip install winotes
```

---

## Quick Start

```python
from winotes import Winotes

sn = Winotes()

# List all notes
for note in sn.list_notes():
    print(note["id"], note["theme"], note["text_plain"])

# Create a note
sn.create_note("Buy milk", theme="Blue")

# Update a note
sn.update_note(note_id, text="Buy oat milk", theme="Green")

# Delete a note (soft delete by default — recoverable)
sn.delete_note(note_id)

# Search notes
results = sn.search_notes("milk")
```

After any write, the app is automatically killed and relaunched so you see the changes immediately.

---

## Custom DB Path

By default the wrapper finds the database at:

```
%LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite
```

You can override this if needed:

```python
sn = Winotes(db_path=r"C:\Users\you\...\plum.sqlite")
```

---

## API Reference

### `Winotes(db_path=DEFAULT_DB_PATH)`

Creates the wrapper. Raises `FileNotFoundError` if the database can't be found.

---

### `list_notes(include_deleted=False) → list[dict]`

Returns all active notes. Pass `include_deleted=True` to also return soft-deleted notes.

Each note dict contains:

| Key | Type | Description |
|---|---|---|
| `id` | `str` | UUID of the note |
| `text_plain` | `str` | Content with internal block markers stripped |
| `text_raw` | `str` | Raw content as stored in the DB |
| `theme` | `str` | Colour theme (see themes below) |
| `is_open` | `bool` | Whether the note window is open |
| `is_always_on_top` | `bool` | Whether the note is pinned on top |
| `created_at` | `datetime` | Creation time (UTC) |
| `updated_at` | `datetime` | Last modified time (UTC) |
| `deleted_at` | `datetime \| None` | Deletion time, or `None` if active |
| `window_position` | `str \| None` | Raw window position string |
| `parent_id` | `str` | ID of the parent user record |

---

### `get_note(note_id) → dict | None`

Returns a single note by its ID, or `None` if not found.

---

### `create_note(text, theme="Yellow", is_open=True, is_always_on_top=False) → dict`

Creates a new note and returns it. Restarts the app.

```python
sn.create_note("hello world", theme="Pink")
sn.create_note("line one\nline two", theme="Blue", is_always_on_top=True)
```

---

### `update_note(note_id, text=None, theme=None, is_open=None, is_always_on_top=None) → dict | None`

Updates any combination of fields on an existing note. Only the arguments you pass are changed. Returns the updated note, or `None` if not found. Restarts the app.

```python
sn.update_note(note_id, text="updated content")
sn.update_note(note_id, theme="Charcoal", is_always_on_top=True)
```

---

### `delete_note(note_id, soft=True) → bool`

Deletes a note. Returns `True` if a note was affected. Restarts the app.

- `soft=True` (default) — sets `DeletedAt` timestamp; the note is hidden but recoverable via `restore_note()`.
- `soft=False` — permanently removes the row from the database.

---

### `restore_note(note_id) → dict | None`

Clears the `DeletedAt` timestamp on a soft-deleted note, making it active again.

---

### `search_notes(query, case_sensitive=False) → list[dict]`

Returns all active notes whose plain text contains `query`.

```python
results = sn.search_notes("TODO")
results = sn.search_notes("API", case_sensitive=True)
```

---

## Themes

The valid theme values are:

`Yellow` · `Blue` · `Green` · `Pink` · `Purple` · `Gray` · `Charcoal`

---

## CLI Demo

Running the script directly prints a summary of all current notes:

```
python -m winotes
```

```
============================================================
WINOTES — current notes
============================================================

  [Yellow  ]  8b154dd4-a795-4c3e-b2cf-6aa7010e0937
  Created : 2026-02-28 10:49
  Preview : Hello!
```

---

## How It Works

Windows Sticky Notes stores all data in a SQLite database (`plum.sqlite`) with no official access API. This wrapper:

1. **Reads** notes directly from the `Note` table via `sqlite3`.
2. **Parses** the internal text format — each paragraph is prefixed with a `\id=<uuid>` block marker which is stripped on read and regenerated on write.
3. **Converts** timestamps from Windows FILETIME (100-nanosecond ticks since 1601-01-01) to Python `datetime` objects.
4. **After writing**, kills `Microsoft.Notes.exe` and relaunches the app via `explorer.exe shell:AppsFolder\...` so it re-reads the database cleanly.

---

## Caveats

- This relies on undocumented internals and could break if Microsoft changes the database schema in a future update.
- If Sticky Notes is syncing with a Microsoft account, externally created/modified notes may be overwritten by the next sync cycle.
- Hard-deleting a note (`soft=False`) that exists on the server may cause sync conflicts.
