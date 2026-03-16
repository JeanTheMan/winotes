import os
import re
import sqlite3
import subprocess
import time
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from typing import Optional

DEFAULT_DB_PATH = os.path.expandvars(
    r"%LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe"
    r"\LocalState\plum.sqlite"
)

THEMES = {"Yellow", "Blue", "Green", "Pink", "Purple", "Gray", "Charcoal"}

# Windows FILETIME epoch: 1601-01-01 00:00:00 UTC
# Python datetime epoch:  1970-01-01 00:00:00 UTC
_EPOCH_DIFF_100NS = 116_444_736_000_000_000  # 100-nanosecond intervals

def _now_filetime() -> int:
    """Return the current time as a Windows FILETIME (100-ns ticks)."""
    unix_ns = time.time_ns()
    return (unix_ns // 100) + _EPOCH_DIFF_100NS


_FILETIME_EPOCH = datetime(1601, 1, 1, tzinfo=timezone.utc)


def _filetime_to_datetime(ft: Optional[int]) -> Optional[datetime]:
    """Convert a Windows FILETIME integer to a UTC datetime (or None)."""
    if ft is None:
        return None
    from datetime import timedelta
    return _FILETIME_EPOCH + timedelta(microseconds=ft // 10)

# Text format helpers
# Sticky Notes stores text as a run of lines, each (optionally) prefixed with
# a block-id marker:  \id=<uuid> <rest of line>
# We parse these transparently and can reconstruct them on write.

_BLOCK_RE = re.compile(r"^\\id=([0-9a-f\-]{36}) ?(.*)", re.DOTALL)


def _parse_text(raw: Optional[str]) -> str:
    """Strip all \\id=<uuid> block markers and return plain text."""
    if not raw:
        return ""
    lines = raw.split("\n")
    plain = []
    for line in lines:
        m = _BLOCK_RE.match(line)
        plain.append(m.group(2) if m else line)
    return "\n".join(plain)


def _build_text(plain: str) -> str:
    """
    Convert plain multi-line text back to the Sticky Notes internal format,
    assigning a fresh UUID to each line/paragraph block.
    """
    lines = plain.split("\n")
    parts = []
    for line in lines:
        block_id = str(uuid.uuid4())
        parts.append(f"\\id={block_id} {line}")
    return "\n".join(parts)

class Winotes:

    def __init__(self, db_path: str = DEFAULT_DB_PATH):
        self.db_path = db_path
        if not os.path.exists(db_path):
            raise FileNotFoundError(
                f"plum.sqlite not found at:\n  {db_path}\n"
                "Check the path or pass db_path= explicitly."
            )

    @contextmanager
    def _connect(self, readonly: bool = False):
        uri = f"file:{self.db_path}"
        if readonly:
            uri += "?mode=ro"
        con = sqlite3.connect(uri, uri=True)
        con.row_factory = sqlite3.Row
        try:
            yield con
            if not readonly:
                con.commit()
        finally:
            con.close()

    @staticmethod
    def _restart_app() -> None:
        """Kill the running Sticky Notes process (if any), then relaunch it."""
        # Kill - ignore errors if it wasn't running
        subprocess.run(
            ["taskkill", "/f", "/im", "Microsoft.Notes.exe"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(0.5)   # brief pause so the DB file is fully released
        # Relaunch via the shell: protocol (works for UWP apps without a direct .exe path)
        subprocess.Popen(
            ["explorer.exe",
             "shell:AppsFolder\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

    def list_notes(self, include_deleted: bool = False) -> list[dict]:
        """
        Return all (non-deleted) notes as a list of dicts.

        Each dict contains:
          id, text_raw, text_plain, theme, is_open, is_always_on_top,
          created_at, updated_at, deleted_at, window_position, parent_id
        """
        with self._connect(readonly=True) as con:
            if include_deleted:
                rows = con.execute('SELECT * FROM "Note"').fetchall()
            else:
                rows = con.execute(
                    'SELECT * FROM "Note" WHERE "DeletedAt" IS NULL'
                ).fetchall()
        return [self._row_to_dict(r) for r in rows]

    def get_note(self, note_id: str) -> Optional[dict]:
        """Return a single note by its Id, or None if not found."""
        with self._connect(readonly=True) as con:
            row = con.execute(
                'SELECT * FROM "Note" WHERE "Id" = ?', (note_id,)
            ).fetchone()
        return self._row_to_dict(row) if row else None

    def create_note(
        self,
        text: str,
        theme: str = "Yellow",
        is_open: bool = True,
        is_always_on_top: bool = False,
        parent_id: Optional[str] = None,
    ) -> dict:
        """
        Insert a brand-new note and return it as a dict.

        Parameters
        ----------
        text            : Plain-text content (newlines supported).
        theme           : One of Yellow, Blue, Green, Pink, Purple, Gray, Charcoal.
        is_open         : Whether the note window is open.
        is_always_on_top: Pin note on top of other windows.
        parent_id       : ParentId (leave None to match the existing user record).
        """
        if theme not in THEMES:
            raise ValueError(f"theme must be one of {sorted(THEMES)}")

        note_id = str(uuid.uuid4())
        now = _now_filetime()
        raw_text = _build_text(text)

        # Try to borrow the ParentId from an existing note if not supplied
        if parent_id is None:
            existing = self.list_notes()
            parent_id = existing[0]["parent_id"] if existing else str(uuid.uuid4())

        with self._connect() as con:
            con.execute(
                """
                INSERT INTO "Note" (
                    "Id", "ParentId", "Text", "Theme",
                    "IsOpen", "IsAlwaysOnTop",
                    "IsFutureNote", "PendingInsightsScan",
                    "IsRemoteDataInvalid", "RemoteSchemaVersion",
                    "CreatedAt", "UpdatedAt", "DeletedAt", "Type"
                ) VALUES (?,?,?,?,?,?,0,0,0,0,?,?,NULL,NULL)
                """,
                (
                    note_id, parent_id, raw_text, theme,
                    int(is_open), int(is_always_on_top),
                    now, now,
                ),
            )
        result = self.get_note(note_id)
        self._restart_app()
        return result

    def update_note(
        self,
        note_id: str,
        text: Optional[str] = None,
        theme: Optional[str] = None,
        is_open: Optional[bool] = None,
        is_always_on_top: Optional[bool] = None,
    ) -> Optional[dict]:
        """
        Update one or more fields of an existing note.
        Only the keyword arguments you pass will be changed.

        Returns the updated note dict, or None if not found.
        """
        note = self.get_note(note_id)
        if note is None:
            return None

        if theme is not None and theme not in THEMES:
            raise ValueError(f"theme must be one of {sorted(THEMES)}")

        sets, params = [], []
        if text is not None:
            sets.append('"Text" = ?')
            params.append(_build_text(text))
        if theme is not None:
            sets.append('"Theme" = ?')
            params.append(theme)
        if is_open is not None:
            sets.append('"IsOpen" = ?')
            params.append(int(is_open))
        if is_always_on_top is not None:
            sets.append('"IsAlwaysOnTop" = ?')
            params.append(int(is_always_on_top))

        if not sets:
            return note  # nothing to do

        sets.append('"UpdatedAt" = ?')
        params.append(_now_filetime())
        params.append(note_id)

        with self._connect() as con:
            con.execute(
                f'UPDATE "Note" SET {", ".join(sets)} WHERE "Id" = ?',
                params,
            )
        result = self.get_note(note_id)
        self._restart_app()
        return result

    def delete_note(self, note_id: str, soft: bool = True) -> bool:
        """
        Delete a note.

        soft=True  (default) — sets DeletedAt timestamp (recoverable).
        soft=False           — permanently removes the row from the DB.

        Returns True if a row was affected, False if the note wasn't found.
        """
        with self._connect() as con:
            if soft:
                cur = con.execute(
                    'UPDATE "Note" SET "DeletedAt" = ?, "UpdatedAt" = ? WHERE "Id" = ?',
                    (_now_filetime(), _now_filetime(), note_id),
                )
            else:
                cur = con.execute(
                    'DELETE FROM "Note" WHERE "Id" = ?', (note_id,)
                )
        affected = cur.rowcount > 0
        if affected:
            self._restart_app()
        return affected

    def restore_note(self, note_id: str) -> Optional[dict]:
        """Un-delete a soft-deleted note (clears DeletedAt)."""
        with self._connect() as con:
            con.execute(
                'UPDATE "Note" SET "DeletedAt" = NULL, "UpdatedAt" = ? WHERE "Id" = ?',
                (_now_filetime(), note_id),
            )
        return self.get_note(note_id)

    def search_notes(self, query: str, case_sensitive: bool = False) -> list[dict]:
        """Return all non-deleted notes whose plain text contains *query*."""
        notes = self.list_notes()
        if not case_sensitive:
            query = query.lower()
        return [
            n for n in notes
            if query in (n["text_plain"] if case_sensitive else n["text_plain"].lower())
        ]

    # Internal helpers

    @staticmethod
    def _row_to_dict(row: sqlite3.Row) -> dict:
        d = dict(row)
        d["text_plain"] = _parse_text(d.get("Text"))
        d["text_raw"] = d.pop("Text", None)
        d["id"] = d.pop("Id")
        d["parent_id"] = d.pop("ParentId", None)
        d["theme"] = d.pop("Theme", None)
        d["is_open"] = bool(d.pop("IsOpen", 0))
        d["is_always_on_top"] = bool(d.pop("IsAlwaysOnTop", 0))
        d["created_at"] = _filetime_to_datetime(d.pop("CreatedAt", None))
        d["updated_at"] = _filetime_to_datetime(d.pop("UpdatedAt", None))
        d["deleted_at"] = _filetime_to_datetime(d.pop("DeletedAt", None))
        d["window_position"] = d.pop("WindowPosition", None)
        # remove less-useful fields from the top-level dict
        for k in ("Type", "RemoteId", "ChangeKey", "LastServerVersion",
                  "RemoteSchemaVersion", "IsRemoteDataInvalid",
                  "PendingInsightsScan", "IsFutureNote",
                  "CreationNoteIdAnchor"):
            d.pop(k, None)
        return d


if __name__ == "__main__":
    sn = Winotes()

    print("=" * 60)
    print("WINOTES — current notes")
    print("=" * 60)
    notes = sn.list_notes()
    if not notes:
        print("  (no notes found)")
    for n in notes:
        created = n["created_at"].strftime("%Y-%m-%d %H:%M") if n["created_at"] else "?"
        preview = n["text_plain"][:80].replace("\n", " ↵ ")
        print(f"\n  [{n['theme']:8}]  {n['id']}")
        print(f"  Created : {created}")
        print(f"  Preview : {preview}")
    print()