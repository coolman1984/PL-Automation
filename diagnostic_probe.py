"""Diagnostic #2: identify what Excel really sees inside the source file.

Read-only. Reports workbook identity, FileFormat, full sheet list including
visibility states, defined-name count, VBA presence, and open-time ReadOnly.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

# Windows' legacy console encoding cannot print the leading star in the real
# workbook filename.  Never let diagnostic output abort workbook inspection.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.errors import PLAutomationError
from src.excel_session import ExcelSession

VISIBILITY = {0: "visible", -1: "very-hidden", 1: "hidden"}


def main() -> int:
    source = Path(sys.argv[1])
    t0 = time.monotonic()

    def stamp(label: str) -> None:
        print(f"[{time.monotonic() - t0:7.2f}s] {label}", flush=True)

    session = ExcelSession.create(visible=False)
    try:
        stamp("opening read-only")
        wb = session.open_workbook(source, read_only=True, update_links=False)
        stamp("opened")
        print(f"FullName={wb.FullName}", flush=True)
        print(f"FileFormat={wb.FileFormat}", flush=True)
        print(f"ReadOnly={wb.ReadOnly}", flush=True)
        try:
            print(f"Names.Count={wb.Names.Count}", flush=True)
        except Exception as exc:
            print(f"Names.Count=<unavailable: {exc}>", flush=True)
        try:
            _ = bool(wb.VBProject)
            vb = True
        except Exception as exc:
            vb = f"<unavailable: {str(exc)[:100]}>"
        print(f"VBProject={vb}", flush=True)

        sheets = wb.Sheets
        total = int(sheets.Count)
        print(f"Sheets.Count={total}", flush=True)
        for i in range(1, total + 1):
            sh = sheets(i)
            vis = VISIBILITY.get(int(sh.Visible), str(sh.Visible))
            kind = "worksheet" if int(sh.Type) == -4167 else f"type:{sh.Type}"
            try:
                used = sh.UsedRange
                geometry = (
                    f"R{int(used.Row)}C{int(used.Column)} "
                    f"{int(used.Rows.Count)}x{int(used.Columns.Count)}"
                )
            except Exception as exc:
                geometry = f"<no-geometry: {str(exc)[:60]}>"
            print(f"  [{i}] {sh.Name!r} vis={vis} {kind} used={geometry}", flush=True)
        session.close_workbook(wb, save_changes=False)
    finally:
        session.close()
    stamp("done cleanly")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except PLAutomationError as exc:
        print(f"PL_ERROR [{exc.code}] {exc.message}")
        raise SystemExit(3)
