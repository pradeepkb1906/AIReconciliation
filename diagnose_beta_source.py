"""Diagnose the SyntaxWarning toast on Beta.

USAGE
-----
1. On Beta, open the AI Reconciliation tool source view, click anywhere
   in the source pane, press Cmd-A then Cmd-C.
2. Paste the clipboard content into a fresh file:
       pbpaste > /tmp/beta_recon.py
3. Run this diagnostic against it:
       python3 diagnose_beta_source.py /tmp/beta_recon.py

WHAT IT FINDS
-------------
The toast "<exec>:N: SyntaxWarning: invalid escape sequence '\\'" with an
invisible second character is caused by a backslash followed by a non-
printable control byte (0x01-0x1f or 0x7f). Common culprits are stray
ANSI escape sequences (ESC=0x1b) or BEL (0x07) that copy-paste pipelines
sometimes embed in clipboards.

This script:
  * runs the same compile() OWUI does
  * captures every SyntaxWarning with line + the exact byte
  * prints a hex dump of the surrounding 40 bytes so you can see what
    actually landed in Beta's DB
  * writes a scrubbed copy to /tmp/beta_recon_scrubbed.py for re-upload
"""
from __future__ import annotations

import io
import sys
import warnings
from pathlib import Path


def main() -> None:
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    src_path = Path(sys.argv[1])
    src_bytes = src_path.read_bytes()
    src = src_bytes.decode("utf-8", errors="replace")

    # 1. Capture every SyntaxWarning that compile() emits.
    captured: list[tuple[str, int, str, str]] = []

    def hook(message, category, filename, lineno, file=None, line=None):
        captured.append((str(filename), lineno, category.__name__, str(message)))

    warnings.showwarning = hook
    warnings.simplefilter("always")
    try:
        compile(src, "<exec>", "exec")
    except SyntaxError as exc:
        print(f"[!] SyntaxError on line {exc.lineno}: {exc.msg}")

    if not captured:
        print("[ok] compile() emitted no SyntaxWarning - the source is clean.")
    else:
        print(f"[!] compile() emitted {len(captured)} SyntaxWarning(s):")
        for filename, lineno, cat, msg in captured:
            print(f"    <{filename}>:{lineno}: {cat}: {msg!r}")

    # 2. Hunt for invisible control chars and backslash-followed-by-control.
    invisible = set(range(0x01, 0x09)) | {0x0B, 0x0C} | set(range(0x0E, 0x20)) | {0x7F}
    print()
    print("=== Invisible control characters anywhere in source ===")
    ctrl_hits = [(i, b) for i, b in enumerate(src_bytes) if b in invisible]
    if not ctrl_hits:
        print("    none")
    else:
        for i, b in ctrl_hits[:30]:
            line = src_bytes[:i].count(b"\n") + 1
            col = i - (src_bytes.rfind(b"\n", 0, i) + 1)
            ctx = src_bytes[max(0, i - 20) : i + 20]
            print(f"    L{line} col{col}: 0x{b:02x}  ctx={ctx!r}")

    print()
    print("=== Backslash followed by an invalid escape character ===")
    bs_hits = []
    for i in range(len(src_bytes) - 1):
        if src_bytes[i] == 0x5C:
            nxt = src_bytes[i + 1]
            valid = set(b"\\'\"abfnrtv01234567xNuU\n")
            if nxt not in valid:
                bs_hits.append((i, nxt))
    if not bs_hits:
        print("    none")
    else:
        for i, nxt in bs_hits[:30]:
            line = src_bytes[:i].count(b"\n") + 1
            col = i - (src_bytes.rfind(b"\n", 0, i) + 1)
            ctx = src_bytes[max(0, i - 15) : i + 5]
            disp = chr(nxt) if 0x20 <= nxt < 0x7F else f"\\x{nxt:02x}"
            print(f"    L{line} col{col}: \\{disp}  ctx={ctx!r}")

    # 3. Write a scrubbed copy.
    scrubbed = "".join(c for c in src if ord(c) not in invisible)
    out = Path("/tmp/beta_recon_scrubbed.py")
    out.write_text(scrubbed, encoding="utf-8")
    print()
    print(f"[ok] scrubbed copy written to {out} ({len(scrubbed)} chars)")
    print("     Re-upload this file to Beta's tool source pane to clear the toast.")


if __name__ == "__main__":
    main()
