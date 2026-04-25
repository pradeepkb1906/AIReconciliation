"""Progressive-chunk bisection for the SyntaxWarning toast.

Strategy
--------
Compile the source one growing chunk at a time:
  10%  ->  did SyntaxWarning fire?
  20%  ->  did SyntaxWarning fire?
  ...
  100% ->  did SyntaxWarning fire?

The first chunk that fires the warning isolates the offending block.
Then bisect within that chunk (50% / 75% / 87.5% ...) to pinpoint the
exact line. This isolates where in the source a stray bad escape lives,
without needing to upload anything.

USAGE
-----
    python3 bisect_syntax_warning.py reconciliation_tool.py
"""
from __future__ import annotations

import io
import sys
import warnings
from pathlib import Path


def compile_with_warnings(src: str) -> list[tuple[int, str, str]]:
    """Return list of (lineno, category, message) emitted during compile."""
    captured: list[tuple[int, str, str]] = []

    def hook(message, category, filename, lineno, file=None, line=None):
        captured.append((lineno, category.__name__, str(message)))

    prev = warnings.showwarning
    prev_filters = warnings.filters[:]
    warnings.showwarning = hook
    warnings.simplefilter("always")
    try:
        compile(src, "<exec>", "exec")
    except SyntaxError as exc:
        captured.append((exc.lineno or 0, "SyntaxError", exc.msg))
    finally:
        warnings.showwarning = prev
        warnings.filters = prev_filters
    return captured


def safe_truncate_at_line(src: str, line_count: int) -> str:
    """Cut the source at line_count.

    If we land mid-string-literal we tack on the closing triple-quote so the
    chunk is a self-contained Python source. We also count unbalanced ('
    and ") and append matching closers as a best-effort.
    """
    lines = src.splitlines(keepends=True)
    chunk = "".join(lines[:line_count])

    # Track whether we're inside a triple-quoted string at the cut point
    # using a tiny tokenizer-like scan that respects single/double, raw, and
    # triple variants.
    triple_dq = chunk.count('"""') - chunk.count('\\"""')
    triple_sq = chunk.count("'''") - chunk.count("\\'''")
    if triple_dq % 2 == 1:
        chunk += '\n"""\n'
    if triple_sq % 2 == 1:
        chunk += "\n'''\n"
    return chunk


def is_syntax_warning(category: str, message: str) -> bool:
    """Filter for invalid-escape SyntaxWarning only — not SyntaxError from cut."""
    return category == "SyntaxWarning" and "invalid escape" in message.lower()


def main() -> None:
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    src_path = Path(sys.argv[1])
    src = src_path.read_text(encoding="utf-8")
    lines = src.splitlines(keepends=True)
    total = len(lines)
    print(f"Source: {src_path}")
    print(f"Total lines: {total}")
    print()

    # Whole-file check first
    full_warns = compile_with_warnings(src)
    full_syn = [w for w in full_warns if is_syntax_warning(w[1], w[2])]
    if not full_syn:
        print("[ok] Whole file emits no invalid-escape SyntaxWarning.")
        print("    The source you provided is clean. If Beta still toasts the")
        print("    warning, it has different bytes than this file, OR the toast")
        print("    is sticky from a prior process. Restart the OWUI worker.")
        return
    print(f"[!] Whole file emits {len(full_syn)} invalid-escape SyntaxWarning(s):")
    for ln, cat, msg in full_syn:
        print(f"    L{ln}: {msg}")
    print()

    # Phase 1: 10% increments — only count invalid-escape warnings
    print("=== Phase 1: 10% growing chunks (counting invalid-escape only) ===")
    print(f"{'pct':>5}  {'lines':>6}  {'esc-warn':>8}  detail")
    first_failing_pct = None
    for pct in range(10, 110, 10):
        line_count = max(1, (total * pct) // 100)
        chunk = safe_truncate_at_line(src, line_count)
        warns = compile_with_warnings(chunk)
        syn = [w for w in warns if is_syntax_warning(w[1], w[2])]
        flag = len(syn)
        first_msg = f"L{syn[0][0]}: {syn[0][2][:80]}" if syn else "clean"
        marker = " <-- FIRST FAILURE" if (flag and first_failing_pct is None) else ""
        if flag and first_failing_pct is None:
            first_failing_pct = pct
        print(f"{pct:>4}%  {line_count:>6}  {flag:>8}  {first_msg}{marker}")

    if first_failing_pct is None:
        print()
        print("[odd] Whole file fired but no chunk did - likely a cut-string artifact.")
        return

    # Phase 2: bisect inside the failing chunk
    print()
    print(f"=== Phase 2: bisect within first {first_failing_pct}% ===")
    prev_pct = first_failing_pct - 10
    lo = max(1, (total * prev_pct) // 100)
    hi = max(1, (total * first_failing_pct) // 100)
    while hi - lo > 1:
        mid = (lo + hi) // 2
        chunk = safe_truncate_at_line(src, mid)
        warns = compile_with_warnings(chunk)
        syn = [w for w in warns if is_syntax_warning(w[1], w[2])]
        if syn:
            hi = mid
            label = f"L{syn[0][0]}: {syn[0][2][:60]}"
            print(f"  cut at line {mid:>5}: WARN  {label}")
        else:
            lo = mid
            print(f"  cut at line {mid:>5}: clean")

    # Show the offending line range
    print()
    print(f"=== Offending line is between {lo+1} and {hi} ===")
    for ln in range(max(1, lo - 2), min(total, hi + 3) + 1):
        marker = " <-- offending" if ln in (lo + 1, hi) else ""
        print(f"  L{ln:>5}: {lines[ln-1].rstrip()[:100]}{marker}")


if __name__ == "__main__":
    main()
