"""Seed / resync the AI Reconciliation Tool + Model into the running Open WebUI instance.

Run with Open WebUI's own Python so sqlite and Pydantic schemas match:

    /Users/pradeepbasavarajappa/.local/share/uv/tools/open-webui/bin/python seed_openwebui.py

Workflow (from CLAUDE.md mandatory sync recipe):
    1. Resolve the live OWUI PID and its DB via lsof (don't hardcode).
    2. Upsert the tool and model rows in a single transaction.
    3. SHA-verify DB content equals local file content.
    4. Sync local file mtimes to the same epoch as updated_at.
    5. (Caller restarts OWUI separately, then HTTP-probes localhost:8080.)
"""
from __future__ import annotations

import hashlib
import json
import os
import sqlite3
import subprocess
import sys
import time
import uuid
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL_PY = HERE / "reconciliation_tool.py"
MODEL_MD = HERE / "RECONCILIATION_MODEL.md"
README_MD = HERE / "README.md"

TOOL_ID = "ai_reconciliation_tool"
MODEL_ID = "ai-reconciliation-model"
BASE_MODEL_ID = "claude-opus-4-6"


def resolve_db_path() -> str:
    """Resolve the live OWUI SQLite DB via lsof on the running process."""
    try:
        pid = subprocess.check_output(
            "pgrep -f 'open-webui serve' | head -1", shell=True
        ).decode().strip()
        if not pid:
            sys.exit("Open WebUI is not running — start it first.")
        db = subprocess.check_output(
            f"lsof -p {pid} 2>/dev/null | grep webui.db | awk '{{print $NF}}' | sort -u | head -1",
            shell=True,
        ).decode().strip()
        if not db or not Path(db).exists():
            sys.exit(f"Could not resolve DB path from PID {pid}.")
        return db
    except subprocess.CalledProcessError as e:
        sys.exit(f"Failed to resolve OWUI DB path: {e}")


def get_admin_user_id(con: sqlite3.Connection) -> str:
    row = con.execute(
        "SELECT id FROM user WHERE role='admin' ORDER BY created_at LIMIT 1"
    ).fetchone()
    if not row:
        sys.exit("No admin user found.")
    return row[0]


def tool_specs() -> list:
    """Function-calling spec exposed to the LLM."""
    return [
        {
            "name": "reconcile",
            "description": (
                "Launch the AI Reconciliation wizard (inline iframe). Call with NO arguments when the "
                "user expresses a reconciliation intent. The wizard walks them through Scope & "
                "Jurisdiction, Upload Datasets (CSV / TSV / XLSX / XLS / DOCX / PPTX / PDF / JSON), "
                "Matching Rules, Run, and Review & Download (XLSX / CSV / PDF / DOCX) — all "
                "client-side with an ooXML server fallback. Advanced: only pass left_records and "
                "right_records if you have ALREADY parsed both datasets from chat context and need a "
                "static server-rendered report (CDN-blocked environments)."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "sector": {
                        "type": "string",
                        "description": "Optional sector hint to pre-select in the wizard.",
                        "enum": [
                            "Banking",
                            "Investment Banking",
                            "Insurance",
                            "Healthcare",
                            "Asset Management",
                            "Pharma Clinical",
                            "Energy",
                            "Telecommunications",
                            "Retail",
                            "Manufacturing",
                            "Public Sector",
                            "Technology",
                            "Transportation",
                            "Other",
                        ],
                    },
                    "region": {
                        "type": "string",
                        "description": "Optional region hint to pre-select in the wizard.",
                        "enum": ["USA", "European Union", "United Kingdom", "Global"],
                    },
                    "as_of": {"type": "string", "description": "Optional reporting date YYYY-MM-DD."},
                    "output_name": {"type": "string", "description": "Optional base filename for downloads."},
                    "left_records": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "ADVANCED: only if pre-parsed. Omit to launch the wizard.",
                    },
                    "right_records": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "ADVANCED: only if pre-parsed. Omit to launch the wizard.",
                    },
                    "key_fields": {"type": "array", "items": {"type": "string"}},
                    "amount_fields": {"type": "array", "items": {"type": "string"}},
                    "tolerance": {"type": "number"},
                    "regulations": {"type": "array", "items": {"type": "string"}},
                    "narrative": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "severity": {"type": "string", "enum": ["high", "medium", "low"]},
                                "regulation": {"type": "string"},
                                "text": {"type": "string"},
                            },
                        },
                    },
                },
                "required": [],
            },
        }
    ]


def tool_meta() -> dict:
    return {
        "description": "IBM Consulting Advantage — AI Reconciliation tool (Banking, Investment Banking, Insurance, Healthcare, Asset Management, Pharma Clinical; USA + Europe).",
        "manifest": {
            "title": "AI Reconciliation",
            "author": "IBM Consulting Advantage",
            "version": "1.0.0",
            "description": "Reconciles two datasets across regulated sectors and regions; inline report + XLSX/DOCX/PPTX downloads (CDN-first with ooXML server fallback).",
        },
    }


def model_meta() -> dict:
    return {
        "profile_image_url": "/static/favicon.png",
        "description": "IBM Consulting Advantage — AI Reconciliation. Stepwise inline wizard: Scope → Upload (CSV, TSV, XLSX, DOCX, PPTX, PDF, JSON) → Matching rules → Run → Download (XLSX, CSV, PDF, DOCX). USA, EU, UK, Global. Banking, Investment Banking, Insurance, Healthcare, Asset Management, Pharma, Energy, Telco, Retail, Manufacturing, Public Sector, Tech, Transportation. CDN-first with ooXML server fallback. Every output carries run ID, SHA-256 of inputs, and regulation tags for audit traceability.",
        "capabilities": {
            "vision": False,
            "usage": True,
            "citations": False,
            "web_search": False,
            "image_generation": False,
            "code_interpreter": False,
        },
        "suggestion_prompts": [
            {"content": "Start a Banking nostro reconciliation for USA (FFIEC / BCBS 239)."},
            {"content": "Launch an EMIR / MiFID II trade-vs-clearing reconciliation for the European Union."},
            {"content": "Begin an HIPAA 837-vs-835 claims reconciliation with PII redaction."},
            {"content": "Open the reconciliation wizard — I have custodian and administrator NAV extracts to compare."},
            {"content": "Start a reconciliation for manufacturing PO vs GRN, EU jurisdiction."},
        ],
        "tags": [
            {"name": "ibm-consulting-advantage"},
            {"name": "reconciliation"},
            {"name": "banking"},
            {"name": "insurance"},
            {"name": "healthcare"},
            {"name": "investment-banking"},
            {"name": "asset-management"},
            {"name": "pharma-clinical"},
            {"name": "usa"},
            {"name": "europe"},
        ],
        "toolIds": [TOOL_ID],
        "builtinTools": [],
    }


def model_params(system_prompt: str) -> dict:
    # Per CLAUDE.md: Bedrock Claude via litellm — pick ONE of temperature / top_p.
    return {
        "system": system_prompt,
        "temperature": 0.3,
        "function_calling": "native",
    }


def upsert(con: sqlite3.Connection, table: str, row: dict) -> None:
    cols = ",".join(row.keys())
    placeholders = ",".join("?" for _ in row)
    pk = row["id"]
    exists = con.execute(f"SELECT 1 FROM {table} WHERE id=?", (pk,)).fetchone()
    if exists:
        sets = ",".join(f"{k}=?" for k in row if k != "id")
        vals = [v for k, v in row.items() if k != "id"] + [pk]
        con.execute(f"UPDATE {table} SET {sets} WHERE id=?", vals)
        print(f"[update] {table}: {pk}")
    else:
        con.execute(f"INSERT INTO {table} ({cols}) VALUES ({placeholders})", list(row.values()))
        print(f"[insert] {table}: {pk}")


def main() -> None:
    if not TOOL_PY.exists() or not MODEL_MD.exists():
        sys.exit("Missing source files — expected reconciliation_tool.py and RECONCILIATION_MODEL.md next to this script.")

    db_path = resolve_db_path()
    print(f"[db] {db_path}")

    tool_src = TOOL_PY.read_text(encoding="utf-8")
    system_prompt = MODEL_MD.read_text(encoding="utf-8")
    now = int(time.time())

    con = sqlite3.connect(db_path, timeout=30)
    con.execute("PRAGMA busy_timeout=30000")
    try:
        admin_id = get_admin_user_id(con)
        print(f"[admin] user_id = {admin_id}")

        upsert(con, "tool", {
            "id": TOOL_ID,
            "user_id": admin_id,
            "name": "AI Reconciliation",
            "content": tool_src,
            "specs": json.dumps(tool_specs()),
            "meta": json.dumps(tool_meta()),
            "valves": None,
            "updated_at": now,
            "created_at": now,
        })

        upsert(con, "model", {
            "id": MODEL_ID,
            "user_id": admin_id,
            "base_model_id": BASE_MODEL_ID,
            "name": "IBM Consulting Advantage — AI Reconciliation",
            "meta": json.dumps(model_meta()),
            "params": json.dumps(model_params(system_prompt)),
            "is_active": 1,
            "updated_at": now,
            "created_at": now,
        })

        con.commit()

        # Hash-verify DB content equals local file content
        db_content = con.execute(
            "SELECT content FROM tool WHERE id=?", (TOOL_ID,)
        ).fetchone()[0]
        db_sha = hashlib.sha256(db_content.encode()).hexdigest()
        file_sha = hashlib.sha256(tool_src.encode()).hexdigest()
        if db_sha != file_sha:
            sys.exit(f"[error] SHA mismatch after upsert\n  db:   {db_sha}\n  file: {file_sha}")
        print(f"[sha] tool db==file ok ({db_sha[:16]}…)")

        db_prompt = json.loads(
            con.execute("SELECT params FROM model WHERE id=?", (MODEL_ID,)).fetchone()[0]
        )["system"]
        if db_prompt != system_prompt:
            sys.exit("[error] model.params.system mismatch")
        print(f"[sha] model system prompt verified ({hashlib.sha256(system_prompt.encode()).hexdigest()[:16]}…)")

        # Sync local file mtimes to the same epoch as updated_at
        for f in [TOOL_PY, MODEL_MD]:
            os.utime(f, (now, now))
        if README_MD.exists():
            os.utime(README_MD, (now, now))
        print(f"[mtime] synced epoch={now}")
    finally:
        con.close()

    print("\n[next] Restart Open WebUI to force tool re-import, then probe localhost:8080:")
    print("  for p in $(pgrep -f 'open-webui serve'); do kill -TERM $p; done; sleep 3; \\")
    print("    DYLD_FALLBACK_LIBRARY_PATH=/opt/homebrew/lib nohup /Users/pradeepbasavarajappa/.local/bin/open-webui serve > /tmp/owui-restart.log 2>&1 & disown; \\")
    print("    sleep 7; curl -s -o /dev/null -w 'HTTP %{http_code}\\n' http://localhost:8080/")


if __name__ == "__main__":
    main()
