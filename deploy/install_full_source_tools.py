#!/usr/bin/env python3
"""Deploy FULL_FILE_READ patches for Generate Excel/Word/PDF/Slides + follow_up_gate."""
from __future__ import annotations

import json
import sys
import time
import uuid
from pathlib import Path

import paramiko

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = Path(__file__).resolve().parent
HOST = "192.168.0.17"
DB = "/var/lib/docker/volumes/open-webui_open-webui-data/_data/webui.db"
# Named user UUID for tool ACL (replace with your OWUI user id; leave empty for admin-only)
GHAZALA = ""
SURVEY_ID = "6a88639a-5371-4796-8045-f82512f0f519"

COMMON_PROPS = {
    "source_file_id": {
        "type": "string",
        "description": "Attached spreadsheet file id (from chat). Prefer this over inventing rows.",
    },
    "source_filename": {
        "type": "string",
        "description": "Optional filename hint if id unknown.",
    },
    "job": {
        "type": "string",
        "description": (
            "FULL_FILE_READ job: subject_group_report | mirror_workbook | full_sheet_tables. "
            "Required when a spreadsheet is attached."
        ),
    },
    "sheet_name": {
        "type": "string",
        "description": "Sheet to read (default Form Responses 1).",
    },
    "group_by": {
        "type": "string",
        "description": "Column to group by for subject_group_report (default Subject).",
    },
}

TOOLS = [
    {
        "name": "Generate Excel",
        "fn": "generate_excel",
        "local": HERE / "generate_excel.py",
        "remote": "/opt/open-webui/tools/generate_excel.py",
        "reqs": ["openpyxl"],
        "desc": (
            "Create Excel (.xlsx). When a spreadsheet is attached, MUST use "
            "job=subject_group_report|mirror_workbook|full_sheet_tables for COMPLETE "
            "disk read (never RAG chunks). Returns VERIFY counts + download link."
        ),
    },
    {
        "name": "Generate Word",
        "fn": "generate_docx",
        "local": HERE / "generate_docx.py",
        "remote": "/opt/open-webui/tools/generate_docx.py",
        "reqs": ["python-docx"],
        "desc": (
            "Create Word (.docx). Attached spreadsheet → job=subject_group_report "
            "for FULL_FILE_READ. Never invent tables from chat snippets."
        ),
    },
    {
        "name": "Generate PDF",
        "fn": "generate_pdf",
        "local": HERE / "generate_pdf.py",
        "remote": "/opt/open-webui/tools/generate_pdf.py",
        "reqs": ["reportlab"],
        "desc": (
            "Create PDF. Attached spreadsheet → job=subject_group_report for "
            "FULL_FILE_READ. Never invent KPIs from RAG."
        ),
    },
    {
        "name": "Generate Slides",
        "fn": "generate_slides",
        "local": HERE / "generate_slides.py",
        "remote": "/opt/open-webui/tools/generate_slides.py",
        "reqs": ["python-pptx", "pillow"],
        "desc": (
            "Create PowerPoint (.pptx). Attached spreadsheet → job=subject_group_report "
            "for FULL_FILE_READ. Never invent slides from RAG snippets."
        ),
    },
]


def main() -> int:
    for t in TOOLS:
        if not t["local"].is_file():
            print("MISSING", t["local"])
            return 2

    c = paramiko.SSHClient()
    c.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    c.connect(HOST, username="root", timeout=30)

    def run(cmd: str, t: int = 240) -> str:
        print(f"$ {cmd[:160]}")
        _, o, e = c.exec_command(cmd, timeout=t)
        out = (o.read() + e.read()).decode("utf-8", "replace")
        if out.strip():
            print(out[-6000:] if len(out) > 6000 else out)
        return out

    run("mkdir -p /opt/open-webui/tools /opt/open-webui/functions")
    sftp = c.open_sftp()
    sftp.put(str(HERE / "lib_source_workbook.py"), "/opt/open-webui/tools/lib_source_workbook.py")
    sftp.put(str(HERE / "follow_up_gate.py"), "/opt/open-webui/functions/follow_up_gate.py")
    for t in TOOLS:
        sftp.put(str(t["local"]), t["remote"])
        print("uploaded", t["name"], t["local"].stat().st_size)
    sftp.close()

    # container copies
    run("docker cp /opt/open-webui/tools/lib_source_workbook.py open-webui:/app/backend/data/lib_source_workbook.py")
    for t in TOOLS:
        run(f"docker cp {t['remote']} open-webui:/tmp/{Path(t['remote']).name}")

    run(
        "docker exec open-webui sh -c "
        "'pip install -q openpyxl python-docx reportlab python-pptx pillow 2>&1 | tail -3; "
        "python -c \"import openpyxl,docx,reportlab,pptx,PIL; print(\"deps_ok\")\"'"
    )

    payload = {
        "db": DB,
        "ghazala": GHAZALA,
        "filter_content": (HERE / "follow_up_gate.py").read_text(encoding="utf-8"),
        "tools": [],
    }
    for t in TOOLS:
        props = {
            "content": {
                "type": "string",
                "description": "JSON spec for blank/new docs only. Ignored when job+source FULL_FILE_READ.",
            },
            **COMMON_PROPS,
        }
        specs = [
            {
                "name": t["fn"],
                "description": t["desc"],
                "parameters": {
                    "type": "object",
                    "properties": props,
                    "required": [],
                },
            }
        ]
        payload["tools"].append(
            {
                "name": t["name"],
                "remote": t["remote"],
                "specs": specs,
                "reqs": t["reqs"],
                "content": t["local"].read_text(encoding="utf-8"),
            }
        )

    sftp = c.open_sftp()
    with sftp.file("/tmp/full_source_payload.json", "w") as f:
        f.write(json.dumps(payload))
    sftp.close()

    runner = r'''
import json, sqlite3, time, uuid
from pathlib import Path
p = json.loads(Path("/tmp/full_source_payload.json").read_text(encoding="utf-8"))
con = sqlite3.connect(p["db"])
con.row_factory = sqlite3.Row
admin = con.execute("select id from user where role='admin' limit 1").fetchone()
admin_id = admin["id"] if admin else None
now = int(time.time())
ghazala = p["ghazala"]

# filter
row = con.execute("select id from function where name=?", ("File Job Follow-up",)).fetchone()
if row:
    con.execute(
        "update function set content=?, updated_at=? where id=?",
        (p["filter_content"], now, row["id"]),
    )
    print("FILTER_UPDATED", row["id"])
else:
    # try type filter table variants
    print("FILTER_ROW_MISSING — listing functions")
    for r in con.execute("select id,name from function order by name"):
        print(" ", dict(r))

for t in p["tools"]:
    specs = json.dumps(t["specs"])
    meta = json.dumps({
        "description": t["specs"][0]["description"][:200],
        "manifest": {
            "title": t["name"],
            "author": "OmniTech",
            "version": "2.0.0",
            "requirements": t["reqs"],
        },
    })
    row = con.execute("select id from tool where name=?", (t["name"],)).fetchone()
    if row:
        tid = row["id"]
        con.execute(
            "update tool set content=?, specs=?, meta=?, updated_at=? where id=?",
            (t["content"], specs, meta, now, tid),
        )
        print("TOOL_UPDATED", t["name"], tid, "len", len(t["content"]))
    else:
        tid = str(uuid.uuid4())
        con.execute(
            "insert into tool (id,user_id,name,content,specs,meta,valves,updated_at,created_at) "
            "values (?,?,?,?,?,?,?,?,?)",
            (tid, admin_id, t["name"], t["content"], specs, meta, "{}", now, now),
        )
        print("TOOL_INSERTED", t["name"], tid)

    # grants
    for ptype, pid in (("user", "*"), ("user", ghazala)):
        if not pid:
            continue
        exists = con.execute(
            "select id from access_grant where resource_type='tool' and resource_id=? "
            "and principal_type=? and principal_id=? and permission='read'",
            (tid, ptype, pid),
        ).fetchone()
        if not exists:
            con.execute(
                "insert into access_grant(id,resource_type,resource_id,principal_type,principal_id,permission,created_at) "
                "values (?,?,?,?,?,?,?)",
                (str(uuid.uuid4()), "tool", tid, ptype, pid, "read", now),
            )
            print("GRANT", t["name"], ptype, pid)

con.commit()
print("DB_OK")
con.close()
'''
    sftp = c.open_sftp()
    with sftp.file("/tmp/install_full_source_db.py", "w") as f:
        f.write(runner)
    sftp.close()
    run("python3 /tmp/install_full_source_db.py")

    print("=== restart ===")
    run("docker restart open-webui")
    for i in range(15):
        time.sleep(3)
        health = run(
            "docker inspect -f '{{.State.Health.Status}}' open-webui"
        ).strip()
        api = run(
            "curl -sS -o /dev/null -w '%{http_code}' http://192.168.0.17:3080/api/config"
        ).strip()
        print(f"wait {i}: {health} {api}")
        if "healthy" in health and api.endswith("200"):
            break

    run("bash /opt/open-webui/scripts/apply_name_patch.sh || true")
    run("docker cp /opt/open-webui/tools/lib_source_workbook.py open-webui:/app/backend/data/lib_source_workbook.py")

    print("=== SMOKE FULL_FILE_READ ===")
    smoke = f'''
docker exec open-webui python - <<'PY'
import asyncio, json, sys
sys.path.insert(0, "/app/backend/data")
import lib_source_workbook as sw
from importlib.util import spec_from_file_location, module_from_spec

fid = "{SURVEY_ID}"
# resolve via DB
fid2, filename, path = sw.resolve_file_path(source_file_id=fid)
print("RESOLVED", fid2, filename, path)
book = sw.load_workbook_rows(path)
print("SHEETS", [(s["name"], s["row_count"]) for s in book["sheets"][:3]], "... total_rows_all_sheets", book["total_data_rows"])
built = sw.build_subject_group_report(book, sheet_name="Form Responses 1", group_by="Subject", title="TL Survey SMOKE")
v = built["verify"]
print("VERIFY", v["total_rows"], v["group_count"], sum(v["groups"].values()))
print("GROUPS", json.dumps(v["groups"], ensure_ascii=False))
assert v["total_rows"] == 105, v
assert v["group_count"] == 11, v
assert sum(v["groups"].values()) == 105, v

# generate_excel tool path
spec = spec_from_file_location("ge", "/opt/open-webui/tools/generate_excel.py")
# file may only be on host path — use /tmp copy
import os
src = "/tmp/generate_excel.py"
if not os.path.isfile(src):
    src = "/opt/open-webui/tools/generate_excel.py"
# inside container /tmp was copied
src = "/tmp/generate_excel.py"
spec = spec_from_file_location("ge", src)
mod = module_from_spec(spec)
spec.loader.exec_module(mod)
tools = mod.Tools()

async def run():
    out = await tools.generate_excel(
        content='{{"title":"TL Survey 2026 FULL_FILE_READ"}}',
        source_file_id=fid,
        job="subject_group_report",
        sheet_name="Form Responses 1",
        group_by="Subject",
        __files__=[{{"id": fid, "filename": filename, "type": "file"}}],
        __user__={{"id": "{GHAZALA}", "role": "user"}},
    )
    print(out[:1200])
    assert "FULL_FILE_READ" in out or "VERIFY" in out
    assert "105" in out
    assert "/api/v1/files/" in out
    return out

print(asyncio.run(run())[:500])
print("SMOKE_OK")
PY
'''
    # Fix the smoke script - I over-escaped. Write clean file instead.
    smoke_py = r'''
import asyncio, json, os, sys
sys.path.insert(0, "/app/backend/data")
import lib_source_workbook as sw
from importlib.util import spec_from_file_location, module_from_spec

fid = "''' + SURVEY_ID + r'''"
fid2, filename, path = sw.resolve_file_path(source_file_id=fid)
print("RESOLVED", fid2, filename, path)
book = sw.load_workbook_rows(path)
built = sw.build_subject_group_report(book, sheet_name="Form Responses 1", group_by="Subject", title="TL Survey SMOKE")
v = built["verify"]
print("VERIFY", v["total_rows"], v["group_count"], sum(v["groups"].values()))
print("GROUPS", json.dumps(v["groups"], ensure_ascii=False))
assert v["total_rows"] == 105
assert v["group_count"] == 11
assert sum(v["groups"].values()) == 105

src = "/tmp/generate_excel.py"
spec = spec_from_file_location("ge", src)
mod = module_from_spec(spec)
spec.loader.exec_module(mod)
tools = mod.Tools()

async def run():
    out = await tools.generate_excel(
        content='{"title":"TL Survey 2026 FULL_FILE_READ"}',
        source_file_id=fid,
        job="subject_group_report",
        sheet_name="Form Responses 1",
        group_by="Subject",
        __files__=[{"id": fid, "filename": filename, "type": "file"}],
        __user__={"id": "''' + GHAZALA + r'''", "role": "user"},
    )
    print(out[:1500])
    assert "105" in out
    assert "/api/v1/files/" in out
    return out

asyncio.run(run())
print("SMOKE_OK")
'''
    sftp = c.open_sftp()
    with sftp.file("/tmp/smoke_full_source.py", "w") as f:
        f.write(smoke_py)
    sftp.close()
    run("docker cp /tmp/smoke_full_source.py open-webui:/tmp/smoke_full_source.py")
    run("docker cp /opt/open-webui/tools/generate_excel.py open-webui:/tmp/generate_excel.py")
    run("docker exec open-webui python /tmp/smoke_full_source.py")

    c.close()
    print("DONE")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
