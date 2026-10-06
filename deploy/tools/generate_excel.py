"""
title: Generate Excel
author: OmniTech
description: Generate native Excel (.xlsx). Prefer FULL attached-file read via source_file_id/job — never RAG chunks.
requirements: openpyxl
required_open_webui_version: 0.4.0
version: 2.0.0
license: MIT
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
import sys
import uuid
from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Optional

from pydantic import BaseModel, Field

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    _HAS_XLSX = True
except Exception:
    _HAS_XLSX = False


def _load_source_lib():
    for p in (
        "/app/backend/data",
        "/opt/open-webui/tools",
        os.path.dirname(os.path.abspath(__file__)),
    ):
        if p and p not in sys.path and os.path.isdir(p):
            sys.path.insert(0, p)
    import lib_source_workbook as sw  # type: ignore

    return sw


def _slugify(text: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "-", (text or "workbook").strip().lower()).strip("-")
    return (s[:48] or "workbook")


def _parse_content(content: str | dict) -> dict:
    if isinstance(content, dict):
        return content
    raw = (content or "{}").strip()
    for fence in ("```json", "```"):
        if raw.startswith(fence):
            raw = raw[len(fence) :]
            if raw.endswith("```"):
                raw = raw[:-3]
            raw = raw.strip()
            break
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            return json.loads(m.group(0))
        raise


def _hex_fill(hex_color: str, default="141C2B"):
    h = (hex_color or default).strip().lstrip("#").upper()
    if len(h) != 6:
        h = default
    return PatternFill("solid", fgColor=h)


class Tools:
    def __init__(self):
        self.valves = self.Valves()

    class Valves(BaseModel):
        header_color: str = Field(default="141C2B", description="Header fill hex")
        accent_color: str = Field(default="7EA8D9", description="Accent fill hex")

    def _style_header_row(self, ws, row_idx: int, ncols: int, fill_hex: str):
        fill = _hex_fill(fill_hex)
        font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        thin = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin

    def _autosize(self, ws, max_width=40):
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            width = 8
            for cell in col:
                val = "" if cell.value is None else str(cell.value)
                width = max(width, min(max_width, len(val) + 2))
            ws.column_dimensions[letter].width = width

    def _write_sheet(self, wb, sheet_spec: dict, *, is_first: bool, header_hex: str):
        name = str(sheet_spec.get("name") or sheet_spec.get("title") or "Sheet")[:31]
        # sanitize sheet name
        name = re.sub(r"[:\\/?*\[\]]", "-", name) or "Sheet"
        if is_first:
            ws = wb.active
            ws.title = name
        else:
            # avoid duplicate titles
            base = name
            i = 2
            existing = set(wb.sheetnames)
            while name in existing:
                name = f"{base[:28]}_{i}"
                i += 1
            ws = wb.create_sheet(name)

        row = 1
        title = sheet_spec.get("title") or sheet_spec.get("heading")
        if title:
            ws.cell(row=row, column=1, value=str(title)).font = Font(
                bold=True, size=14, color=header_hex, name="Calibri"
            )
            row += 1
        if sheet_spec.get("subtitle"):
            ws.cell(row=row, column=1, value=str(sheet_spec["subtitle"])).font = Font(
                size=10, color="4B5563", name="Calibri"
            )
            row += 1
        if title or sheet_spec.get("subtitle"):
            row += 1

        # KPIs block
        stats = sheet_spec.get("stats") or sheet_spec.get("kpis") or []
        if stats:
            ws.cell(row=row, column=1, value="KPI").font = Font(bold=True)
            ws.cell(row=row, column=2, value="Value").font = Font(bold=True)
            self._style_header_row(ws, row, 2, header_hex)
            row += 1
            for st in stats:
                if isinstance(st, dict):
                    ws.cell(row=row, column=1, value=str(st.get("label", "")))
                    ws.cell(row=row, column=2, value=st.get("value", ""))
                else:
                    ws.cell(row=row, column=1, value=str(st))
                row += 1
            row += 1

        headers = sheet_spec.get("headers") or sheet_spec.get("columns") or []
        rows = sheet_spec.get("rows") or []
        table = sheet_spec.get("table")
        if isinstance(table, dict):
            headers = table.get("headers") or table.get("columns") or headers
            rows = table.get("rows") or rows

        if headers:
            for c, h in enumerate(headers, 1):
                ws.cell(row=row, column=c, value=str(h))
            self._style_header_row(ws, row, len(headers), header_hex)
            row += 1
            alt = PatternFill("solid", fgColor="F8FAFC")
            thin = Border(
                left=Side(style="thin", color="E5E7EB"),
                right=Side(style="thin", color="E5E7EB"),
                top=Side(style="thin", color="E5E7EB"),
                bottom=Side(style="thin", color="E5E7EB"),
            )
            for r_i, rdata in enumerate(rows):
                if isinstance(rdata, dict):
                    vals = [rdata.get(h, "") for h in headers]
                else:
                    vals = list(rdata) if isinstance(rdata, (list, tuple)) else [rdata]
                while len(vals) < len(headers):
                    vals.append("")
                for c, val in enumerate(vals[: len(headers)], 1):
                    cell = ws.cell(row=row, column=c, value=val)
                    cell.border = thin
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    if r_i % 2 == 1:
                        cell.fill = alt
                row += 1

        notes = sheet_spec.get("notes") or sheet_spec.get("bullets") or []
        if notes:
            row += 1
            ws.cell(row=row, column=1, value="Notes").font = Font(bold=True, color=header_hex)
            row += 1
            for n in notes:
                if isinstance(n, dict):
                    text = n.get("text") or n.get("title") or str(n)
                else:
                    text = str(n)
                ws.cell(row=row, column=1, value=f"• {text}")
                row += 1

        self._autosize(ws)
        ws.freeze_panes = "A2" if not title else "A4"
        return ws

    def _build_xlsx(self, spec: dict) -> bytes:
        wb = Workbook()
        header_hex = (spec.get("header_color") or self.valves.header_color).lstrip("#")
        sheets = spec.get("sheets") or []
        if not sheets:
            # single-sheet convenience
            sheets = [
                {
                    "name": spec.get("sheet_name") or "Report",
                    "title": spec.get("title"),
                    "subtitle": spec.get("subtitle"),
                    "headers": spec.get("headers"),
                    "rows": spec.get("rows"),
                    "table": spec.get("table"),
                    "stats": spec.get("stats") or spec.get("kpis"),
                    "notes": spec.get("notes") or spec.get("bullets"),
                }
            ]

        for i, sh in enumerate(sheets):
            if not isinstance(sh, dict):
                continue
            self._write_sheet(wb, sh, is_first=(i == 0), header_hex=header_hex)

        # meta sheet optional
        if spec.get("author") or spec.get("title"):
            meta = wb.create_sheet("_meta")
            meta["A1"] = "Title"
            meta["B1"] = spec.get("title") or ""
            meta["A2"] = "Author"
            meta["B2"] = spec.get("author") or ""
            meta["A3"] = "Generated"
            meta["B3"] = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
            meta.sheet_state = "hidden"

        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def _save(self, data: bytes, *, title: str, user_dict: Optional[dict]):
        slug = _slugify(title)
        day = datetime.now(timezone.utc).strftime("%Y%m%d")
        short = uuid.uuid4().hex[:6]
        filename = f"workbook-{slug}_{day}_{short}.xlsx"
        fid = str(uuid.uuid4())
        stored = f"{fid}_{filename}"
        uploads = "/app/backend/data/uploads"
        cache = "/app/backend/data/cache/files"
        try:
            os.makedirs(uploads, mode=0o775, exist_ok=True)
            os.makedirs(cache, mode=0o775, exist_ok=True)
            with open(os.path.join(uploads, stored), "wb") as fh:
                fh.write(data)
            with open(os.path.join(cache, filename), "wb") as fh:
                fh.write(data)
            uid = user_dict.get("id") if isinstance(user_dict, dict) else None
            db_path = "/app/backend/data/webui.db"
            file_hash = hashlib.sha256(data).hexdigest()
            now = int(datetime.now(timezone.utc).timestamp())
            meta = {
                "name": filename,
                "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "size": len(data),
                "file_hash": file_hash,
                "data": {},
            }
            con = sqlite3.connect(db_path)
            try:
                if not uid:
                    row = con.execute("select id from user where role='admin' limit 1").fetchone()
                    uid = row[0] if row else "system"
                cols = [r[1] for r in con.execute("pragma table_info(file)")]
                vals = {
                    "id": fid,
                    "user_id": uid,
                    "filename": filename,
                    "meta": json.dumps(meta),
                    "created_at": now,
                    "hash": file_hash,
                    "data": json.dumps({}),
                    "updated_at": now,
                    "path": f"/app/backend/data/uploads/{stored}",
                }
                use = [k for k in vals if k in cols]
                con.execute(
                    f"insert into file ({','.join(use)}) values ({','.join('?' for _ in use)})",
                    [vals[k] for k in use],
                )
                con.commit()
            finally:
                con.close()
            return filename, f"/api/v1/files/{fid}/content", None
        except Exception as exc:
            return filename, None, str(exc)

    async def generate_excel(
        self,
        content: str = "{}",
        source_file_id: str = "",
        source_filename: str = "",
        job: str = "",
        sheet_name: str = "",
        group_by: str = "Subject",
        __files__: Optional[list] = None,
        __event_emitter__: Any = None,
        __user__: Optional[dict] = None,
        __request__: Any = None,
    ) -> str:
        """Create a native Excel (.xlsx) workbook and return a download link.

        FULL-FILE RULE (mandatory when a spreadsheet is attached):
        - Pass source_file_id from the attached file (or leave blank to auto-pick chat .xlsx).
        - Set job to one of:
          - subject_group_report — group ALL rows by Subject (or group_by); one sheet per group + Overview
          - mirror_workbook — copy every sheet/row from the source file exactly
          - full_sheet_tables — export chosen sheet (all rows) as data sheets
        - Do NOT invent rows from chat/RAG snippets. This tool reads the complete file on disk.

        Optional `content` JSON (blank/new workbooks only — ignored for data when job+source set):
        { "title":"...", "sheets":[{"name":"...","headers":[...],"rows":[[...]]}] }

        Reproduce the returned markdown link EXACTLY (do NOT prefix sandbox:).
        """
        if not _HAS_XLSX:
            return (
                "[TOOL_RESULT — use as final reply]\n\n"
                "I couldn't generate the Excel file: openpyxl is not installed."
            )

        verify_md = ""
        job_l = (job or "").strip().lower()
        has_source_hint = bool(
            (source_file_id or "").strip()
            or (source_filename or "").strip()
            or job_l
            or __files__
        )

        # Auto job when files attached and job blank but content empty/default
        if has_source_hint and not job_l:
            # If content already has real sheets/rows, allow legacy path unless source_file_id set
            if (source_file_id or "").strip() or (source_filename or "").strip():
                job_l = "subject_group_report"

        if job_l in (
            "subject_group_report",
            "subjectwise",
            "subject_wise",
            "subject-wise",
            "mirror_workbook",
            "mirror",
            "full_sheet_tables",
            "full_sheet",
        ):
            try:
                sw = _load_source_lib()
                uid = __user__.get("id") if isinstance(__user__, dict) else None
                fid, filename, path = sw.resolve_file_path(
                    source_file_id=source_file_id,
                    source_filename=source_filename,
                    files=__files__ or [],
                    user_id=uid,
                )
                if __event_emitter__:
                    try:
                        await __event_emitter__(
                            {
                                "type": "status",
                                "data": {
                                    "description": f"FULL_FILE_READ · {filename}…",
                                    "done": False,
                                },
                            }
                        )
                    except Exception:
                        pass
                book = sw.load_workbook_rows(path)
                if job_l in ("mirror_workbook", "mirror"):
                    sheets = []
                    for s in book["sheets"]:
                        sheets.append(
                            {
                                "name": s["name"][:31],
                                "title": s["name"],
                                "headers": s["headers"],
                                "rows": [
                                    [("" if c is None else c) for c in row]
                                    for row in s["rows"]
                                ],
                                "stats": [
                                    {
                                        "label": "Rows",
                                        "value": str(s["row_count"]),
                                    }
                                ],
                                "notes": ["FULL_FILE_READ mirror — complete sheet"],
                            }
                        )
                    spec = {
                        "title": f"Mirror · {filename}",
                        "author": "OmniTech FULL_FILE_READ",
                        "sheets": sheets,
                    }
                    verify_md = (
                        f"**VERIFY · FULL_FILE_READ**\n- Source: `{filename}` (`{fid}`)\n"
                        f"- Sheets: {len(sheets)}\n"
                        f"- Total data rows: **{book['total_data_rows']}**"
                    )
                elif job_l in ("full_sheet_tables", "full_sheet"):
                    s = None
                    for cand in book["sheets"]:
                        if not sheet_name or sheet_name.lower() in cand["name"].lower():
                            s = cand
                            if sheet_name:
                                break
                    if s is None:
                        s = book["sheets"][0]
                    spec = {
                        "title": f"{s['name']} · full export",
                        "author": "OmniTech FULL_FILE_READ",
                        "sheets": [
                            {
                                "name": s["name"][:31],
                                "title": s["name"],
                                "headers": s["headers"],
                                "rows": [
                                    [("" if c is None else c) for c in row]
                                    for row in s["rows"]
                                ],
                                "stats": [
                                    {"label": "Rows", "value": str(s["row_count"])}
                                ],
                                "notes": ["FULL_FILE_READ — all rows"],
                            }
                        ],
                    }
                    verify_md = (
                        f"**VERIFY · FULL_FILE_READ**\n- Source: `{filename}` (`{fid}`)\n"
                        f"- Sheet: `{s['name']}`\n- Rows: **{s['row_count']}**"
                    )
                else:
                    # parse optional title from content
                    title = "Subject-wise Report"
                    try:
                        meta = _parse_content(content) if content and content.strip() not in ("{}", "") else {}
                        if isinstance(meta, dict) and meta.get("title"):
                            title = str(meta["title"])
                    except Exception:
                        pass
                    built = sw.build_subject_group_report(
                        book,
                        sheet_name=sheet_name or "Form Responses 1",
                        group_by=group_by or "Subject",
                        title=title,
                    )
                    spec = built["excel"]
                    verify_md = sw.verify_block_markdown(
                        built["verify"], fid, filename
                    )
            except Exception as exc:
                return (
                    "[TOOL_RESULT — use as final reply]\n\n"
                    f"FULL_FILE_READ failed: {exc}\n\n"
                    "Attach the .xlsx again and call generate_excel with "
                    "job='subject_group_report' (or mirror_workbook)."
                )
        else:
            try:
                spec = _parse_content(content)
            except Exception as exc:
                return f"[TOOL_RESULT — use as final reply]\n\nInvalid JSON for generate_excel: {exc}"
            # Guard: refuse tiny invented reports when chat has spreadsheet attaches
            if __files__:
                xlsx_attached = False
                for item in __files__ or []:
                    if not isinstance(item, dict):
                        continue
                    n = str(item.get("filename") or item.get("name") or "").lower()
                    if n.endswith((".xlsx", ".xlsm", ".xls", ".csv")):
                        xlsx_attached = True
                        break
                if xlsx_attached:
                    return (
                        "[TOOL_RESULT — use as final reply]\n\n"
                        "A spreadsheet is attached. Do **not** pass hand-built rows from chat/RAG. "
                        "Call again with job='subject_group_report' (or mirror_workbook / full_sheet_tables) "
                        "so the tool can FULL_FILE_READ the attached file."
                    )

        if __event_emitter__:
            try:
                await __event_emitter__(
                    {"type": "status", "data": {"description": "Building Excel workbook…", "done": False}}
                )
            except Exception:
                pass

        try:
            data = self._build_xlsx(spec)
        except Exception as exc:
            return f"[TOOL_RESULT — use as final reply]\n\nXLSX render failed: {exc}"

        fname, url, err = self._save(data, title=str(spec.get("title") or "workbook"), user_dict=__user__)
        if err or not url:
            return f"[TOOL_RESULT — use as final reply]\n\nXLSX built but save failed: {err}"

        kb = max(1, len(data) // 1024)
        head = verify_md + "\n\n" if verify_md else ""
        if __event_emitter__:
            try:
                await __event_emitter__(
                    {
                        "type": "message",
                        "data": {
                            "content": (
                                f"\n\n---\n\n📊 **Excel workbook ready** · {kb} KB\n\n"
                                f"{head}"
                                f"⬇️ [Download {fname}]({url})\n\n---\n"
                            )
                        },
                    }
                )
                await __event_emitter__(
                    {"type": "status", "data": {"description": "XLSX ready", "done": True}}
                )
            except Exception:
                pass

        return (
            "[TOOL_RESULT — reproduce the markdown link EXACTLY as written below "
            "(do NOT change the URL, do NOT prefix sandbox:). Do not include this line.]\n\n"
            f"{head}"
            "Here is the Excel workbook:\n\n"
            f"[{fname}]({url})"
        )
