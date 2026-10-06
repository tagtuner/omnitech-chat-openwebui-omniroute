"""
OmniTech full-source workbook reader for OWUI tools.
Reads COMPLETE .xlsx/.xlsm/.csv from disk — never RAG chunks.
Deploy to: /opt/open-webui/tools/lib_source_workbook.py
           and /app/backend/data/lib_source_workbook.py (container)
"""
from __future__ import annotations

import csv
import json
import os
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Optional

DB_PATHS = (
    "/app/backend/data/webui.db",
    "/var/lib/docker/volumes/open-webui_open-webui-data/_data/webui.db",
)
UPLOAD_ROOTS = (
    "/app/backend/data/uploads",
    "/var/lib/docker/volumes/open-webui_open-webui-data/_data/uploads",
)


def _db() -> str:
    for p in DB_PATHS:
        if os.path.isfile(p):
            return p
    return DB_PATHS[0]


def _short_header(h: str) -> str:
    h = (h or "").strip()
    if not h:
        return h
    m = re.search(r"\[([^\]]+)\]", h)
    if h.startswith("Q1."):
        return "Q1 Attendance"
    if h.startswith("Q2.") and m:
        return "Q2 " + m.group(1).rstrip(".")
    if h.startswith("Q3.") and m:
        return "Q3 " + m.group(1).rstrip(".")
    if h.startswith("Q4.") and m:
        return "Q4 " + m.group(1).rstrip(".")
    if h.startswith("Q5."):
        return "Q5 Progress"
    if h.startswith("Q6."):
        return "Q6 Confidence"
    if h.startswith("Q7."):
        return "Q7 Helps most"
    if h.startswith("Q8."):
        return "Q8 Does well"
    if h.startswith("Q9."):
        return "Q9 Improve"
    if h.startswith("Q10."):
        return "Q10 More support"
    if h.startswith("Q11."):
        return "Q11 Extra for CAIE"
    if h.startswith("Q12."):
        return "Q12 Other feedback"
    return h if len(h) <= 80 else h[:77] + "..."


def resolve_file_path(
    *,
    source_file_id: str = "",
    source_filename: str = "",
    files: Optional[list] = None,
    user_id: Optional[str] = None,
) -> tuple[str, str, str]:
    """
    Return (file_id, filename, absolute_path).
    Prefers explicit id, then __files__ attach list, then DB lookup by filename.
    """
    files = files or []
    fid = (source_file_id or "").strip()
    fname_hint = (source_filename or "").strip()

    # From chat attach metadata
    candidates: list[dict] = []
    for item in files:
        if not isinstance(item, dict):
            continue
        if item.get("type", "file") not in ("file", None, ""):
            # allow missing type
            if item.get("type") not in (None, "", "file"):
                continue
        item_id = item.get("id") or item.get("file_id") or item.get("url")
        if isinstance(item_id, str) and item_id.startswith(("http://", "https://", "data:")):
            continue
        if not item_id and not item.get("filename") and not item.get("name"):
            continue
        candidates.append(item)

    if not fid and candidates:
        if fname_hint:
            low = fname_hint.lower()
            for item in candidates:
                n = str(item.get("filename") or item.get("name") or "").lower()
                if low in n or n in low:
                    fid = str(item.get("id") or item.get("file_id") or "")
                    break
        if not fid:
            # newest / first spreadsheet-like
            for item in candidates:
                n = str(item.get("filename") or item.get("name") or "").lower()
                if n.endswith((".xlsx", ".xlsm", ".xls", ".csv")):
                    fid = str(item.get("id") or item.get("file_id") or "")
                    break
        if not fid and candidates:
            fid = str(candidates[0].get("id") or candidates[0].get("file_id") or "")

    con = sqlite3.connect(_db())
    con.row_factory = sqlite3.Row
    try:
        row = None
        if fid:
            row = con.execute(
                "SELECT id, user_id, filename, path FROM file WHERE id=?", (fid,)
            ).fetchone()
        if row is None and fname_hint:
            like = f"%{fname_hint}%"
            q = "SELECT id, user_id, filename, path FROM file WHERE filename LIKE ? ORDER BY created_at DESC LIMIT 5"
            rows = list(con.execute(q, (like,)))
            if user_id:
                mine = [r for r in rows if r["user_id"] == user_id]
                row = mine[0] if mine else (rows[0] if rows else None)
            else:
                row = rows[0] if rows else None
        if row is None:
            raise FileNotFoundError(
                "Source spreadsheet not found. Attach an .xlsx/.csv and pass source_file_id, "
                "or ensure the file is still in this chat."
            )
        fid = row["id"]
        filename = row["filename"] or "source.xlsx"
        path = (row["path"] or "").strip()
    finally:
        con.close()

    # Resolve path on disk
    tried = []
    if path:
        tried.append(path)
        if os.path.isfile(path):
            return fid, filename, path
        # container vs host mapping
        if path.startswith("/app/backend/data/"):
            alt = path.replace(
                "/app/backend/data/",
                "/var/lib/docker/volumes/open-webui_open-webui-data/_data/",
                1,
            )
            tried.append(alt)
            if os.path.isfile(alt):
                return fid, filename, alt

    for root in UPLOAD_ROOTS:
        if not os.path.isdir(root):
            continue
        # exact stored name pattern {id}_filename
        for p in Path(root).glob(f"{fid}_*"):
            if p.is_file():
                return fid, filename, str(p)
        direct = Path(root) / filename
        if direct.is_file():
            return fid, filename, str(direct)

    raise FileNotFoundError(
        f"DB has file id={fid} name={filename} but disk path missing. Tried: {tried}"
    )


def load_workbook_rows(path: str) -> dict[str, Any]:
    """
    Full read. Returns:
      {
        path, sheets: [{name, headers, rows(list[list]), row_count}],
        total_data_rows
      }
    """
    low = path.lower()
    if low.endswith(".csv"):
        with open(path, "r", encoding="utf-8-sig", newline="") as fh:
            reader = csv.reader(fh)
            all_rows = [list(r) for r in reader]
        headers = [str(c) if c is not None else "" for c in (all_rows[0] if all_rows else [])]
        data = all_rows[1:] if len(all_rows) > 1 else []
        return {
            "path": path,
            "sheets": [
                {
                    "name": "Sheet1",
                    "headers": headers,
                    "rows": data,
                    "row_count": len(data),
                }
            ],
            "total_data_rows": len(data),
        }

    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    sheets = []
    total = 0
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            rows_iter = ws.iter_rows(values_only=True)
            try:
                header_row = next(rows_iter)
            except StopIteration:
                sheets.append({"name": name, "headers": [], "rows": [], "row_count": 0})
                continue
            headers = [
                "" if c is None else str(c).strip() for c in header_row
            ]
            # trim trailing empty headers
            while headers and headers[-1] == "":
                headers.pop()
            data = []
            for raw in rows_iter:
                vals = list(raw[: len(headers)]) if headers else list(raw)
                if all(v is None or str(v).strip() == "" for v in vals):
                    continue
                # pad
                if len(vals) < len(headers):
                    vals.extend([None] * (len(headers) - len(vals)))
                data.append(vals)
            total += len(data)
            sheets.append(
                {
                    "name": name,
                    "headers": headers,
                    "rows": data,
                    "row_count": len(data),
                }
            )
    finally:
        wb.close()
    return {"path": path, "sheets": sheets, "total_data_rows": total}


def _find_sheet(book: dict, sheet_name: str = "") -> dict:
    sheets = book["sheets"]
    if not sheets:
        raise ValueError("Workbook has no sheets")
    want = (sheet_name or "").strip().lower()
    if want:
        for s in sheets:
            if s["name"].strip().lower() == want:
                return s
        for s in sheets:
            if want in s["name"].strip().lower():
                return s
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available: {[s['name'] for s in sheets]}"
        )
    # prefer Form Responses
    for s in sheets:
        if "form response" in s["name"].lower() or s["name"].lower() == "responses":
            return s
    # widest sheet
    return max(sheets, key=lambda s: s["row_count"])


def _col_index(headers: list[str], *names: str) -> Optional[int]:
    norms = [re.sub(r"\s+", " ", (h or "").strip().lower()) for h in headers]
    for name in names:
        n = name.lower()
        for i, h in enumerate(norms):
            if h == n:
                return i
    for name in names:
        n = name.lower()
        for i, h in enumerate(norms):
            if n in h:
                return i
    return None


def build_subject_group_report(
    book: dict,
    *,
    sheet_name: str = "",
    group_by: str = "Subject",
    title: str = "Subject-wise Report",
) -> dict[str, Any]:
    """
    Deterministic full-file subject (or any group) report for generate_excel/docx/pdf.
    Returns excel-style spec + verify block.
    """
    sheet = _find_sheet(book, sheet_name)
    headers = sheet["headers"]
    rows = sheet["rows"]
    g_idx = _col_index(headers, group_by, "subject", "Subject")
    if g_idx is None:
        raise ValueError(
            f"Group column '{group_by}' not found. Headers: {headers[:20]}"
        )
    class_idx = _col_index(headers, "Class", "Level")
    short_headers = [_short_header(h) for h in headers]

    groups: dict[str, list[list]] = defaultdict(list)
    for r in rows:
        key = r[g_idx]
        if key is None or str(key).strip() == "":
            key = "(blank)"
        groups[str(key).strip()].append(r)

    # stable alpha order
    ordered = sorted(groups.items(), key=lambda kv: kv[0].lower())

    verify = {
        "source_sheet": sheet["name"],
        "group_by": headers[g_idx],
        "total_rows": len(rows),
        "group_count": len(ordered),
        "groups": {k: len(v) for k, v in ordered},
        "mode": "FULL_FILE_READ",
        "assumption": False,
    }

    excel_sheets = []
    doc_sections = []
    pdf_sections = []

    for gname, grows in ordered:
        # class breakdown
        class_counts: Counter = Counter()
        if class_idx is not None:
            for r in grows:
                cv = r[class_idx]
                class_counts[str(cv).strip() if cv is not None else "(blank)"] += 1

        # frequency for closed questions (non-text heavy)
        open_keys = {"Q7", "Q8", "Q9", "Q10", "Q11", "Q12", "Helps most", "Does well", "Improve", "More support", "Extra for CAIE", "Other feedback"}
        freq_rows = []
        open_rows = []
        for hi, h in enumerate(headers):
            if hi == g_idx:
                continue
            sh = short_headers[hi]
            vals = []
            for r in grows:
                if hi >= len(r):
                    continue
                v = r[hi]
                if v is None or str(v).strip() == "":
                    continue
                vals.append(str(v).strip())
            if not vals:
                continue
            is_open = any(k.lower() in sh.lower() for k in open_keys) or sh.startswith("Q7") or sh.startswith("Q8") or sh.startswith("Q9") or sh.startswith("Q10") or sh.startswith("Q11") or sh.startswith("Q12")
            # also treat long average length as open
            avg_len = sum(len(v) for v in vals) / max(1, len(vals))
            if is_open or avg_len > 40 or any("," in v and len(v) > 50 for v in vals):
                for v in vals:
                    open_rows.append([sh, v])
            else:
                c = Counter(vals)
                for ans, n in c.most_common():
                    pct = f"{(100.0 * n / len(grows)):.0f}%"
                    freq_rows.append([sh, ans, n, pct])

        stats = [
            {"label": "Total Responses", "value": str(len(grows))},
        ]
        for ck, cn in sorted(class_counts.items()):
            stats.append({"label": f"Class {ck}", "value": str(cn)})

        # Excel sheet
        excel_sheets.append(
            {
                "name": gname[:31],
                "title": f"{gname} — Survey Report",
                "stats": stats,
                "headers": ["Question", "Answer", "Count", "Share"],
                "rows": freq_rows,
                "notes": [
                    f"FULL_FILE_READ · n={len(grows)} · source sheet '{sheet['name']}'",
                    "Open-text answers follow in Open Text block (all rows, not sampled).",
                ]
                + (
                    ["--- Open text ---"]
                    + [f"{a}: {b}" for a, b in open_rows[:500]]
                    if open_rows
                    else ["(no open-text answers)"]
                ),
            }
        )
        # also dedicated open-text table sheet if many
        if open_rows:
            excel_sheets.append(
                {
                    "name": (gname[:24] + " text")[:31],
                    "title": f"{gname} — Open text (all)",
                    "headers": ["Question", "Answer"],
                    "rows": open_rows,
                    "stats": [{"label": "Open answers", "value": str(len(open_rows))}],
                }
            )

        bullets = [f"Total responses: {len(grows)}"]
        if class_counts:
            bullets.append(
                "Class mix: "
                + ", ".join(f"{k}={v}" for k, v in sorted(class_counts.items()))
            )
        # top closed themes
        for sh, ans, n, pct in freq_rows[:12]:
            bullets.append(f"{sh}: {ans} ({n}, {pct})")
        for a, b in open_rows[:15]:
            bullets.append(f"{a}: {b}")

        table_headers = ["Question", "Answer", "Count", "Share"]
        table_rows = [[str(x) for x in r] for r in freq_rows[:80]]
        sec = {
            "heading": gname,
            "body": f"Full-file aggregate for {gname}. Responses: {len(grows)}.",
            "bullets": bullets[:40],
            "stats": [{"value": s["value"], "label": s["label"]} for s in stats],
            "table": {"headers": table_headers, "rows": table_rows},
        }
        doc_sections.append(sec)
        pdf_sections.append(sec)

    # overview sheet first
    overview_rows = [[k, len(v)] for k, v in ordered]
    excel_sheets.insert(
        0,
        {
            "name": "Overview",
            "title": title,
            "stats": [
                {"label": "Total responses", "value": str(len(rows))},
                {"label": "Groups", "value": str(len(ordered))},
                {"label": "Mode", "value": "FULL_FILE_READ"},
            ],
            "headers": [headers[g_idx], "Responses"],
            "rows": overview_rows,
            "notes": [
                "VERIFY: sum of group counts must equal Total responses.",
                f"Source sheet: {sheet['name']}",
                json.dumps(verify["groups"], ensure_ascii=False),
            ],
        },
    )

    return {
        "verify": verify,
        "excel": {"title": title, "author": "OmniTech FULL_FILE_READ", "sheets": excel_sheets},
        "docx": {
            "title": title,
            "subtitle": f"FULL_FILE_READ · {len(rows)} rows · {len(ordered)} groups",
            "author": "OmniTech",
            "sections": [
                {
                    "heading": "Verification",
                    "body": (
                        f"Mode=FULL_FILE_READ. Total rows={len(rows)}. "
                        f"Groups={len(ordered)}. Counts: {verify['groups']}"
                    ),
                    "bullets": [f"{k}: {n}" for k, n in overview_rows],
                },
                *doc_sections,
            ],
        },
        "pdf": {
            "title": title,
            "subtitle": f"FULL_FILE_READ · {len(rows)} rows",
            "sections": [
                {
                    "heading": "Verification",
                    "body": f"Total rows={len(rows)}; groups={len(ordered)}",
                    "bullets": [f"{k}: {n}" for k, n in overview_rows],
                    "table": {
                        "headers": [headers[g_idx], "Responses"],
                        "rows": [[str(a), str(b)] for a, b in overview_rows],
                    },
                },
                *pdf_sections,
            ],
        },
        "slides": {
            "title": title,
            "slides": [
                {
                    "title": "Verification (FULL_FILE_READ)",
                    "bullets": [
                        f"Total rows: {len(rows)}",
                        f"Groups: {len(ordered)}",
                        *[f"{k}: {n}" for k, n in overview_rows],
                    ],
                },
                *[
                    {
                        "title": gname[:60],
                        "bullets": [
                            f"Responses: {len(grows)}",
                            *[
                                f"{short_headers[i]} sample kept in Excel/PDF full tables"
                                for i in range(min(3, len(short_headers)))
                            ],
                        ],
                    }
                    for gname, grows in ordered
                ],
            ],
        },
    }


def _header_idx_map(headers: list[str]) -> dict[str, int]:
    """Map common progress-report columns → index."""
    norms = [re.sub(r"\s+", " ", (h or "").strip().lower()) for h in headers]
    out: dict[str, int] = {}

    def find(*names: str) -> Optional[int]:
        for n in names:
            n = n.lower()
            for i, h in enumerate(norms):
                if h == n:
                    return i
        for n in names:
            n = n.lower()
            for i, h in enumerate(norms):
                if n in h:
                    return i
        return None

    out["name"] = find("name", "student", "student name")
    out["subject"] = find("subject", "subj")
    out["forecast"] = find("forecast", "f", "predicted", "pred")
    out["caie"] = find("caie", "c", "actual", "result")
    # letter columns sometimes only F / C as headers
    if out.get("forecast") is None:
        for i, h in enumerate(norms):
            if h in ("f", "fc"):
                out["forecast"] = i
                break
    if out.get("caie") is None:
        for i, h in enumerate(norms):
            if h in ("c", "cie"):
                out["caie"] = i
                break
    return out


def _pick_level_sheet(book: dict, level: str) -> Optional[dict]:
    """Prefer AS GB / A2 GB, then AS/A2, then As/A2."""
    level = level.upper()
    names = [s["name"] for s in book["sheets"]]
    prefs = []
    if level == "AS":
        prefs = ["AS GB", "ASGB", "AS", "As", "AS Girls", "AS Students"]
    else:
        prefs = ["A2 GB", "A2GB", "A2", "A2 Girls", "A2 Students"]
    lower_map = {n.lower().strip(): s for n, s in zip(names, book["sheets"])}
    for p in prefs:
        s = lower_map.get(p.lower())
        if s and s["row_count"] > 0:
            return s
    # fuzzy contains
    for s in book["sheets"]:
        n = s["name"].lower()
        if level.lower() in n and "gb" in n and s["row_count"] > 0:
            return s
    for s in book["sheets"]:
        n = s["name"].lower().strip()
        if n == level.lower() and s["row_count"] > 0:
            return s
    return None


def _sheet_to_forecast_rows(sheet: dict) -> list[dict]:
    headers = sheet["headers"]
    idx = _header_idx_map(headers)
    need = ["name", "subject", "forecast", "caie"]
    missing = [k for k in need if idx.get(k) is None]
    if missing:
        raise ValueError(
            f"Sheet '{sheet['name']}' missing columns {missing}. Headers: {headers[:20]}"
        )
    rows = []
    for raw in sheet["rows"]:
        def cell(key: str):
            i = idx[key]
            if i is None or i >= len(raw):
                return ""
            v = raw[i]
            return "" if v is None else str(v).strip()

        name = cell("name")
        subject = cell("subject")
        if not name and not subject:
            continue
        # skip header repeats
        if name.lower() in ("name", "student", "student name"):
            continue
        rows.append(
            {
                "name": name,
                "subject": subject,
                "forecast": cell("forecast"),
                "caie": cell("caie"),
            }
        )
    return rows


def build_forecast_pack_from_workbook(
    book: dict,
    *,
    title: str = "Students Progress Report",
    cohort: str = "GB",
    as_sheet: str = "",
    a2_sheet: str = "",
) -> dict[str, Any]:
    """
    FULL_FILE_READ → generate_forecast_pack spec.
    Prefers sheets AS GB + A2 GB (full cohort).
    """
    if as_sheet:
        as_s = _find_sheet(book, as_sheet)
    else:
        as_s = _pick_level_sheet(book, "AS")
    if a2_sheet:
        a2_s = _find_sheet(book, a2_sheet)
    else:
        a2_s = _pick_level_sheet(book, "A2")

    if as_s is None and a2_s is None:
        available = [s["name"] for s in book["sheets"]]
        raise ValueError(
            "Could not find AS/A2 sheets (prefer 'AS GB' + 'A2 GB'). "
            f"Available: {available}"
        )

    as_rows = _sheet_to_forecast_rows(as_s) if as_s else []
    a2_rows = _sheet_to_forecast_rows(a2_s) if a2_s else []

    verify = {
        "mode": "FULL_FILE_READ",
        "assumption": False,
        "as_sheet": as_s["name"] if as_s else None,
        "a2_sheet": a2_s["name"] if a2_s else None,
        "as_rows": len(as_rows),
        "a2_rows": len(a2_rows),
        "total_rows": len(as_rows) + len(a2_rows),
        "cohort": cohort or "GB",
    }
    spec = {
        "title": title,
        "subtitle": f"FULL_FILE_READ · cohort {cohort or 'GB'}",
        "author": "OmniTech FULL_FILE_READ",
        "cohort": cohort or "GB",
        "as_rows": as_rows,
        "a2_rows": a2_rows,
        "actions": [],
    }
    return {"verify": verify, "spec": spec}


def verify_forecast_markdown(verify: dict, file_id: str, filename: str) -> str:
    return "\n".join(
        [
            "**VERIFY · FULL_FILE_READ** (not RAG chunks)",
            f"- Source file: `{filename}` (`{file_id}`)",
            f"- AS sheet: `{verify.get('as_sheet')}` · rows **{verify.get('as_rows')}**",
            f"- A2 sheet: `{verify.get('a2_sheet')}` · rows **{verify.get('a2_rows')}**",
            f"- Total student rows: **{verify.get('total_rows')}**",
            f"- Cohort: `{verify.get('cohort')}`",
        ]
    )


def verify_block_markdown(verify: dict, file_id: str, filename: str) -> str:
    groups = verify.get("groups") or {}
    lines = [
        "**VERIFY · FULL_FILE_READ** (not RAG chunks)",
        f"- Source file: `{filename}` (`{file_id}`)",
        f"- Sheet: `{verify.get('source_sheet')}`",
        f"- Total rows: **{verify.get('total_rows')}**",
        f"- Groups: **{verify.get('group_count')}**",
    ]
    for k, n in groups.items():
        lines.append(f"  - {k}: {n}")
    return "\n".join(lines)
