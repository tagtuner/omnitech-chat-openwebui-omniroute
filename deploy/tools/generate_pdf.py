"""
title: Generate PDF
author: OmniTech
author_url: https://github.com
description: Generate native PDF reports from a JSON spec (cover, KPIs, sections, tables).
requirements: reportlab
required_open_webui_version: 0.4.0
version: 1.0.0
license: MIT
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
import uuid
from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Optional

from pydantic import BaseModel, Field

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
        KeepTogether,
        HRFlowable,
        ListFlowable,
        ListItem,
    )

    _HAS_RL = True
except Exception:
    _HAS_RL = False


def _slugify(text: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "-", (text or "report").strip().lower()).strip("-")
    return (s[:48] or "report")


def _hex(c: str, default: str = "141C2B") -> str:
    if not isinstance(c, str):
        return default
    h = c.strip().lstrip("#")
    if len(h) == 3:
        h = "".join(ch * 2 for ch in h)
    if len(h) != 6:
        return default
    try:
        int(h, 16)
        return h.upper()
    except ValueError:
        return default


def _color(c: str):
    h = _hex(c)
    return colors.HexColor(f"#{h}")


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
        # salvage first {...}
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            return json.loads(m.group(0))
        raise


class Tools:
    def __init__(self):
        self.valves = self.Valves()

    class Valves(BaseModel):
        default_theme: str = Field(
            default="navy",
            description="Theme: navy | slate | forest | midnight",
        )
        page_size: str = Field(
            default="A4",
            description="A4 or A4_landscape",
        )
        footer_label: str = Field(default="", description="Footer label override")

    _THEMES = {
        "navy": {"primary": "141C2B", "accent": "7EA8D9", "gold": "E8B83A", "ink": "1A1A2E"},
        "slate": {"primary": "27303F", "accent": "3B82F6", "gold": "F59E0B", "ink": "111827"},
        "forest": {"primary": "1A3D1B", "accent": "97BC62", "gold": "C9A227", "ink": "14261A"},
        "midnight": {"primary": "1E2761", "accent": "CADCFC", "gold": "C99A3B", "ink": "141A45"},
    }

    def _theme(self, spec: dict) -> dict:
        name = (spec.get("theme") or self.valves.default_theme or "navy").lower()
        if name == "auto":
            name = "navy"
        base = dict(self._THEMES.get(name, self._THEMES["navy"]))
        for k in ("primary", "accent", "gold", "ink"):
            if spec.get(k):
                base[k] = _hex(spec[k], base[k])
        return {"name": name, **base}

    def _styles(self, theme: dict):
        ss = getSampleStyleSheet()
        primary = _color(theme["primary"])
        accent = _color(theme["accent"])
        ink = _color(theme["ink"])
        styles = {
            "cover_title": ParagraphStyle(
                "cover_title",
                parent=ss["Title"],
                fontName="Helvetica-Bold",
                fontSize=26,
                leading=32,
                textColor=colors.white,
                alignment=TA_CENTER,
                spaceAfter=8,
            ),
            "cover_sub": ParagraphStyle(
                "cover_sub",
                parent=ss["Normal"],
                fontName="Helvetica",
                fontSize=12,
                leading=16,
                textColor=colors.Color(0.85, 0.9, 0.95),
                alignment=TA_CENTER,
                spaceAfter=6,
            ),
            "h1": ParagraphStyle(
                "h1",
                parent=ss["Heading1"],
                fontName="Helvetica-Bold",
                fontSize=16,
                leading=20,
                textColor=primary,
                spaceBefore=10,
                spaceAfter=6,
            ),
            "h2": ParagraphStyle(
                "h2",
                parent=ss["Heading2"],
                fontName="Helvetica-Bold",
                fontSize=12,
                leading=15,
                textColor=primary,
                spaceBefore=8,
                spaceAfter=4,
            ),
            "body": ParagraphStyle(
                "body",
                parent=ss["Normal"],
                fontName="Helvetica",
                fontSize=10,
                leading=14,
                textColor=ink,
                spaceAfter=4,
            ),
            "bullet": ParagraphStyle(
                "bullet",
                parent=ss["Normal"],
                fontName="Helvetica",
                fontSize=10,
                leading=13,
                textColor=ink,
                leftIndent=8,
            ),
            "kpi_val": ParagraphStyle(
                "kpi_val",
                parent=ss["Normal"],
                fontName="Helvetica-Bold",
                fontSize=18,
                leading=22,
                textColor=primary,
                alignment=TA_CENTER,
            ),
            "kpi_lbl": ParagraphStyle(
                "kpi_lbl",
                parent=ss["Normal"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                textColor=colors.HexColor("#4B5563"),
                alignment=TA_CENTER,
            ),
            "meta": ParagraphStyle(
                "meta",
                parent=ss["Normal"],
                fontName="Helvetica",
                fontSize=9,
                textColor=colors.HexColor("#6B7280"),
                alignment=TA_CENTER,
            ),
            "cell": ParagraphStyle(
                "cell",
                parent=ss["Normal"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                textColor=ink,
            ),
            "cell_h": ParagraphStyle(
                "cell_h",
                parent=ss["Normal"],
                fontName="Helvetica-Bold",
                fontSize=8,
                leading=10,
                textColor=colors.white,
            ),
        }
        styles["_accent"] = accent
        styles["_primary"] = primary
        styles["_gold"] = _color(theme["gold"])
        return styles

    def _esc(self, text: Any) -> str:
        s = "" if text is None else str(text)
        return (
            s.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\n", "<br/>")
        )

    def _build_pdf(self, spec: dict) -> bytes:
        theme = self._theme(spec)
        styles = self._styles(theme)
        page = landscape(A4) if str(self.valves.page_size).lower().endswith("landscape") else A4
        if str(spec.get("orientation", "")).lower() == "landscape":
            page = landscape(A4)

        buf = BytesIO()
        footer = (spec.get("footer") or self.valves.footer_label or "").strip()
        title = spec.get("title") or "Report"

        def _footer(canvas, doc):
            canvas.saveState()
            canvas.setFillColor(_color(theme["primary"]))
            canvas.rect(0, 0, page[0], 14 * mm, fill=1, stroke=0)
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica", 8)
            label = footer or title
            canvas.drawString(15 * mm, 5 * mm, label[:80])
            canvas.drawRightString(page[0] - 15 * mm, 5 * mm, f"Page {doc.page}")
            canvas.restoreState()

        doc = SimpleDocTemplate(
            buf,
            pagesize=page,
            leftMargin=16 * mm,
            rightMargin=16 * mm,
            topMargin=16 * mm,
            bottomMargin=20 * mm,
            title=str(title)[:120],
            author=str(spec.get("author") or "OmniTech Chat")[:80],
        )

        story: list = []
        # Cover
        story.append(Spacer(1, 28 * mm))
        cover_data = [[Paragraph(self._esc(title), styles["cover_title"])]]
        if spec.get("subtitle"):
            cover_data.append([Paragraph(self._esc(spec["subtitle"]), styles["cover_sub"])])
        meta_bits = []
        if spec.get("author"):
            meta_bits.append(str(spec["author"]))
        if spec.get("date"):
            meta_bits.append(str(spec["date"]))
        else:
            meta_bits.append(datetime.now(timezone.utc).strftime("%Y-%m-%d"))
        cover_data.append([Paragraph(self._esc(" · ".join(meta_bits)), styles["cover_sub"])])
        cover = Table(cover_data, colWidths=[page[0] - 40 * mm])
        cover.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), _color(theme["primary"])),
                    ("TOPPADDING", (0, 0), (-1, -1), 14),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
                    ("LEFTPADDING", (0, 0), (-1, -1), 18),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 18),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        story.append(cover)
        story.append(Spacer(1, 10 * mm))
        if spec.get("eyebrow"):
            story.append(Paragraph(self._esc(spec["eyebrow"]), styles["meta"]))
        story.append(PageBreak())

        sections = spec.get("sections") or spec.get("pages") or []
        if not sections and spec.get("slides"):
            # tolerate slides-shaped JSON from models that confuse tools
            sections = []
            for s in spec["slides"]:
                sections.append(
                    {
                        "heading": s.get("title") or s.get("heading") or "Section",
                        "body": s.get("body") or s.get("subtitle") or "",
                        "bullets": s.get("bullets") or s.get("points") or s.get("items") or [],
                        "stats": s.get("stats") or [],
                        "table": s.get("table"),
                    }
                )

        for sec in sections:
            if not isinstance(sec, dict):
                continue
            block = []
            heading = sec.get("heading") or sec.get("title") or ""
            if heading:
                block.append(Paragraph(self._esc(heading), styles["h1"]))
                block.append(
                    HRFlowable(
                        width="100%",
                        thickness=1.2,
                        color=styles["_accent"],
                        spaceAfter=6,
                    )
                )
            if sec.get("subheading"):
                block.append(Paragraph(self._esc(sec["subheading"]), styles["h2"]))

            # KPIs / stats
            stats = sec.get("stats") or sec.get("kpis") or []
            if stats:
                cells = []
                for st in stats[:6]:
                    if isinstance(st, dict):
                        val = st.get("value", "")
                        lbl = st.get("label", "")
                    else:
                        val, lbl = str(st), ""
                    cells.append(
                        [
                            Paragraph(self._esc(val), styles["kpi_val"]),
                            Paragraph(self._esc(lbl), styles["kpi_lbl"]),
                        ]
                    )
                # one row of KPI cards
                row = []
                for card in cells:
                    t = Table([[card[0]], [card[1]]], colWidths=[(page[0] - 40 * mm) / max(len(cells), 1) - 4])
                    t.setStyle(
                        TableStyle(
                            [
                                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F6FA")),
                                ("BOX", (0, 0), (-1, -1), 0.5, styles["_accent"]),
                                ("TOPPADDING", (0, 0), (-1, -1), 8),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ]
                        )
                    )
                    row.append(t)
                kpi_table = Table([row], colWidths=[(page[0] - 32 * mm) / len(row)] * len(row))
                kpi_table.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]))
                block.append(Spacer(1, 2 * mm))
                block.append(kpi_table)
                block.append(Spacer(1, 4 * mm))

            body = sec.get("body") or sec.get("text") or sec.get("content") or ""
            if body:
                for para in str(body).split("\n"):
                    if para.strip():
                        block.append(Paragraph(self._esc(para.strip()), styles["body"]))

            bullets = sec.get("bullets") or sec.get("points") or sec.get("items") or []
            if bullets:
                items = []
                for b in bullets:
                    if isinstance(b, dict):
                        line = b.get("text") or b.get("title") or b.get("description") or str(b)
                    else:
                        line = str(b)
                    items.append(ListItem(Paragraph(self._esc(line), styles["bullet"]), leftIndent=10))
                block.append(ListFlowable(items, bulletType="bullet", start="•"))

            # table
            table = sec.get("table")
            if isinstance(table, dict):
                headers = table.get("headers") or table.get("columns") or []
                rows = table.get("rows") or []
                if headers and rows:
                    data = [[Paragraph(self._esc(h), styles["cell_h"]) for h in headers]]
                    for r in rows:
                        if isinstance(r, dict):
                            data.append(
                                [
                                    Paragraph(self._esc(r.get(h, "")), styles["cell"])
                                    for h in headers
                                ]
                            )
                        else:
                            vals = list(r) if isinstance(r, (list, tuple)) else [r]
                            while len(vals) < len(headers):
                                vals.append("")
                            data.append(
                                [
                                    Paragraph(self._esc(vals[i]), styles["cell"])
                                    for i in range(len(headers))
                                ]
                            )
                    col_w = (page[0] - 32 * mm) / max(len(headers), 1)
                    t = Table(data, colWidths=[col_w] * len(headers), repeatRows=1)
                    t.setStyle(
                        TableStyle(
                            [
                                ("BACKGROUND", (0, 0), (-1, 0), _color(theme["primary"])),
                                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D1D5DB")),
                                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("TOPPADDING", (0, 0), (-1, -1), 4),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                            ]
                        )
                    )
                    block.append(Spacer(1, 3 * mm))
                    block.append(t)

            if sec.get("page_break"):
                story.append(KeepTogether(block))
                story.append(PageBreak())
            else:
                story.extend(block)
                story.append(Spacer(1, 6 * mm))

        # Closing
        closing = spec.get("closing")
        if isinstance(closing, dict):
            story.append(PageBreak())
            story.append(Paragraph(self._esc(closing.get("title") or "Recommendations"), styles["h1"]))
            for b in closing.get("takeaways") or closing.get("bullets") or []:
                story.append(Paragraph(f"• {self._esc(b)}", styles["body"]))
            if closing.get("contact"):
                story.append(Spacer(1, 4 * mm))
                story.append(Paragraph(self._esc(closing["contact"]), styles["meta"]))

        if len(story) <= 3:
            story.append(Paragraph("No sections provided in the JSON spec.", styles["body"]))

        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        return buf.getvalue()

    def _save(self, data: bytes, *, title: str, user_dict: Optional[dict]) -> tuple[str, Optional[str], Optional[str]]:
        slug = _slugify(title)
        day = datetime.now(timezone.utc).strftime("%Y%m%d")
        short = uuid.uuid4().hex[:6]
        filename = f"report-{slug}_{day}_{short}.pdf"
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
                "content_type": "application/pdf",
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

    async def generate_pdf(
        self,
        content: str = "{}",
        __event_emitter__: Any = None,
        __user__: Optional[dict] = None,
        __request__: Any = None,
    ) -> str:
        """Create a native PDF report and return a download link. Use whenever the
        user asks for a PDF, report PDF, printable report, or similar.

        The `content` parameter MUST be a SINGLE JSON string (no markdown fence).
        Structure:

        {
          "title": "Report title",
          "subtitle": "optional",
          "author": "School / name",
          "theme": "navy",
          "orientation": "portrait",
          "footer": "Students Progress Report 2026",
          "sections": [
            {
              "heading": "KPIs",
              "stats": [{"value": "51%", "label": "On target"}],
              "body": "optional paragraph",
              "bullets": ["point 1", "point 2"],
              "table": {
                "headers": ["Name", "Forecast", "CAIE", "Δ"],
                "rows": [["Aisha", "A", "B", "-1"]]
              },
              "page_break": false
            }
          ],
          "closing": {
            "title": "Recommendations",
            "takeaways": ["..."]
          }
        }

        DESIGN: Prefer navy theme for Intellect/Ghazala reports. Keep tables readable.
        Reproduce the returned markdown link EXACTLY (do NOT prefix sandbox:).
        """
        if not _HAS_RL:
            return (
                "[TOOL_RESULT — use as final reply]\n\n"
                "I couldn't generate the PDF: reportlab is not installed."
            )
        try:
            spec = _parse_content(content)
        except Exception as exc:
            return (
                "[TOOL_RESULT — use as final reply]\n\n"
                f"Invalid JSON for generate_pdf: {exc}"
            )

        if __event_emitter__:
            try:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {"description": "Building PDF…", "done": False},
                    }
                )
            except Exception:
                pass

        try:
            pdf_bytes = self._build_pdf(spec)
        except Exception as exc:
            return (
                "[TOOL_RESULT — use as final reply]\n\n"
                f"PDF render failed: {exc}"
            )

        fname, url, err = self._save(
            pdf_bytes, title=str(spec.get("title") or "report"), user_dict=__user__
        )
        if err or not url:
            return (
                "[TOOL_RESULT — use as final reply]\n\n"
                f"PDF built but save failed: {err}"
            )

        kb = max(1, len(pdf_bytes) // 1024)
        if __event_emitter__:
            try:
                await __event_emitter__(
                    {
                        "type": "message",
                        "data": {
                            "content": (
                                f"\n\n---\n\n📄 **PDF ready** · {kb} KB\n\n"
                                f"⬇️ [Download {fname}]({url})\n\n---\n"
                            )
                        },
                    }
                )
                await __event_emitter__(
                    {"type": "status", "data": {"description": "PDF ready", "done": True}}
                )
            except Exception:
                pass

        return (
            "[TOOL_RESULT — reproduce the markdown link EXACTLY as written below "
            "(do NOT change the URL, do NOT prefix sandbox:). Do not include this line.]\n\n"
            "Here is the PDF report:\n\n"
            f"[{fname}]({url})"
        )
