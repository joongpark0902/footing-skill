# -*- coding: utf-8 -*-
"""보고서 draft / 근거 조서를 하나의 정규화 JSON 스키마로 추출한다.

지원 형식: .pptx  .xlsx/.xlsm  .htm/.html  .docx  .pdf  .dsd(DART 편집기)

사용법
    python extract.py <파일...> -o out.json [--as report|workpaper]

출력 스키마
    {"source":..., "kind":"report|workpaper", "type":"pptx", "meta":{...},
     "blocks":[{"id":int, "loc":str, "kind":"cell|para", "text":str,
                "number":float|None, "num_kind":"amount|percent|date|None",
                "context":{"row":str,"col":str,"table":str},
                "style":{...}, "geom":{...}}]}

blocks 는 파일 종류가 달라도 같은 모양이다. 하위 스크립트(tie_out / format_audit)는
이 스키마만 알면 되고 원본 포맷은 몰라도 된다.
"""
import argparse
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

EMU_CM = 360000.0

# ---------------------------------------------------------------- 숫자 정규화

_DATE_RE = re.compile(r"^\s*(19|20)\d{2}\s*[.\-/년]\s*\d{1,2}\s*[.\-/월]\s*\d{1,2}\s*일?\s*$")
_NUM_RE = re.compile(r"^\s*\(?\s*[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*\)?\s*(원|%|주|건|배)?\s*$")
_BARE_RE = re.compile(r"^\s*\(?\s*[-+]?\d+(?:\.\d+)?\s*\)?\s*(원|%|주|건|배)?\s*$")


def parse_number(text):
    """셀/문단 텍스트 하나가 통째로 수치면 (값, 종류)를 돌려준다.

    문장 속에 섞인 숫자는 잡지 않는다 — 대사 대상은 표·항목 단위 수치이고,
    문장 속 숫자까지 넣으면 우연 일치 노이즈가 폭증한다.
    """
    if text is None:
        return None, None
    t = str(text).strip()
    if not t or t in {"-", "–", "—", "N/A", "n/a"}:
        return None, None
    if _DATE_RE.match(t):
        return None, "date"
    m = _NUM_RE.match(t) or _BARE_RE.match(t)
    if not m:
        return None, None
    unit = m.group(1) or ""
    neg = t.strip().startswith("(") and t.strip().endswith(")")
    body = re.sub(r"[^\d.\-+]", "", t)
    if body in {"", "-", "+", "."}:
        return None, None
    try:
        v = float(body)
    except ValueError:
        return None, None
    if neg:
        v = -abs(v)
    if unit == "%":
        return v, "percent"
    if unit in {"주", "건", "배"}:
        return v, "count"
    return v, "amount"


def as_number(value):
    """엑셀 셀처럼 이미 수치인 값을 정규화한다."""
    if isinstance(value, bool):
        return None, None
    if isinstance(value, (int, float)):
        return float(value), "amount"
    return parse_number(value)


# ---------------------------------------------------------------- PPTX

def extract_pptx(path):
    from pptx import Presentation

    prs = Presentation(path)
    blocks = []
    meta = {
        "slide_w_cm": round(prs.slide_width / EMU_CM, 2),
        "slide_h_cm": round(prs.slide_height / EMU_CM, 2),
        "slides": 0,
    }

    def cm(v):
        return round(v / EMU_CM, 3) if v is not None else None

    def run_style(para):
        name = size = None
        bold = False
        for r in para.runs:
            if r.font.name and name is None:
                name = r.font.name
            if r.font.size and size is None:
                size = r.font.size.pt
            if r.font.bold:
                bold = True
        return name, size, bold

    doc = {"statements": [], "notes": [], "front": [], "meta": {"source": "pptx"}}
    unit_rx = re.compile(r"\(\s*단\s*위\s*[:：]\s*([^)]*)\)")
    for si, slide in enumerate(prs.slides, 1):
        meta["slides"] = si
        note = {"no": "S%d" % si, "title": "", "items": []}
        doc["notes"].append(note)
        slide_unit = ""
        for shp in slide.shapes:
            if shp.has_text_frame:
                m = unit_rx.search(shp.text_frame.text)
                if m:
                    slide_unit = m.group(1).strip()
        try:
            if slide.shapes.title is not None and slide.shapes.title.text.strip():
                note["title"] = slide.shapes.title.text.strip().replace("\n", " ")
        except Exception:
            pass
        shapes = sorted(slide.shapes, key=lambda x: ((x.top or 0), (x.left or 0)))
        for shp in shapes:
            geom = {"left": cm(shp.left), "top": cm(shp.top),
                    "w": cm(shp.width), "h": cm(shp.height)}
            if shp.has_table:
                tb = shp.table
                widths = [cm(c.width) for c in tb.columns]
                heights = [cm(r.height) for r in tb.rows]
                grid = [[c.text for c in row.cells] for row in tb.rows]
                col_hdr = grid[0] if grid else []
                # doc: 첫 행에 숫자가 없으면 헤더
                g2 = []
                for row in grid:
                    g2.append([{"text": t.strip(), "num": parse_number(t)[0], "span": False} for t in row])
                hdr_rows = []
                while g2 and all(c["num"] is None for c in g2[0]) and len(hdr_rows) < 2:
                    hdr_rows.append([c["text"] for c in g2.pop(0)])
                note["items"].append({"kind": "table", "unit": slide_unit, "name": shp.name,
                                      "header": hdr_rows, "rows": g2})
                for ri, row in enumerate(tb.rows):
                    row_hdr = grid[ri][0] if grid[ri] else ""
                    for ci, cell in enumerate(row.cells):
                        text = cell.text
                        name = size = None
                        bold = False
                        align = None
                        for p in cell.text_frame.paragraphs:
                            n, s, b = run_style(p)
                            name = name or n
                            size = size or s
                            bold = bold or b
                            align = align or (str(p.alignment).split(" ")[0]
                                              if p.alignment is not None else None)
                        num, kind = parse_number(text)
                        blocks.append({
                            "loc": "S%d/%s[%d,%d]" % (si, shp.name, ri, ci),
                            "kind": "cell",
                            "text": text,
                            "number": num,
                            "num_kind": kind,
                            "context": {"table": shp.name,
                                        "row": row_hdr if ci else "",
                                        "col": col_hdr[ci] if ci < len(col_hdr) and ri else ""},
                            "style": {"font": name, "size": size, "bold": bold,
                                      "align": align,
                                      "anchor": (str(cell.vertical_anchor).split(" ")[0]
                                                 if cell.vertical_anchor is not None else None),
                                      "col_w_cm": widths[ci] if ci < len(widths) else None,
                                      "row_h_cm": heights[ri] if ri < len(heights) else None,
                                      "margins": [cm(cell.margin_left), cm(cell.margin_right),
                                                  cm(cell.margin_top), cm(cell.margin_bottom)]},
                            "geom": dict(geom, rows=len(heights), cols=len(widths),
                                         row_h_sum=round(sum(h or 0 for h in heights), 3)),
                        })
            elif shp.has_text_frame:
                tf = shp.text_frame
                for pi, p in enumerate(tf.paragraphs):
                    if not p.text.strip():
                        continue
                    if p.text.strip() != note["title"]:
                        note["items"].append({"kind": "para", "text": p.text.strip()})
                    name, size, bold = run_style(p)
                    num, kind = parse_number(p.text)
                    blocks.append({
                        "loc": "S%d/%s p%d" % (si, shp.name, pi),
                        "kind": "para",
                        "text": p.text,
                        "number": num,
                        "num_kind": kind,
                        "context": {"table": "", "row": "", "col": ""},
                        "style": {"font": name, "size": size, "bold": bold,
                                  "align": (str(p.alignment).split(" ")[0]
                                            if p.alignment is not None else None),
                                  "runs": len(p.runs),
                                  "shape_type": str(shp.shape_type).split(" ")[0]},
                        "geom": geom,
                    })
    for n in doc["notes"]:
        if not n["title"]:  # 제목 개체틀이 없으면 맨 위 문단을 제목으로
            first = next((it["text"] for it in n["items"] if it["kind"] == "para"), "")
            n["title"] = first[:60]
    doc["notes"] = [n for n in doc["notes"] if n["items"]]
    return meta, blocks, doc


# ---------------------------------------------------------------- XLSX

def extract_xlsx(path, max_rows=5000):
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    blocks = []
    meta = {"sheets": [ws.title for ws in wb.worksheets]}
    for ws in wb.worksheets:
        nrow = min(ws.max_row or 0, max_rows)
        ncol = min(ws.max_column or 0, 60)
        grid = [[ws.cell(r, c).value for c in range(1, ncol + 1)]
                for r in range(1, nrow + 1)]

        def row_header(ri):
            for v in grid[ri]:
                if isinstance(v, str) and v.strip():
                    return v.strip()[:40]
            return ""

        def col_header(ri, ci):
            for rr in range(ri - 1, max(-1, ri - 12), -1):
                v = grid[rr][ci]
                if isinstance(v, str) and v.strip():
                    return v.strip()[:40]
            return ""

        for ri in range(nrow):
            rh = row_header(ri)
            for ci in range(ncol):
                v = grid[ri][ci]
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                num, kind = as_number(v)
                blocks.append({
                    "loc": "%s!%s%d" % (ws.title, _col_letter(ci + 1), ri + 1),
                    "kind": "cell",
                    "text": str(v),
                    "number": num,
                    "num_kind": kind,
                    "context": {"table": ws.title, "row": rh,
                                "col": col_header(ri, ci) if num is not None else ""},
                    "style": {},
                    "geom": {"sheet": ws.title, "r": ri + 1, "c": ci + 1},
                })
    return meta, blocks


def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- HTM / DOCX / PDF

def _int(v, default=1):
    try:
        n = int(str(v).strip())
        return n if 1 <= n <= 50 else default
    except (TypeError, ValueError):
        return default


def _to_grid(rows):
    """(text, colspan, rowspan) 행 목록을 직사각 그리드로 편다.

    colspan/rowspan 을 무시하면 열이 밀려서 합계 검산이 전부 오탐이 된다.
    병합된 칸의 값은 첫 칸에만 두고 나머지는 빈칸으로 채운다.
    """
    grid = []
    carry = {}          # col -> 남은 행 수 (이번 행보다 아래로 몇 줄 더 덮는가)
    for row in rows:
        out = {}
        fresh = {}
        c = 0

        def place_carry():
            nonlocal c
            while c in carry:
                out[c] = ""
                c += 1

        for text, cs, rs in row:
            place_carry()
            for k in range(cs):
                out[c + k] = text if k == 0 else ""
                if rs > 1:
                    fresh[c + k] = rs - 1
            c += cs
        place_carry()
        # 이번 행에서 새로 생긴 rowspan 은 다음 행부터 적용한다.
        # 여기서 같이 깎으면 병합이 한 행 일찍 끝나 이후 행이 통째로 밀린다.
        for k in list(carry):
            carry[k] -= 1
            if carry[k] <= 0:
                del carry[k]
        carry.update(fresh)
        width = (max(out) + 1) if out else 0
        grid.append([out.get(i, "") for i in range(width)])
    return grid


def extract_htm(path):
    from html.parser import HTMLParser

    class P(HTMLParser):
        """표는 반드시 <table> 단위로 나눈다.

        HWP·Word 에서 변환된 HTML 은 전체를 감싸는 바깥 table 이 흔하고,
        표를 하나로 합쳐 버리면 합계 검산이 전부 오탐이 된다. 스택으로
        가장 안쪽 table 에 행을 귀속시킨다.
        """

        def __init__(self):
            super().__init__()
            self.tables = []
            self.stack = []
            self.cur_row = None
            self.cur_cell = None
            self.cur_span = (1, 1)
            self.mute = 0
            self.buf = []

        def handle_starttag(self, tag, attrs):
            if tag in ("style", "script"):
                self.mute += 1
            elif tag == "table":
                self.tables.append([])
                self.stack.append(len(self.tables) - 1)
            elif tag == "tr":
                self.cur_row = []
            elif tag in ("td", "th"):
                self.cur_cell = []
                d = dict(attrs)
                self.cur_span = (_int(d.get("colspan"), 1), _int(d.get("rowspan"), 1))
            elif tag in ("p", "br", "div"):
                self._flush_para()

        def handle_endtag(self, tag):
            if tag in ("style", "script"):
                self.mute = max(0, self.mute - 1)
            elif tag == "table":
                if self.stack:
                    self.stack.pop()
            elif tag == "tr":
                if self.cur_row is not None and self.stack:
                    self.tables[self.stack[-1]].append(self.cur_row)
                self.cur_row = None
            elif tag in ("td", "th") and self.cur_cell is not None:
                text = " ".join(self.cur_cell).strip()
                if self.cur_row is not None:
                    self.cur_row.append((text,) + self.cur_span)
                self.cur_cell = None
                self.cur_span = (1, 1)
            elif tag in ("p", "div"):
                self._flush_para()

        def handle_data(self, data):
            if self.mute:
                return
            d = data.replace("\xa0", " ")
            if not d.strip():
                return
            if self.cur_cell is not None:
                self.cur_cell.append(d.strip())
            elif not self.stack:
                self.buf.append(d.strip())

        def _flush_para(self):
            if self.buf:
                paras.append(" ".join(self.buf).strip())
                self.buf = []

    paras = []
    raw = open(path, "rb").read()
    for enc in ("utf-8", "cp949", "euc-kr", "latin-1"):
        try:
            html = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    p = P()
    p.feed(html)
    p._flush_para()

    blocks = []
    kept = 0
    for ti, raw_rows in enumerate(p.tables):
        if not raw_rows:
            continue
        rows = _to_grid(raw_rows)
        # 바깥 레이아웃용 table (셀이 1개뿐인 껍데기)은 표로 세지 않는다
        if max((len(r) for r in rows), default=0) < 2:
            continue
        kept += 1
        tname = "표%d" % ti
        hdr = rows[0]
        for ri, row in enumerate(rows):
            rh = next((c for c in row if c and parse_number(c)[0] is None), "")
            for ci, cell in enumerate(row):
                if not cell.strip():
                    continue
                num, kind = parse_number(cell)
                blocks.append({
                    "loc": "%s[%d,%d]" % (tname, ri, ci), "kind": "cell", "text": cell,
                    "number": num, "num_kind": kind,
                    "context": {"table": tname, "row": rh if ci else "",
                                "col": hdr[ci] if ci < len(hdr) and ri else ""},
                    "style": {}, "geom": {},
                })
    for pi, para in enumerate(paras):
        if not para.strip():
            continue
        num, kind = parse_number(para)
        blocks.append({"loc": "p%d" % pi, "kind": "para", "text": para,
                       "number": num, "num_kind": kind,
                       "context": {"table": "", "row": "", "col": ""},
                       "style": {}, "geom": {}})
    return {"paras": len(paras), "tables": kept}, blocks


def extract_docx(path):
    import docx  # python-docx

    d = docx.Document(path)
    blocks = []
    for pi, p in enumerate(d.paragraphs):
        if not p.text.strip():
            continue
        num, kind = parse_number(p.text)
        blocks.append({"loc": "p%d" % pi, "kind": "para", "text": p.text,
                       "number": num, "num_kind": kind,
                       "context": {"table": "", "row": "", "col": ""},
                       "style": {"style": p.style.name if p.style else None}, "geom": {}})
    for ti, t in enumerate(d.tables):
        grid = [[c.text for c in row.cells] for row in t.rows]
        hdr = grid[0] if grid else []
        for ri, row in enumerate(grid):
            for ci, cell in enumerate(row):
                num, kind = parse_number(cell)
                blocks.append({"loc": "표%d[%d,%d]" % (ti, ri, ci), "kind": "cell",
                               "text": cell, "number": num, "num_kind": kind,
                               "context": {"table": "표%d" % ti,
                                           "row": row[0] if ci else "",
                                           "col": hdr[ci] if ci < len(hdr) and ri else ""},
                               "style": {}, "geom": {}})
    return {"paras": len(d.paragraphs), "tables": len(d.tables)}, blocks


def extract_pdf(path):
    import fitz

    doc = fitz.open(path)
    blocks = []
    for pi in range(doc.page_count):
        for li, line in enumerate(doc[pi].get_text().splitlines()):
            if not line.strip():
                continue
            num, kind = parse_number(line)
            blocks.append({"loc": "p%d/l%d" % (pi + 1, li), "kind": "para",
                           "text": line.strip(), "number": num, "num_kind": kind,
                           "context": {"table": "", "row": "", "col": ""},
                           "style": {}, "geom": {}})
    return {"pages": doc.page_count}, blocks


# ---------------------------------------------------------------- main

EXTRACTORS = {
    ".pptx": ("pptx", extract_pptx, "report"),
    ".xlsx": ("xlsx", extract_xlsx, "workpaper"),
    ".xlsm": ("xlsx", extract_xlsx, "workpaper"),
    ".htm": ("htm", extract_htm, "report"),
    ".html": ("htm", extract_htm, "report"),
    ".docx": ("docx", extract_docx, "report"),
    ".pdf": ("pdf", extract_pdf, "report"),
    ".dsd": ("dsd", None, "report"),   # dsd.py 로 위임 (doc 구조도 함께 냄)
    ".xml": ("dsd", None, "report"),   # DART 공개 원문 XML (document.xml 의 *_00760.xml 등) — dsd.py 공개 모드
}


def extract_one(path, kind_override=None):
    ext = os.path.splitext(path)[1].lower()
    if ext not in EXTRACTORS:
        raise SystemExit("지원하지 않는 형식: %s" % ext)
    typ, fn, default_kind = EXTRACTORS[ext]
    doc = None
    if typ == "dsd":
        from dsd import extract_dsd
        meta, blocks, doc = extract_dsd(path)
    elif typ == "pptx":
        meta, blocks, doc = fn(path)
    else:
        meta, blocks = fn(path)
    for i, b in enumerate(blocks):
        b["id"] = i
    out = {
        "source": os.path.abspath(path),
        "name": os.path.basename(path),
        "kind": kind_override or default_kind,
        "type": typ,
        "meta": meta,
        "blocks": blocks,
    }
    if doc is not None:
        out["doc"] = doc
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--as", dest="kind", choices=["report", "workpaper"])
    a = ap.parse_args()

    docs = [extract_one(f, a.kind) for f in a.files]
    payload = docs[0] if len(docs) == 1 else {"docs": docs}
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)

    for d in docs:
        nums = sum(1 for b in d["blocks"] if b["number"] is not None)
        print("%s  [%s/%s]  blocks %d  numbers %d"
              % (d["name"], d["type"], d["kind"], len(d["blocks"]), nums))
    print("-> %s" % a.out)


if __name__ == "__main__":
    main()
