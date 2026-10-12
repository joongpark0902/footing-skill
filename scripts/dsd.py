# -*- coding: utf-8 -*-
"""DART 편집기 dsd(zip: contents.xml + meta.xml) 파서.

extract.py 가 호출한다. 두 가지를 돌려준다.

1. blocks — 다른 포맷과 같은 정규화 블록 (tie_out / format_audit 용)
2. doc    — build_workbook.py 가 쓰는 구조화 문서
   {"statements": [{"key":"BS","title":"재무상태표","unit":"원","annual":bool,
                    "periods":["제 48(당) 기","제 47(전) 기"],
                    "rows":[{"label":str,"level":int|None,"acode":str,
                             "vals":[float|None,...],"dup":[bool,...]}]}],
    "notes": [{"no":"3","title":"재고자산",
               "items":[{"kind":"para","text":str} |
                        {"kind":"table","unit":"천원","header":[[str,...],...],
                         "rows":[[{"text":str,"num":float|None},...],...]}]}],
    "front": [str,...]   # 감사보고서 본문 문단
   }

dsd 재무제표 표(ACLASS=FINANCE)는 TE 셀에 ALEVEL(계층)·ADELIM(열)·ACODE(계정코드)가 붙고,
기간 하나가 '세부·합계' 두 열로 나뉜다. 여기서는 두 열을 기간당 하나로 합친다
(둘 다 값이 있으면 합계열을 쓰고 dup=True).

DART 에 공시된 원문 XML(OpenDART document.xml 의 *_00760.xml 등)도 같은 태그 구조지만
ALEVEL·ACLASS=FINANCE 가 빠져 있다. 그 경우 '재무상태표' 같은 제목표 뒤의 표를 재무제표로 보고,
계층은 dsd 관례(로마숫자 0 · 번호 1 · 괄호번호 2 · 들여쓰기 문단 1)로 추정한다.
.xml 을 그대로 넘겨도 되고 zip(contents.xml) 이어도 된다.
"""
import re
import xml.etree.ElementTree as ET
import zipfile

from extract import parse_number

STMT_KEYS = [
    ("BS", r"재무상태표|대차대조표"),
    ("PL", r"포괄손익계산서|손익계산서"),
    ("MC", r"제조원가명세서"),
    ("CE", r"자본변동표"),
    ("CF", r"현금흐름표"),
    ("RE", r"이익잉여금처분계산서|결손금처리계산서"),
]

_NOTE_START = re.compile(r"^\s*(\d{1,2})\s*\.\s*(\S.*)$")
_UNIT_RE = re.compile(r"\(\s*단\s*위\s*[:：]\s*([^)]*)\)")


def _txt(el):
    t = "".join(el.itertext())
    return t.replace("&cr;", "\n").replace(" ", "\n").replace("　", " ")


def _clean(s):
    return re.sub(r"[ \t]+", " ", s).strip()


def _lines(el):
    return [_clean(x) for x in _txt(el).split("\n")]


def _norm_stmt_title(s):
    return re.sub(r"\s+", "", s)


def _stmt_key(title):
    t = _norm_stmt_title(title)
    for k, pat in STMT_KEYS:
        if re.search(pat, t):
            return k
    return None


# ---------------------------------------------------------------- 표 그리드

def _grid(table):
    """TR/TH/TD/TE 를 rowspan·colspan 을 펼친 2차원 셀 리스트로.

    각 셀: {"text", "num", "tag", "attrs", "span": bool(펼쳐진 복제 셀)}
    """
    rows = []
    pending = {}  # col -> (remaining_rows, cell)
    for tr in table.iter("TR"):
        cells = [c for c in tr if c.tag in ("TH", "TD", "TE", "TU")]
        out = []
        col = 0
        ci = 0
        while ci < len(cells) or any(pending.get(c, (0,))[0] > 0 for c in range(col, col + 1)):
            if col in pending and pending[col][0] > 0:
                rem, cell = pending[col]
                out.append(dict(cell, span=True))
                pending[col] = (rem - 1, cell)
                col += 1
                continue
            if ci >= len(cells):
                break
            c = cells[ci]
            ci += 1
            text = _clean(_txt(c))
            num, kind = parse_number(text)
            cell = {"text": text, "num": num, "num_kind": kind, "tag": c.tag,
                    "attrs": dict(c.attrib), "span": False}
            cs = int(c.get("COLSPAN", "1") or 1)
            rs = int(c.get("ROWSPAN", "1") or 1)
            for k in range(cs):
                out.append(dict(cell, span=(k > 0)))
                if rs > 1:
                    pending[col] = (rs - 1, cell)
                col += 1
        # 남은 pending 열 처리 (행 끝에 rowspan 셀이 있는 경우)
        while col in pending and pending[col][0] > 0:
            rem, cell = pending[col]
            out.append(dict(cell, span=True))
            pending[col] = (rem - 1, cell)
            col += 1
        rows.append(out)
    return rows


def _is_header_row(row, public=False):
    if not row:
        return False
    if all(c["tag"] == "TH" for c in row):
        return True
    # 공개 XML 주석 표는 TH 대신 배경색(USERMARK=BC0X…) 입힌 TD 로 머리행을 만든다
    return public and all((c["attrs"].get("USERMARK") or "").startswith("BC0X") for c in row)


# ---------------------------------------------------------------- 재무제표

_ROMAN = re.compile(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+\s*[.．]")
_NUMBERED = re.compile(r"^\d{1,2}\s*[.．]")
_PAREN = re.compile(r"^\(\d{1,2}\)")
_CE_BAL = re.compile(r"기초|기말|\d{4}\s*[.년]")
_NOTE_COL = re.compile(r"^주\s*석$")


def _infer_level(key, label, attrs):
    """공개 XML(ALEVEL 없음)용 계층 추정 — dsd 관례(로마숫자 0 · 번호 1 · 괄호번호 2)에 맞춘다."""
    s = label.strip()
    if key.startswith("CE"):                       # 자본변동표: 잔액행 0, 변동행 1
        return 0 if _CE_BAL.search(s) else 1
    if _ROMAN.match(s):
        return 0
    if _PAREN.match(s):
        return 2
    if _NUMBERED.match(s):
        return 1
    if (attrs.get("USERMARK") or "").startswith("P"):   # P15·P16 = 들여쓰기 문단
        return 1
    return 0


def _parse_statement(title, tables, unit, public=False):
    key = _stmt_key(title)
    periods = []
    rows = []
    for tb in tables:
        if not public and tb.get("ACLASS") != "FINANCE":
            continue
        g = _grid(tb)
        hdr = [r for r in g if _is_header_row(r)]
        body = [r for r in g if not _is_header_row(r)]
        # 기간 열 묶음: 헤더 행(셀 2개 이상인 첫 행)에서 과목 이후 셀을 span 기준으로 묶는다.
        # '주석' 열은 기간이 아니므로 뺀다.
        groups = []
        if hdr:
            h0 = next((r for r in hdr if sum(1 for c in r if not c["span"]) >= 2), hdr[0])
            i = 1
            while i < len(h0):
                j = i
                while j + 1 < len(h0) and h0[j + 1]["span"]:
                    j += 1
                if not _NOTE_COL.match(h0[i]["text"]):
                    groups.append((h0[i]["text"], list(range(i, j + 1))))
                i = j + 1
        if not periods:
            periods = [g_[0] for g_ in groups]
        for r in body:
            if not r:
                continue
            lab = r[0]
            level = lab["attrs"].get("ALEVEL")
            if level not in (None, ""):
                level = int(level)
            else:
                level = _infer_level(key, lab["text"], lab["attrs"]) if public else None
            vals, dup = [], []
            for _, cols in groups:
                nums = [r[c]["num"] for c in cols if c < len(r) and r[c]["num"] is not None]
                vals.append(nums[-1] if nums else None)
                dup.append(len(nums) > 1)
            if not groups:  # 헤더가 없으면 열 전부를 각각 기간으로
                vals = [c["num"] for c in r[1:]]
                dup = [False] * len(vals)
            rows.append({"label": lab["text"], "level": level,
                         "acode": lab["attrs"].get("ACODE", ""), "vals": vals, "dup": dup})
    if public and key.startswith("PL"):
        # K-IFRS 손익계산서는 단계식이라 'Ⅳ.영업이익' 뒤에 들여쓴 금융수익·금융비용이 와도 그 구성항목이 아니다
        parent = ""
        for r in rows:
            if r["level"] == 0:
                parent = r["label"]
            elif r["level"] == 1 and re.search(r"이익|손실", parent):
                r["level"] = 0
    annual = not re.search(r"반기|분기|중간", " ".join(periods) + title)
    return {"key": key, "title": title, "unit": unit, "annual": annual,
            "periods": periods, "rows": rows}


def _find_unit(tables):
    for tb in tables:
        for tu in tb.iter("TU"):
            if tu.get("AUNIT") in ("WON", "THOUSANDWON", "MILLIONWON"):
                m = _UNIT_RE.search(_txt(tu))
                return _clean(m.group(1)) if m else _clean(_txt(tu))
        m = _UNIT_RE.search(_txt(tb))
        if m:
            return _clean(m.group(1))
    return ""


# ---------------------------------------------------------------- 주석

def _note_table(tb, unit, public=False):
    g = _grid(tb)
    header, body = [], []
    for r in g:
        # 표 첫 행이 '(단위: 천원)' 한 칸뿐이면 단위행 (공개 XML 은 단위를 표 안에 넣는다)
        if public and not header and not body and len({c["text"] for c in r}) == 1 \
                and _UNIT_RE.search(r[0]["text"]):
            unit = _clean(_UNIT_RE.search(r[0]["text"]).group(1))
            continue
        if _is_header_row(r, public) and not body:
            header.append([c["text"] for c in r])
        else:
            body.append([{"text": c["text"], "num": c["num"], "span": c["span"]} for c in r])
    return {"kind": "table", "unit": unit, "header": header, "rows": body}


_TITLE_BREAK = re.compile(r"(당기말과|당기와|당기 ?중|당사는|당사의|당사가|회사는|회사의|회사가|다음은|보고기간|"
                          r"재무제표|전기말과|\(\d+\)|\d{1,2}\.\d)")


def _split_title(rest):
    """공개 XML 은 주석 제목과 첫 문장 사이 줄바꿈이 사라져 붙어 있다 → 본문 첫 어절 앞에서 자른다."""
    m = _TITLE_BREAK.search(rest, 1)
    if m and m.start() <= 40:
        return _clean(rest[:m.start()]), _clean(rest[m.start():])
    return _clean(rest), ""


def _parse_notes(section, public=False):
    notes = []
    cur = None
    unit_next = ""
    pre = []  # 첫 주석 이전 문단 (기간 표 등)

    def start(no, rest):
        """새 주석 시작. 제목 뒤에 붙어 온 본문이 있으면 돌려준다(호출자가 다시 줄 처리)."""
        nonlocal cur
        title, body = _split_title(rest) if public else (_clean(rest), "")
        cur = {"no": no, "title": title, "items": []}
        notes.append(cur)
        return body

    def is_start(m):
        if not m or (public and re.match(r"\d", m.group(2))):
            return False                  # '4.1.3 유동성 위험' 같은 하위 번호는 주석 시작이 아니다
        return cur is None or int(m.group(1)) >= int(cur["no"]) + 1

    for ch in list(section):
        if ch.tag == "TITLE":
            continue
        if ch.tag == "P":
            txt = _txt(ch)
            if public:   # '7' + 줄바꿈 + '. 영업부문' 처럼 번호와 점이 갈라진 제목
                txt = re.sub(r"(?m)^\s*(\d{1,2})\s*\n\s*\.\s*", r"\1. ", txt)
            pending = [l for l in (_clean(x) for x in txt.split("\n")) if l]
            first = True
            while pending:
                l = pending.pop(0)
                m = _NOTE_START.match(l) if (first or public) else None   # dsd 는 문단 첫 줄만 제목 후보
                first = False
                if is_start(m):
                    body = start(m.group(1), m.group(2))
                    if body:
                        pending.insert(0, body)
                    continue
                if public and cur is not None:
                    # 앞 주석의 마지막 문장에 다음 주석 번호가 이어 붙은 경우 ("…않았습니다.21. 자본잉여금…")
                    # 문장 끝(. 또는 ) ) 바로 뒤의 번호, 점 뒤는 글자여야 한다 ("주석 4.1.2", "20.9%" 는 제외)
                    m = re.search(r"(?:(?<=\.)|(?<=\.\s)|(?<=\)\s))(%d)\s*\.\s*(?=[가-힣A-Za-z(])"
                                  % (int(cur["no"]) + 1), l)
                    if m and m.start() > 0:
                        cur["items"].append({"kind": "para", "text": _clean(l[:m.start()])})
                        body = start(m.group(1), l[m.end():])
                        if body:
                            pending.insert(0, body)
                        continue
                (cur["items"] if cur else pre).append({"kind": "para", "text": l})
        elif ch.tag == "TABLE":
            txt = _clean(_txt(ch))
            m = _UNIT_RE.search(txt)
            if m and len(list(ch.iter("TR"))) == 1 and len(txt) < 40:
                unit_next = _clean(m.group(1))
                continue
            item = _note_table(ch, unit_next, public)
            unit_next = ""
            (cur["items"] if cur else pre).append(item)
        elif ch.tag == "TABLE-GROUP":
            for tb in ch.iter("TABLE"):
                item = _note_table(tb, unit_next, public)
                (cur["items"] if cur else pre).append(item)
            unit_next = ""
        # PGBRK 등은 무시
    return notes, pre


# ---------------------------------------------------------------- 진입점

def _iter_tables(el):
    """el 의 직계 TABLE(TABLE-GROUP 안 포함). 하위 SECTION 은 내려가지 않는다."""
    for ch in el:
        if ch.tag.startswith("SECTION-"):
            continue
        if ch.tag == "TABLE":
            yield ch
        elif ch.tag == "TABLE-GROUP":
            yield from ch.iter("TABLE")


def _public_statements(body):
    """공개 XML: '(첨부)재무제표' 섹션에서 '재무상태표' 같은 제목표 뒤에 오는 헤더 있는 표들을 재무제표로."""
    out = []
    for sec in body.iter():
        if not sec.tag.startswith("SECTION-"):
            continue
        t = sec.find("TITLE")
        if t is None or "재무제표" not in _norm_stmt_title(_txt(t)):
            continue
        cur = None
        for tb in _iter_tables(sec):
            g = _grid(tb)
            if not g or not g[0]:
                continue
            has_th = any(_is_header_row(r) for r in g)
            first = g[0][0]["text"].split("\n")[0]
            key = _stmt_key(first) if (not has_th and len(g) <= 6) else None
            if key:
                m = _UNIT_RE.search(_txt(tb))
                cur = {"title": first, "tables": [], "unit": _clean(m.group(1)) if m else ""}
                out.append(cur)
            elif cur is not None and has_th:
                cur["tables"].append(tb)
        if out:
            break
    return [(c["title"], c["tables"], c["unit"]) for c in out if c["tables"]]


def _add_statement(doc, st):
    if not st["rows"]:
        return
    # 같은 key 가 둘이면 (별도·연결 등) 뒤에 번호
    if any(s["key"] == st["key"] for s in doc["statements"]):
        st["key"] = st["key"] + str(sum(1 for s in doc["statements"] if s["key"].startswith(st["key"])) + 1)
    doc["statements"].append(st)


def parse_dsd(path):
    if zipfile.is_zipfile(path):
        xml = zipfile.ZipFile(path).read("contents.xml")
    else:                                         # DART 공개 원문 XML 그대로
        with open(path, "rb") as f:
            xml = f.read()
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        # 공개 원문 XML 은 본문에 <당기말> 같은 꺾쇠 표기·날것 & 가 섞여 잘 안 열린다 → 보정 후 재시도
        from xml_fix import fix_xml
        root = ET.fromstring(fix_xml(xml.decode("utf-8", "replace")))
    body = root.find("BODY")
    if body is None:
        body = root
    doc = {"statements": [], "notes": [], "front": [], "meta": {}}
    cn = root.find(".//COMPANY-NAME")
    doc["meta"]["company"] = _clean(_txt(cn)) if cn is not None else ""
    dn = root.find(".//DOCUMENT-NAME")
    doc["meta"]["docname"] = _clean(_txt(dn)) if dn is not None else ""

    # 1) 재무제표: TITLE 다음에 오는 TABLE 들을 다음 TITLE 전까지 모은다 (dsd)
    for parent in body.iter():
        kids = list(parent)
        for i, el in enumerate(kids):
            if el.tag != "TITLE":
                continue
            key = _stmt_key(_clean(_txt(el)))
            if not key:
                continue
            tables = []
            for nx in kids[i + 1:]:
                if nx.tag == "TITLE":
                    break
                if nx.tag == "TABLE":
                    tables.append(nx)
                elif nx.tag == "TABLE-GROUP":
                    tables.extend(nx.iter("TABLE"))
            if not any(t.get("ACLASS") == "FINANCE" for t in tables):
                continue
            _add_statement(doc, _parse_statement(_clean(_txt(el)), tables, _find_unit(tables)))
    # 1') 공개 XML: FINANCE 표가 하나도 없으면 제목표 기준으로 찾는다
    public = not doc["statements"]
    if public:
        for title, tables, unit in _public_statements(body):
            _add_statement(doc, _parse_statement(title, tables, unit or _find_unit(tables), public=True))
        doc["meta"]["source"] = "public_xml"

    # 2) 주석
    for sec in body.iter():
        if not sec.tag.startswith("SECTION-"):
            continue
        t = sec.find("TITLE")
        if t is not None and _clean(_txt(t)) in ("주석", "주  석", "재무제표에 대한 주석"):
            notes, pre = _parse_notes(sec, public=public)
            doc["notes"] = notes
            doc["notes_pre"] = pre
            break

    # 3) 감사보고서 본문·외부감사 실시내용 문단
    for sec in body.iter():
        if not sec.tag.startswith("SECTION-"):
            continue
        t = sec.find("TITLE")
        tt = _clean(_txt(t)) if t is not None else ""
        if re.search(r"감사인의|검토인의|외부감사 실시내용", tt):
            for p in sec.iter("P"):
                for l in _lines(p):
                    if l:
                        doc["front"].append(l)
    return doc


def blocks_from_doc(doc):
    blocks = []
    for st in doc["statements"]:
        k = st["key"]
        for ri, r in enumerate(st["rows"]):
            blocks.append({"loc": "%s[%d,0]" % (k, ri), "kind": "cell", "text": r["label"],
                           "number": None, "num_kind": None,
                           "context": {"table": k, "row": "", "col": ""}, "style": {}, "geom": {}})
            for ci, v in enumerate(r["vals"]):
                if v is None:
                    continue
                blocks.append({"loc": "%s[%d,%d]" % (k, ri, ci + 1), "kind": "cell",
                               "text": "{:,.0f}".format(v), "number": v, "num_kind": "amount",
                               "context": {"table": k, "row": r["label"],
                                           "col": st["periods"][ci] if ci < len(st["periods"]) else ""},
                               "style": {}, "geom": {}})
    for n in doc["notes"]:
        ti = 0
        for it in n["items"]:
            if it["kind"] == "para":
                num, kind = parse_number(it["text"])
                blocks.append({"loc": "주%s/p" % n["no"], "kind": "para", "text": it["text"],
                               "number": num, "num_kind": kind,
                               "context": {"table": "", "row": "", "col": ""}, "style": {}, "geom": {}})
            else:
                ti += 1
                hdr = it["header"][-1] if it["header"] else []
                # 머리행도 셀 블록으로 낸다 — tie_out 의 표 합계 검산은 0행을 머리행으로 보고 1행부터 더한다
                for hi, hrow in enumerate(it["header"]):
                    for ci, txt in enumerate(hrow):
                        blocks.append({"loc": "주%s/표%d[%d,%d]" % (n["no"], ti, hi, ci), "kind": "cell",
                                       "text": txt, "number": None, "num_kind": None,
                                       "context": {"table": "주%s/표%d" % (n["no"], ti), "row": "", "col": ""},
                                       "style": {"unit": it.get("unit", "")}, "geom": {}})
                off = len(it["header"])
                for ri, row in enumerate(it["rows"], start=off):
                    for ci, c in enumerate(row):
                        if c.get("span"):
                            continue
                        blocks.append({"loc": "주%s/표%d[%d,%d]" % (n["no"], ti, ri, ci), "kind": "cell",
                                       "text": c["text"], "number": c["num"],
                                       "num_kind": "amount" if c["num"] is not None else None,
                                       "context": {"table": "주%s/표%d" % (n["no"], ti),
                                                   "row": row[0]["text"] if ci else "",
                                                   "col": hdr[ci] if ci < len(hdr) else ""},
                                       "style": {"unit": it.get("unit", "")}, "geom": {}})
    for l in doc["front"]:
        blocks.append({"loc": "본문", "kind": "para", "text": l, "number": None, "num_kind": None,
                       "context": {"table": "", "row": "", "col": ""}, "style": {}, "geom": {}})
    return blocks


def extract_dsd(path):
    doc = parse_dsd(path)
    blocks = blocks_from_doc(doc)
    meta = {"statements": [s["key"] for s in doc["statements"]],
            "notes": len(doc["notes"]), "company": doc["meta"].get("company", "")}
    return meta, blocks, doc
