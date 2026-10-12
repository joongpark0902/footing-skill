# -*- coding: utf-8 -*-
"""보고서의 표 형식·레이아웃·공백·표기 일관성을 기계적으로 감사한다.

사용법
    python format_audit.py 01_report.json -o 02_format.json [--font malgun.ttf]

사람 눈으로는 잘 안 보이지만 인쇄·다른 PC에서 터지는 것들을 잡는다.
    width      셀 텍스트 실측 폭 > 가용 폭 (또는 여유 0.15cm 미만)
    font       같은 표·같은 행 안에서 글꼴·크기 혼재
    align      같은 열 안에서 정렬 혼재
    geom       0.01~0.15cm 단위로 어긋난 도형 좌표, 표 하단선 불일치
    blank      합계행의 빈 칸, 합계행 강조 누락
    space      이중공백·앞뒤공백·탭
    placeholder  ~~~ · TBD · TODO · xx · ??? 같은 미완성 표시
    term       같은 말이 띄어쓰기만 다르게 쓰인 곳 (발생 여부 / 발생여부)
"""
import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TIGHT = 0.15          # 여유가 이보다 작으면 줄바꿈 위험으로 본다
GEOM_TOL = 0.16       # 이 이내로 어긋나면 "맞추려다 어긋난" 것으로 본다
PLACEHOLDER_RE = re.compile(r"(~{3,}|_{3,}|\bTBD\b|\bTODO\b|X{2,}|x{2,}(?![a-z])|\?{3,}|\bxx\b|＊{3,})")
HANGUL_RUN = re.compile(r"[가-힣]+(?:\s+[가-힣]+)*")

_FONT_CACHE = {}


def font_path(bold=False):
    base = r"C:\Windows\Fonts"
    return os.path.join(base, "malgunbd.ttf" if bold else "malgun.ttf")


def text_cm(text, pt, bold=False):
    """맑은 고딕 실측 폭(cm). PIL 이 없으면 근사치로 떨어진다."""
    if not text:
        return 0.0
    try:
        from PIL import ImageFont
        key = (bold, int(pt * 100))
        f = _FONT_CACHE.get(key)
        if f is None:
            f = ImageFont.truetype(font_path(bold), int(pt * 1000))
            _FONT_CACHE[key] = f
        return f.getlength(text) / 1000 / 72 * 2.54
    except Exception:
        w = 0.0
        for ch in text:
            w += pt * (1.0 if ord(ch) > 0x1100 else 0.52)
        return w / 72 * 2.54


def add(out, kind, sev, loc, msg, detail=""):
    out.append({"kind": kind, "sev": sev, "loc": loc, "msg": msg, "detail": detail})


# ---------------------------------------------------------------- 개별 점검

def check_width(blocks, out):
    for b in blocks:
        st = b.get("style") or {}
        text = (b.get("text") or "").strip()
        if not text:
            continue
        pt = st.get("size") or 10
        if b["kind"] == "cell":
            cw = st.get("col_w_cm")
            if not cw:
                continue
            m = st.get("margins") or [0.1, 0.1, 0, 0]
            avail = cw - (m[0] or 0) - (m[1] or 0)
        else:
            g = b.get("geom") or {}
            if not g.get("w") or len(text) > 40:
                continue
            avail = g["w"] - 0.5
        need = max(text_cm(line, pt, st.get("bold")) for line in text.split("\n"))
        if need > avail:
            add(out, "width", "B", b["loc"],
                "텍스트가 폭을 넘겨 줄바꿈된다", "필요 %.2fcm > 가용 %.2fcm — %r" % (need, avail, text[:30]))
        elif avail - need < TIGHT:
            add(out, "width", "B", b["loc"],
                "폭 여유 %.2fcm — 글꼴 대체 시 줄바꿈 위험" % (avail - need),
                "필요 %.2fcm / 가용 %.2fcm — %r" % (need, avail, text[:30]))


def check_font_align(blocks, out):
    tables = defaultdict(list)
    for b in blocks:
        if b["kind"] == "cell" and b["context"].get("table"):
            tables[b["loc"].rsplit("[", 1)[0]].append(b)

    for tname, cells in tables.items():
        rows = defaultdict(list)
        cols = defaultdict(list)
        for b in cells:
            try:
                r, c = (int(x) for x in b["loc"].rsplit("[", 1)[1].rstrip("]").split(","))
            except ValueError:
                continue
            if (b.get("text") or "").strip():
                rows[r].append((c, b))
                cols[c].append((r, b))

        fonts = Counter((b["style"].get("font") or "?") for b in cells
                        if (b.get("text") or "").strip())
        if len([f for f in fonts if f != "?"]) > 1:
            add(out, "font", "B", tname,
                "표 안에서 글꼴 지정이 갈린다", " / ".join("%s %d칸" % (k, v) for k, v in fonts.most_common()))

        for r, items in rows.items():
            fs = {(b["style"].get("font") or "?") for _, b in items}
            if len(fs - {"?"}) > 1:
                add(out, "font", "B", "%s 행%d" % (tname, r),
                    "같은 행 안에서 글꼴이 갈린다" + (" (헤더행)" if r == 0 else ""),
                    " / ".join(sorted(fs)))
            szs = {b["style"].get("size") for _, b in items if b["style"].get("size")}
            if len(szs) > 1:
                add(out, "font", "B", "%s 행%d" % (tname, r),
                    "같은 행 안에서 글자 크기가 갈린다", " / ".join("%gpt" % s for s in sorted(szs)))

        for c, items in cols.items():
            data = [(r, b) for r, b in items if r > 0]
            als = {b["style"].get("align") for _, b in data if b["style"].get("align")}
            if len(als) > 1:
                add(out, "align", "B", "%s 열%d" % (tname, c),
                    "같은 열 안에서 정렬이 갈린다", " / ".join(sorted(als)))


def check_total_row(blocks, out):
    from tie_out import is_total_label  # 같은 판정 기준을 공유한다

    tables = defaultdict(dict)
    for b in blocks:
        if b["kind"] != "cell":
            continue
        loc = b["loc"]
        if "[" not in loc:
            continue
        head, idx = loc.rsplit("[", 1)
        try:
            r, c = (int(x) for x in idx.rstrip("]").split(","))
        except ValueError:
            continue
        tables[head][(r, c)] = b

    for tname, cells in tables.items():
        if not cells:
            continue
        maxr = max(r for r, _ in cells)
        maxc = max(c for _, c in cells)
        has_style = any((b.get("style") or {}).get("size") for b in cells.values())
        total_rows = [r for r in range(1, maxr + 1)
                      if cells.get((r, 0)) and is_total_label(cells[(r, 0)].get("text"))]
        # 소계가 여러 번 나오는 계층 표(재무상태표 등)는 빈 칸이 정상 구조다
        layered = len(total_rows) > 1
        for r in total_rows:
            if layered:
                continue
            blanks = [c for c in range(1, maxc + 1)
                      if (r, c) not in cells or not (cells[(r, c)].get("text") or "").strip()]
            # 위 행에 수치가 있는 열만 본다. 이름·날짜 같은 문자 열이
            # 합계행에서 비는 것은 정상이다.
            filled_above = [c for c in blanks
                            if any((rr, c) in cells and cells[(rr, c)].get("number") is not None
                                   for rr in range(1, r))]
            if filled_above:
                add(out, "blank", "B", "%s 행%d" % (tname, r),
                    "합계행에 빈 칸이 있다 (위 행에는 값이 있음)",
                    "열 " + ", ".join(str(c) for c in filled_above))
            bolds = [cells[(r, c)]["style"].get("bold") for c in range(maxc + 1)
                     if (r, c) in cells and (cells[(r, c)].get("text") or "").strip()]
            if has_style and bolds and not any(bolds):
                add(out, "blank", "B", "%s 행%d" % (tname, r),
                    "합계행에 굵게·음영 등 강조가 없어 데이터행과 구분되지 않는다", "")


def check_geometry(blocks, out):
    shapes = {}
    for b in blocks:
        g = b.get("geom") or {}
        if g.get("left") is None:
            continue
        key = b["loc"].split(" p")[0].rsplit("[", 1)[0]
        shapes[key] = g

    for axis in ("left", "top"):
        vals = [round(g[axis], 2) for g in shapes.values() if g.get(axis) is not None]
        if len(vals) < 3:
            continue
        counts = Counter(vals)
        for modal, n in counts.most_common():
            if n < 2:
                continue
            offenders = sorted({v for v in vals if 0 < abs(v - modal) <= GEOM_TOL})
            if offenders:
                where = [k for k, g in shapes.items()
                         if g.get(axis) is not None and round(g[axis], 2) in offenders]
                add(out, "geom", "B", ", ".join(sorted(where)[:6]),
                    "%s 좌표가 %.2fcm 기준선에서 미세하게 어긋난다" % (axis, modal),
                    "기준 %.2f / 어긋난 값 " % modal + ", ".join("%.2f" % v for v in offenders))
            break

    tops = defaultdict(list)
    for k, g in shapes.items():
        if g.get("rows") and g.get("top") is not None:
            slide = k.split("/", 1)[0]
            tops[(slide, round(g["top"], 2))].append((k, g))
    for (_slide, top), items in tops.items():
        if len(items) < 2:
            continue
        bottoms = {k: round(top + (g.get("h") or 0), 2) for k, g in items}
        if len(set(bottoms.values())) > 1:
            add(out, "geom", "B", ", ".join(sorted(bottoms)),
                "나란한 표의 하단선이 맞지 않는다",
                " / ".join("%s %.2fcm" % (k, v) for k, v in sorted(bottoms.items())))


def is_letter_spaced(text):
    """`자  본  총  계` · `합     계` 처럼 자간을 공백으로 벌린 표 라벨인가.

    HWP·Word 에서 변환된 문서에 흔하다. 이걸 이중공백으로 잡으면
    한 보고서에서 수백 건이 쏟아져 진짜 지적이 묻힌다.
    긴 낱말이 하나도 없으면 문장이 아니라 벌려 쓴 라벨로 본다.
    """
    toks = text.split()
    return bool(toks) and all(len(t) <= 3 for t in toks)


def check_text(blocks, out):
    for b in blocks:
        t = b.get("text") or ""
        if not t.strip():
            continue
        flat = t.replace("\n", " ").replace("\t", " ")
        if "  " in flat and not is_letter_spaced(flat):
            add(out, "space", "C", b["loc"], "이중 공백", repr(t[:60]))
        for line in t.split("\n"):
            if line and line != line.rstrip():
                add(out, "space", "C", b["loc"], "줄 끝 공백", repr(line[-25:]))
            if line and line != line.lstrip() and b["kind"] == "para":
                add(out, "space", "C", b["loc"], "줄 앞 공백", repr(line[:25]))
        if "\t" in t:
            add(out, "space", "C", b["loc"], "탭 문자로 정렬 — 폭이 바뀌면 깨진다", repr(t[:60]))
        m = PLACEHOLDER_RE.search(t)
        if m:
            add(out, "placeholder", "S", b["loc"],
                "미완성 표시가 남아 있다: %r" % m.group(0), repr(t[:80]))


def check_terms(blocks, out):
    """같은 말이 띄어쓰기만 다르게 쓰인 곳을 찾는다."""
    surfaces = defaultdict(set)
    where = defaultdict(set)
    for b in blocks:
        t = b.get("text") or ""
        for run in HANGUL_RUN.findall(t):
            words = run.split()
            for i in range(len(words)):
                for j in (1, 2, 3):
                    if i + j > len(words):
                        break
                    surf = " ".join(words[i:i + j])
                    key = surf.replace(" ", "")
                    if len(key) < 4:
                        continue
                    surfaces[key].add(surf)
                    where[key].add(b["loc"])
    for key, forms in sorted(surfaces.items()):
        if len(forms) > 1:
            add(out, "term", "C", ", ".join(sorted(where[key])[:4]),
                "같은 말의 띄어쓰기가 갈린다", " / ".join(sorted(forms)))


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    with open(a.report, encoding="utf-8") as f:
        doc = json.load(f)
    docs = doc.get("docs", [doc])
    blocks = [b for d in docs for b in d["blocks"]]

    out = []
    check_width(blocks, out)
    check_font_align(blocks, out)
    check_total_row(blocks, out)
    check_geometry(blocks, out)
    check_text(blocks, out)
    check_terms(blocks, out)

    order = {"S": 0, "A": 1, "B": 2, "C": 3}
    out.sort(key=lambda x: (order.get(x["sev"], 9), x["kind"], x["loc"]))
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump({"report": docs[0]["name"], "count": len(out), "items": out},
                  f, ensure_ascii=False, indent=1)

    by = Counter(x["kind"] for x in out)
    print("서식 감사 %d건  —  %s" % (len(out), ", ".join("%s %d" % kv for kv in by.most_common())))
    for x in out:
        print("  [%s/%s] %-26s %s  %s" % (x["sev"], x["kind"], x["loc"][:26], x["msg"], x["detail"][:70]))
    print("-> %s" % a.out)


if __name__ == "__main__":
    main()
