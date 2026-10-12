# -*- coding: utf-8 -*-
"""보고서 수치를 조서·자체 합산에 전수 대사한다.

사용법
    python tie_out.py 01_report.json 01_wp.json [01_wp2.json ...] -o 02_tieout.json

찾는 근거는 네 종류다.
    exact        조서에 같은 값이 있는 셀
    subtotal     조서 같은 열/행의 연속 구간 합이 그 값
    internal     보고서 자체의 행합·열합·총계
    derived      보고서 수치끼리의 산술 파생 (A/B=비율, A-B=차액, A+B)

어디서도 근거가 안 나온 수치는 unmatched 로 남는다. 이게 지적사항 후보다.
보고서가 조용히 넘어간 숫자를 사람이 놓치지 않게 하는 것이 이 스크립트의 목적이다.
"""
import argparse
import json
from collections import defaultdict

import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

EPS = 0.51          # 원 단위 반올림 허용
PCT_EPS = 0.005     # 백분율 소수 둘째자리 반올림 허용
MAX_RUN = 400       # 연속구간 합 탐색 최대 길이
MIN_ABS = 100       # 이 미만 소액은 우연 일치가 많아 exact 매칭에서 제외


def load(path):
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    return d.get("docs", [d])


def close(a, b, eps=EPS):
    return abs(a - b) <= eps


# ---------------------------------------------------------------- 조서 인덱스

class WpIndex:
    def __init__(self, docs):
        self.exact = defaultdict(list)
        self.cols = defaultdict(list)   # (sheet, col) -> [(row, value, loc)]
        self.rows = defaultdict(list)   # (sheet, row) -> [(col, value, loc)]
        self.n = 0
        for d in docs:
            for b in d["blocks"]:
                v = b.get("number")
                if v is None:
                    continue
                self.n += 1
                self.exact[round(v, 2)].append((d["name"], b["loc"], b["context"]))
                g = b.get("geom") or {}
                sh, r, c = g.get("sheet"), g.get("r"), g.get("c")
                if sh is not None and r and c:
                    self.cols[(d["name"], sh, c)].append((r, v, b["loc"]))
                    self.rows[(d["name"], sh, r)].append((c, v, b["loc"]))
        for k in self.cols:
            self.cols[k].sort()
        for k in self.rows:
            self.rows[k].sort()

    def find_exact(self, v):
        hits = self.exact.get(round(v, 2), [])
        if not hits and abs(v) >= MIN_ABS:
            for cand in (round(v), round(v, 1)):
                hits = self.exact.get(cand, [])
                if hits:
                    break
        return hits

    def find_runs(self, v, axis="cols"):
        """연속 구간 합이 v 인 곳을 찾는다. 누적합 + 해시로 O(n)."""
        out = []
        table = self.cols if axis == "cols" else self.rows
        for key, items in table.items():
            if len(items) < 2:
                continue
            seen = {0.0: -1}
            acc = 0.0
            for i, (_, val, _) in enumerate(items):
                acc += val
                target = acc - v
                for cand in (round(target, 2), round(target)):
                    if cand in seen:
                        j = seen[cand]
                        if i - j <= MAX_RUN and i - j >= 2:
                            out.append({
                                "file": key[0], "sheet": key[1],
                                "axis": "열" if axis == "cols" else "행",
                                "index": key[2],
                                "from": items[j + 1][2], "to": items[i][2],
                                "count": i - j,
                            })
                        break
                seen.setdefault(round(acc, 2), i)
                if len(out) >= 6:
                    return out
        return out


# ---------------------------------------------------------------- 보고서 자체 합산

_CACHE = {}   # id(obj) -> 파생 인덱스. 수치 하나마다 표 전체를 다시 훑지 않게 한다 (조서 없는 감사보고서는 수치 1,400개)


def report_tables(doc):
    """pptx/docx 표를 (table_name, {(r,c): block}) 로 재구성."""
    key = ("tables", id(doc))
    if key in _CACHE:
        return _CACHE[key]
    tabs = defaultdict(dict)
    for b in doc["blocks"]:
        if b["kind"] != "cell":
            continue
        loc = b["loc"]
        if "[" not in loc or "]" not in loc:
            continue
        head, idx = loc.rsplit("[", 1)
        try:
            r, c = (int(x) for x in idx.rstrip("]").split(","))
        except ValueError:
            continue
        tabs[head][(r, c)] = b
    _CACHE[key] = tabs
    return tabs


TOTAL_WORDS = ("합계", "합 계", "총계", "총 계", "소계", "소 계", "계")
DASHES = {"-", "–", "—", "", "N/A", "n/a"}


def is_total_label(text):
    t = (text or "").replace(" ", "").strip()
    if not t:
        return False
    return any(t == w.replace(" ", "") or t.endswith(w.replace(" ", ""))
               for w in TOTAL_WORDS)


def _term(cell):
    """합산 항으로서의 값. '-'·빈칸은 0으로 보되 연속성은 끊지 않는다."""
    if cell is None:
        return 0.0
    if (cell["text"] or "").strip() in DASHES:
        return 0.0
    if cell["number"] is None or cell["num_kind"] != "amount":
        return None
    return cell["number"]


def internal_footing(doc):
    """표별 합계를 검산한다.

    ① 헤더가 '합계/계/소계/총계'인 행·열은 반드시 검산하고, 안 맞으면 fail 로 남긴다.
    ② 그 외에도 어떤 셀이 같은 행·열 연속 구간의 합과 일치하면 검증 근거로 기록한다.
    """
    checks = []
    for name, cells in report_tables(doc).items():
        if not cells:
            continue
        maxr = max(r for r, _ in cells)
        maxc = max(c for _, c in cells)
        col_hdr = {c: (cells[(0, c)]["text"] if (0, c) in cells else "") for c in range(maxc + 1)}
        row_hdr = {r: (cells[(r, 0)]["text"] if (r, 0) in cells else "") for r in range(maxr + 1)}

        def best_run(get_cell, lo, hi, target):
            """[lo, hi) 안에서 합이 target 인 가장 긴 연속 구간. 없으면 None."""
            best = None
            for s in range(lo, hi):
                acc = 0.0
                ok_end = None
                for e in range(s, hi):
                    v = _term(get_cell(e))
                    if v is None:
                        break
                    acc += v
                    if close(acc, target) and e > s:
                        ok_end = e
                if ok_end is not None and (best is None or ok_end - s > best[1] - best[0]):
                    best = (s, ok_end)
            return best

        def diagnose(get_cell, lo, hi, target):
            """합계가 안 맞을 때, 진단 가능한 형태로 원인을 좁힌다.

            합계 바로 앞에서 끝나는 연속 구간만 본다. 멀리 떨어진 구간이
            우연히 가깝다고 그걸 보여주면 진단이 아니라 소음이다.

            항 하나의 부호를 뒤집으면 맞는 경우는 거의 항상 '음수 항목을
            괄호 없이 양수로 표시'한 것이다. 그 사실을 cause 로 돌려준다.
            """
            vals = []
            for e in range(hi - 1, lo - 1, -1):
                v = _term(get_cell(e))
                if v is None:
                    break
                vals.insert(0, (e, v))
            if len(vals) < 2:
                return None

            best = None
            for s in range(len(vals) - 1):
                run = vals[s:]
                acc = sum(v for _, v in run)
                gap = abs(acc - target)
                if best is None or gap < best["gap"]:
                    best = {"gap": gap, "from": run[0][0], "to": run[-1][0],
                            "sum": acc, "n": len(run), "cause": ""}
                for i, (idx, v) in enumerate(run):
                    if v and close(acc - 2 * v, target):
                        return {"gap": 0.0, "from": run[0][0], "to": run[-1][0],
                                "sum": target, "n": len(run),
                                "cause": "%s번 항목 %s 을 음수로 보면 일치 — 괄호(음수) 표기 누락으로 보인다"
                                         % (idx, format(int(abs(v)), ","))}
            return best

        # ① 합계 열 (행 방향 합산)
        for c in range(1, maxc + 1):
            if not is_total_label(col_hdr.get(c)):
                continue
            for r in range(1, maxr + 1):
                tot = cells.get((r, c))
                if not tot or tot["number"] is None or tot["num_kind"] != "amount":
                    continue
                run = best_run(lambda x: cells.get((r, x)), 1, c, tot["number"])
                if run:
                    checks.append({"table": name, "type": "행 합계", "at": "[%d,%d]" % (r, c),
                                   "label": "%s / %s" % (row_hdr.get(r, ""), col_hdr.get(c, "")),
                                   "expected": tot["number"], "actual": tot["number"], "diff": 0.0,
                                   "ok": True, "terms": run[1] - run[0] + 1,
                                   "range": "열%d~%d" % run})
                else:
                    dg = diagnose(lambda x: cells.get((r, x)), 1, c, tot["number"])
                    if dg is None:
                        continue
                    checks.append({"table": name, "type": "행 합계", "at": "[%d,%d]" % (r, c),
                                   "label": "%s / %s" % (row_hdr.get(r, ""), col_hdr.get(c, "")),
                                   "expected": dg["sum"], "actual": tot["number"],
                                   "diff": round(tot["number"] - dg["sum"], 2),
                                   "ok": False, "terms": dg["n"], "cause": dg["cause"],
                                   "range": "열%d~%d" % (dg["from"], dg["to"])})

        # ① 합계 행 (열 방향 합산)
        for r in range(1, maxr + 1):
            if not is_total_label(row_hdr.get(r)):
                continue
            for c in range(1, maxc + 1):
                tot = cells.get((r, c))
                if not tot or tot["number"] is None or tot["num_kind"] != "amount":
                    continue
                run = best_run(lambda x: cells.get((x, c)), 1, r, tot["number"])
                if run:
                    checks.append({"table": name, "type": "열 합계", "at": "[%d,%d]" % (r, c),
                                   "label": "%s / %s" % (row_hdr.get(r, ""), col_hdr.get(c, "")),
                                   "expected": tot["number"], "actual": tot["number"], "diff": 0.0,
                                   "ok": True, "terms": run[1] - run[0] + 1,
                                   "range": "행%d~%d" % run})
                else:
                    dg = diagnose(lambda x: cells.get((x, c)), 1, r, tot["number"])
                    if dg is None:
                        continue
                    checks.append({"table": name, "type": "열 합계", "at": "[%d,%d]" % (r, c),
                                   "label": "%s / %s" % (row_hdr.get(r, ""), col_hdr.get(c, "")),
                                   "expected": dg["sum"], "actual": tot["number"],
                                   "diff": round(tot["number"] - dg["sum"], 2),
                                   "ok": False, "terms": dg["n"], "cause": dg["cause"],
                                   "range": "행%d~%d" % (dg["from"], dg["to"])})
    return checks


def _table_lines(doc):
    """표별 행·열의 (앞 항목 합, 마지막 값) 목록. internal_sum_for 가 값마다 다시 만들지 않게 한 번만."""
    key = ("lines", id(doc))
    if key in _CACHE:
        return _CACHE[key]
    lines = []
    for name, cells in report_tables(doc).items():
        maxr = max((r for r, _ in cells), default=-1)
        maxc = max((c for _, c in cells), default=-1)

        def amounts(coords):
            return [cells[k]["number"] for k in coords
                    if k in cells and cells[k]["number"] is not None and cells[k]["num_kind"] == "amount"]

        for r in range(maxr + 1):
            vals = amounts([(r, c) for c in range(maxc + 1)])
            if len(vals) >= 3:
                lines.append((name, "행", r, sum(vals[:-1]), vals[-1], len(vals) - 1))
        for c in range(maxc + 1):
            vals = amounts([(r, c) for r in range(maxr + 1)])
            if len(vals) >= 3:
                lines.append((name, "열", c, sum(vals[:-1]), vals[-1], len(vals) - 1))
    _CACHE[key] = lines
    return lines


def internal_sum_for(value, doc):
    """이 값이 보고서 내 다른 수치들의 합으로 설명되는가.

    행은 마지막 칸이 그 값(= 행 합계)일 때만, 열은 앞 항목 합이 그 값이면 인정한다.
    """
    out = []
    for name, axis, idx, head_sum, last, terms in _table_lines(doc):
        if not close(head_sum, value):
            continue
        if axis == "행" and not close(last, value):
            continue
        out.append({"table": name, "axis": axis, "index": idx, "terms": terms})
    return out[:4]


KEY_LABELS = ("잔여", "초과", "미달", "차액", "이자", "수령", "해지", "단가", "주식수",
              "납입", "지분", "소진", "환급", "미사용", "잔액", "기한", "만기")


def wp_uncovered(wdocs, report_values):
    """조서에는 있는데 보고서에 반영되지 않은 주요 수치.

    보고서가 조용히 넘어간 숫자를 잡는 역방향 커버리지다. 이번 사례에서
    '잔여액 -623,697,832'·'해지 수령액'·'주식수'가 전부 여기서 나왔다.

    요약 성격 블록만 본다 — 라벨(문자) 옆에 값이 붙어 있고, 그 행의
    채워진 칸이 8개 이하인 곳. 원장·명세 같은 데이터 행은 제외된다.
    """
    seen = set()
    out = []
    for d in wdocs:
        rows = defaultdict(list)
        for b in d["blocks"]:
            g = b.get("geom") or {}
            if g.get("sheet") is None:
                continue
            rows[(g["sheet"], g["r"])].append(b)
        for (sheet, r), items in rows.items():
            if len(items) > 8:
                continue
            items.sort(key=lambda x: x["geom"]["c"])
            for i, b in enumerate(items):
                v = b.get("number")
                if v is None or abs(v) < 1000:
                    continue
                label = ""
                for j in range(i - 1, -1, -1):
                    if items[j].get("number") is None:
                        label = items[j]["text"].strip()[:40]
                        break
                if not label:
                    continue
                if any(close(v, rv) for rv in report_values):
                    continue
                key = (label, round(v, 2))
                if key in seen:
                    continue
                seen.add(key)
                out.append({"file": d["name"], "loc": b["loc"], "label": label,
                            "value": v, "text": b["text"],
                            "keyword": any(k in label for k in KEY_LABELS)})
    out.sort(key=lambda x: (not x["keyword"], -abs(x["value"])))
    return out[:40]


def _big_index(nums):
    """1,000 넘는 수치 목록과 반올림값 → 위치 색인. 차액·합 파생을 쌍 전수조사(n²) 대신 색인으로 찾는다."""
    key = ("big", id(nums))
    if key in _CACHE:
        return _CACHE[key]
    big = [x for x in nums if x[1] is not None and abs(x[1]) > 1000]
    by_key = defaultdict(list)
    for j, (_, b, _) in enumerate(big):
        by_key[round(b)].append(j)
    _CACHE[key] = (big, by_key)
    return big, by_key


def derived_for(value, kind, nums):
    """비율·차액·합 파생을 찾는다. nums = [(loc, value, text)]"""
    out = []
    big, by_key = _big_index(nums)
    if kind == "percent":
        tgt = value / 100.0
        for la, a, _ in big:
            for lb, b, _ in big:
                if a is b or not b:
                    continue
                if abs(a / b - tgt) <= PCT_EPS / 100 * 100:
                    if abs(a / b - tgt) <= 0.00005:
                        out.append({"expr": "%s ÷ %s" % (la, lb), "value": round(a / b * 100, 4)})
        return out[:4]
    for i, (la, a, _) in enumerate(big):
        # a−b=v, b−a=v, a+b=v 를 만족하는 b 는 세 목표값 ±EPS 안에 있다 → 색인으로 후보만 본다 (결과·순서는 전수조사와 같다)
        cand = set()
        for t in (a - value, a + value, value - a):
            k = round(t)
            for kk in (k - 1, k, k + 1):
                cand.update(j for j in by_key.get(kk, ()) if j > i)
        for j in sorted(cand):
            lb, b, _ = big[j]
            if close(a - b, value) and abs(value) > MIN_ABS:
                out.append({"expr": "%s − %s" % (la, lb)})
            elif close(b - a, value) and abs(value) > MIN_ABS:
                out.append({"expr": "%s − %s" % (lb, la)})
            elif close(a + b, value):
                out.append({"expr": "%s + %s" % (la, lb)})
            if len(out) >= 4:
                return out
    return out


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("workpapers", nargs="*")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    rdocs = load(a.report)
    rdoc = rdocs[0]
    wdocs = [d for p in a.workpapers for d in load(p)]
    idx = WpIndex(wdocs)

    def is_noise(b):
        """페이지 번호·한 자리 수치는 대사 대상이 아니다."""
        loc = b["loc"]
        if "슬라이드 번호" in loc or "Slide Number" in loc:
            return True
        v = b["number"]
        return b["num_kind"] == "amount" and abs(v) < MIN_ABS

    targets = [b for b in rdoc["blocks"] if b["number"] is not None and not is_noise(b)]
    nums = [(b["loc"], b["number"], b["text"]) for b in targets]

    matched, unmatched = [], []
    for b in targets:
        v = b["number"]
        ev = []
        for f, loc, ctx in idx.find_exact(v)[:5]:
            ev.append({"type": "exact", "where": "%s %s" % (f, loc),
                       "detail": " / ".join(x for x in (ctx.get("row"), ctx.get("col")) if x)})
        if not ev and b["num_kind"] == "amount" and abs(v) >= MIN_ABS:
            for run in idx.find_runs(v, "cols") + idx.find_runs(v, "rows"):
                ev.append({"type": "subtotal",
                           "where": "%s %s!%s~%s" % (run["file"], run["sheet"], run["from"], run["to"]),
                           "detail": "%s 연속 %d건 합" % (run["axis"], run["count"])})
        for s in internal_sum_for(v, rdoc):
            ev.append({"type": "internal",
                       "where": "%s %s%d" % (s["table"], s["axis"], s["index"]),
                       "detail": "보고서 내부 %d개 항목 합" % s["terms"]})
        if not ev:
            for d in derived_for(v, b["num_kind"], nums):
                ev.append({"type": "derived", "where": d["expr"],
                           "detail": "보고서 수치 파생"})

        rec = {"loc": b["loc"], "text": b["text"], "value": v, "kind": b["num_kind"],
               "context": b["context"], "evidence": ev[:6]}
        (matched if ev else unmatched).append(rec)

    footing = internal_footing(rdoc)
    bad = [f for f in footing if not f["ok"]]
    uncovered = wp_uncovered(wdocs, [b["number"] for b in targets])

    out = {
        "report": rdoc["name"],
        "workpapers": [d["name"] for d in wdocs],
        "stats": {"report_numbers": len(nums), "wp_numbers": idx.n,
                  "matched": len(matched), "unmatched": len(unmatched),
                  "footing_checks": len(footing), "footing_fail": len(bad),
                  "wp_uncovered": len(uncovered)},
        "matched": matched,
        "unmatched": unmatched,
        "footing": footing,
        "wp_uncovered": uncovered,
    }
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    print("보고서 수치 %d개 / 조서 수치 %d개" % (len(nums), idx.n))
    print("  근거 확인 %d  ·  미매칭 %d" % (len(matched), len(unmatched)))
    print("  합계 검산 %d건 중 불일치 %d건" % (len(footing), len(bad)))
    for u in unmatched:
        print("  [미매칭] %-22s %-18s %s" % (u["loc"], u["text"][:18],
                                             " / ".join(x for x in (u["context"].get("row"),
                                                                    u["context"].get("col")) if x)))
    for f_ in bad:
        head = "  [합계불일치] %s %s %s (%s)" % (f_["table"], f_["type"], f_["at"],
                                                f_["label"].strip(" /"))
        if f_.get("cause"):
            print("%s 기재 %s\n      └ %s"
                  % (head, format(int(f_["actual"]), ","), f_["cause"]))
        else:
            print("%s 기재 %s ≠ %s 합 %s"
                  % (head, format(int(f_["actual"]), ","), f_["range"],
                     format(int(f_["expected"]), ",")))
    if uncovered:
        print("  조서에만 있고 보고서에 없는 주요 수치 %d건 (상위 10):" % len(uncovered))
        for u in uncovered[:10]:
            print("    %-22s %18s   %s" % (u["label"][:22], u["text"][:18], u["loc"]))
    print("-> %s" % a.out)


if __name__ == "__main__":
    main()
