---
name: footing
description: 회계·재무 보고서 draft를 근거 조서와 대사(풋팅)해 지적사항을 중요도 순으로 정리한 엑셀을 만든다. 숫자 전수대사·합계검산은 물론 오타·띄어쓰기·표 형식·셀 폭·문맥/비문·조서간 정합성까지 함께 점검한다. 투자금 실사보고서, 반기검토·감사보고서, 손상검사보고서 등 종류를 가리지 않는다. 트리거 — "풋팅", "풋팅해줘", "풋팅작업", "footing", "tie-out", "보고서 대사", "조서 대사", "draft 검증", "보고서 숫자 확인", "보고서랑 조서 맞춰줘", "오타·표형식 체크", 그리고 footing 폴더 경로만 던지는 경우.
---

# 풋팅 — 보고서 draft ↔ 근거 조서 대사

보고서 초안을 발행 전에 붙잡는 작업이다. 산출물은 **`풋팅결과_{보고서명}.xlsx`** 하나다.
두 가지가 한 파일에 들어간다 — ① 지적사항 목록(중요도 순) ② **검증 워크북**(보고서를 옮겨 적고 검증식을 살아 있는 수식으로 심은 재수행 조서, `references/workbook.md`).

## 철칙

1. **보고서 원본을 수정하지 않는다.** 무엇을 고칠지는 실사자가 판단한다. 이 스킬은 지적만 한다.
2. **근거 없는 지적을 쓰지 않는다.** 모든 지적에는 조서 셀 위치·전표번호·재계산식 중 하나가 붙는다.
3. **숫자가 맞는다고 끝이 아니다.** 실제로 중요한 지적은 대개 "숫자 오류"가 아니라 **결론과 근거의 연결**에서 나온다 — 조서에 미결로 남은 확인사항, 보고서가 설명하지 못하는 차액, 조서에만 있고 보고서에 없는 수치.
4. **기계가 판단하지 않는다.** 스크립트는 사실만 뽑고, 승격·기각·중요도는 모델이 정한다.
5. **미매칭은 기본값이 "지적 후보"다.** 근거를 못 찾은 수치를 조용히 넘기지 않는다. 기각하려면 왜 기각인지 쓴다.
6. **보고서에 적힌 합계는 값으로 둔다.** 합계 자리에 `=SUM` 을 넣으면 검증이 항상 참이 된다. 검증셀이 합계와 Σ항목을 비교한다.

## Phase 0 — 입력 확정

작업 시작 시 다음 한 줄을 먼저 출력한다.

```
footing — {보고서명} / 프로파일: {유형} / run: {YYYY-MM-DD-NNN}
```

- 인자가 폴더면 그 안에서 보고서(.dsd/.xml/.pptx/.htm/.docx/.pdf)와 조서(.xlsx/.xlsm)를 찾는다. 인자가 없으면 cwd.
- 후보가 여럿이라 어느 게 보고서인지 모호하면 **묻는다.** 넘겨짚지 않는다.
- 파일명·표지·목차로 유형을 판정하고 `references/profiles/` 에서 프로파일을 고른다. 해당 프로파일이 없으면 `_template.md` 로 진행하고 그 사실을 고지한다.
- 작업폴더는 보고서 폴더 하위 `_footing/{run_id}/`. 최종 엑셀만 보고서 폴더 루트에 놓는다.
- 감사·검토보고서는 **dsd(DART 편집기) 원본**을 받는 게 가장 좋다. 표 계층(ALEVEL)·계정코드가 살아 있어 소계 구조를 기계가 읽는다. **DART 에 공시된 원문 XML**(OpenDART `document.xml` 로 받은 zip 안의 `*_00760.xml` 감사보고서 등)도 같은 파서로 읽는다 — 계층 속성이 없어 로마숫자·번호·들여쓰기로 계층을 추정하고, 재무제표 표는 '재무상태표' 같은 제목표 뒤에서 찾는다. pdf 만 있으면 워크북의 표 재수행은 못 하고 지적사항 목록만 낸다 — 그 사실을 고지한다.

## Phase 1 — 추출

```
python {skill}/scripts/extract.py "{보고서}" -o _footing/{run}/01_report.json
python {skill}/scripts/extract.py "{조서1}" "{조서2}" -o _footing/{run}/01_wp.json
python {skill}/scripts/extract.py "{전기보고서}" -o _footing/{run}/01_prev.json   # 있으면
```

출력 줄의 blocks·numbers 개수가 0이면 파싱 실패다. 그대로 진행하지 말고 원인을 밝힌다.
dsd·pptx 는 `doc`(재무제표·주석 / 슬라이드별 표·문단 구조)도 함께 나온다. 워크북은 이 `doc` 으로 만든다.

## Phase 2 — 기계 감사

```
python {skill}/scripts/tie_out.py _footing/{run}/01_report.json _footing/{run}/01_wp.json -o _footing/{run}/02_tieout.json
python {skill}/scripts/format_audit.py _footing/{run}/01_report.json -o _footing/{run}/02_format.json
python {skill}/scripts/build_workbook.py _footing/{run}/01_report.json -o _footing/{run}/02_workbook.xlsx   # 1차 — 구조 확인용
```

세 스크립트의 **콘솔 출력만 읽고 판단해도 된다.** JSON 전문을 읽는 것은 특정 항목을 파고들 때만.

`tie_out` 이 내놓는 네 가지를 각각 다르게 다룬다.

| 결과 | 뜻 | 처리 |
|---|---|---|
| `matched` | 조서·내부합산·파생에서 근거를 찾음 | 숫자대사 시트 행으로. 근거 유형이 `derived` 뿐이면 조서 근거가 없다는 뜻이니 한 번 더 본다 |
| `unmatched` | 어디서도 근거를 못 찾음 | **지적 후보.** 기각하려면 사유를 쓴다 |
| `footing` `ok:false` | 합계가 안 맞음 | 거의 항상 S급 |
| `wp_uncovered` | 조서에는 있는데 보고서에 없는 수치 | **가장 값진 입력.** 잔여액·초과액·수령액·단가·주식수·납입일이 여기서 나온다 |

`build_workbook` 1차 콘솔은 세 가지를 준다 — `구조 미인식`(산식을 못 정한 소계), `조서 미매칭`(wp 매핑 전이므로 이 단계에선 없음), `차이 상세`(합계·행검증이 안 맞는 셀). 이것이 Phase 3 의 입력이다.

## Phase 3 — 판단 (모델의 일)

`02_tieout.json` · `02_format.json` · 워크북 콘솔 요약 + 보고서 텍스트 전문 + 프로파일을 놓고 지적사항을 확정한다.

기계가 못 하는 것을 여기서 한다.

- **결론 ↔ 조서 정합성** — 조서에 "확인필요·TBD·검토중" 플래그가 남아 있는데 보고서가 "이상 없음"으로 끝나는가. 조서의 미결 건수·금액·비중을 세어 붙인다.
- **주석·설명의 정량 검증** — 보고서가 "~때문이다"라고 쓴 문장을 숫자로 검산한다. 맞아떨어지지 않으면 차액을 계산해 적는다. deal 보고서는 이 검산을 `04_structure.json` 의 `claims` 로 워크북 **서술검산 시트**에 수식으로 남긴다.
- **문맥·비문** — 같은 말 반복, 조사 오용, 인용한 절 제목이 실제 제목과 다른 경우, 용어 혼용(사용목적/사용용도).
- **누락 항목** — 통상 기재하는데 빠진 것(정의·산식·기간·단위·절차·범위).
- **중요도 부여** — `references/severity.md` 기준.
- `format_audit` 의 `term`·`space` 항목은 **전기 서명본 답습 여부**를 따져본다. 답습이면 그 사실을 근거란에 적는다.
- **워크북 구조 정의** — `04_structure.json` 을 쓴다 (`references/workbook.md` 스키마). 구조 미인식 소계의 산식(`rules`), 조서 매핑(`wp`: 파일·시트·라벨열·기간열·부호·합산), 주석↔본문 크로스(`xrefs`), 변동표 열 규칙(`note_rules`), 서술 검산(`claims`). 차이가 뜬 합계는 **먼저 보고서 오류인지 구조 오해인지** 가른다. 구조 오해면 규칙을 고치고, 보고서 오류면 지적으로 올린다.

확정 결과를 `_footing/{run}/03_findings.json` 으로 쓴다. 스키마는 `scripts/build_report.py` 상단 주석에 있다.

작성 규칙:
- `issue` 는 한 문장으로 무엇이 문제인지. `basis` 는 숫자·위치·계산식. `fix` 는 실행 가능한 조치.
- 근거란에 계산이 들어가면 **식을 그대로 쓴다** (`623,697,832 − 292,940,813 = 330,757,019`).
- 검증만 하고 이상 없던 항목은 반드시 `verified` 에 남긴다. 무엇을 봤는지가 조서다.

## Phase 4 — 산출

```
python {skill}/scripts/build_workbook.py _footing/{run}/01_report.json -o "풋팅결과_{보고서명}.xlsx" \
    --findings _footing/{run}/03_findings.json --structure _footing/{run}/04_structure.json [--prev _footing/{run}/01_prev.json]
```

`doc` 이 없는 입력(pdf·docx·htm)이면 대신 `build_report.py _footing/{run}/03_findings.json -o "풋팅결과_{보고서명}.xlsx"`.

산출 후 **엑셀로 실제로 열어 재계산**한다 (COM: `Workbooks.Open` → `CalculateFull` → 오류값 스캔 → 총괄표 값 읽기). 총괄표의 오류 건수가 콘솔의 파이썬 재계산 건수와 같아야 한다. 다르면 수식이 잘못된 것이다.

## Phase 5 — 응답

사용자에게 짧게 넷만 준다.

1. 한 줄 상태 — `대사 {n}건 차이 0 / 지적 {n}건 (S {n} · A {n} · B {n} · C {n})` + 워크북 검증셀 수와 오류 수
2. 숫자 검증 결과 요약 3~5줄
3. **S급 전건**을 한 줄씩
4. 파일 경로

A·B·C급은 나열하지 않는다. 엑셀에 있다.

## 옵션 (인자 끝에 자연어로)

- `유형: 투자금실사|검토감사|손상검사|재무실사` — 프로파일 강제 지정
- `숫자만` — Phase 2 의 `tie_out` 만, 서식·문구 생략
- `서식만` — `format_audit` 만
- `워크북만` — 지적사항 없이 검증 워크북만 (draft 초기 단계에서 구조 확인)
- `재실행` — 기존 `_footing` 폴더 재사용, Phase 3 부터

## 참조

- 검증 워크북 서식·structure.json 스키마 — `references/workbook.md`
- 공통 점검 항목 — `references/checklist.md`
- 중요도 판정 — `references/severity.md`
- 유형별 논점 — `references/profiles/*.md`
