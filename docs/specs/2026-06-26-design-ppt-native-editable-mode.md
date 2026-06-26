# design-ppt 네이티브 편집 경로 (`--mode native`) — 설계서

> 출처 사양: `작업지시서_APS_PPT_플러그인.md` (IT전략팀 · v1.0).
> 작성: IT전략팀 · 언어: 한국어 · 일자: 2026-06-26.
>
> 본 문서는 작업지시서를 **별도 Office Add-in 제품으로 구축하지 않고**, 그 핵심 원칙
> ("이미지로 굽지 않고 모든 요소를 편집 가능한 네이티브 개체로")을 기존 `design-ppt`
> 스킬에 **추가 출력 경로**로 흡수하는 설계다.

---

## 1. 배경 · 목표

### 1.1 현재 상태
- `design-ppt` 스킬(`creating-design-ppt`)은 APS Brand Presentation System(섹션 00–13,
  `assets/design-tokens.md`, `scripts/build_design_ppt.py`)을 갖춘 성숙한 자산이다.
- 현재 빌드는 헤드리스 Chrome `--screenshot`으로 각 `<section>`(1920×1080)을 PNG로 구워
  슬라이드에 full-bleed로 박는다 → **시각 100% 충실하지만 PowerPoint에서 텍스트 편집 불가**
  (원문은 발표자 노트에 보존).

### 1.2 목표
1. 같은 `deck.html`에서, 빌드 모드만 바꾸면 **텍스트·표·도형이 편집 가능한 네이티브 개체**인
   .pptx가 나오게 한다.
2. 기존 이미지 경로는 **무손상 유지**(기본값). 네이티브는 나란히 추가되는 옵션이다.
3. APS 디자인이 플랫·고정 어휘(네이비/페이퍼/블루 단색, 60×4 그라데이션 룰, 단순 박스·표)
   라는 점을 활용해, 픽셀 정확도를 유지하면서 네이티브화한다.

### 1.3 비목표 (Non-Goal)
- Office Add-in(React+TS+Vite, manifest.xml, AppSource) 제품은 **만들지 않는다**.
- 이미지 경로 제거/대체는 하지 않는다(둘 다 유지).
- 새 런타임 의존성(playwright/selenium 등) 도입하지 않는다 — 기존 Chrome subprocess만 사용.
- v1에서 차트·로고 그라데이션 freeform의 완전 네이티브화는 하지 않는다(2차 마일스톤).

---

## 2. 아키텍처 — 방식 A (HTML 단일 소스 + 브라우저 측정 + 선택적 래스터 폴백)

### 2.1 핵심 아이디어
이미지 경로가 이미 Chrome로 `deck.html`을 렌더한다. 네이티브 경로는 **스크린샷 대신
레이아웃을 측정**한다: 브라우저(=CSS 레이아웃 엔진)가 계산한 각 노드의 좌표·스타일을 읽어
그 위치에 python-pptx 네이티브 개체를 배치한다. 좌표를 손으로 계산하지 않으므로 픽셀 정확하다.

```
deck.html ──┬─ --mode image  → Chrome --screenshot → PNG → full-bleed 슬라이드 (현행)
            └─ --mode native → Chrome --dump-dom(측정JS) → layout JSON → python-pptx 개체 (신규)
```

### 2.2 추출 메커니즘 (의존성 추가 없음)
- 기존 `subprocess` + Chrome 패턴 재사용. `capture_slide()`의 `--screenshot`을 네이티브에서는
  **`--dump-dom`**으로 대체한다(동일 플래그: `--headless=new`, `--virtual-time-budget`,
  `--force-device-scale-factor=1`, `--window-size=1920,1080`).
- 섹션 페이지 래퍼(`wrap_section_page`)에 **측정 스크립트**를 추가 주입한다. 로드 시:
  1. DOM을 순회하며 대상 노드를 분류(§3).
  2. 각 노드의 `getBoundingClientRect()`(px) + `getComputedStyle()`(폰트·색·정렬·보더·radius)를
     수집.
  3. 결과 배열을 `JSON.stringify`하여 `<pre id="__layout__">…</pre>`에 기록.
- Python은 `--dump-dom` stdout에서 `<pre id="__layout__">`의 텍스트만 정규식으로 떼어
  `json.loads` → 노드 리스트.

> **주의·검증 포인트**: `--headless=new --dump-dom`가 측정 스크립트 실행 후의 DOM을 반환하는지
> 빌드 셋업 초기에 확인한다(`--virtual-time-budget`로 스크립트 완료 보장). 실패 시 폴백:
> `--remote-debugging-port` + 최소 DevTools 프로토콜 호출(여전히 새 pip 의존성 없음).

### 2.3 단위 환산
- 작업지시서 §3: 캔버스 1920×1080 px ↔ 슬라이드 13.333"×7.5".
- 좌표/크기 → EMU: python-pptx `Emu`/`Cm` 사용. `emu_per_px = SLIDE_W_EMU / 1920`.
- 폰트 pt: computedStyle의 px 폰트값을 그대로 쓰지 않고 작업지시서 환산 `pt = px × 0.54` 적용
  (캔버스 픽셀폰트 → 슬라이드 pt). `native_render.py`에 `px_to_pt`, `px_to_emu` 헬퍼.

---

## 3. 노드 → 네이티브 개체 매핑

| HTML 노드 | 네이티브 개체 | 1차 폴백 |
|---|---|---|
| 텍스트 리프(직접 자식이 텍스트뿐) | `TextBox` — 폰트명·pt·색·굵기·정렬·자간 = computedStyle | — |
| 배경/박스(`background`·`border`·`border-radius`) | `Rectangle` / `RoundRectangle` (fill·line·rounding), 자식보다 뒤 레이어 | — |
| 60×4 액센트 룰·그라데이션 바 | `Rectangle` + **lxml 2-stop 그라데이션**(시작 `#BED600` → 끝 `#2BA6CB`) | 단색 APS Blue `#0b3fd1` |
| `<table>` | 네이티브 표: 헤더행 네이비·white·bold, 짝수행 지브라, 합계행 블루, 숫자셀 우정렬, 음수 텍스트 빨강 | — |
| 로고 SVG / 그라데이션 락업 | `APS` 워드마크 텍스트(800 굵기) (작업지시서 §5.7) | 해당 조각만 소형 래스터 |
| 차트(섹션 09, CSS 막대/도넛) | **v1: 해당 영역 래스터** | (2차) 네이티브 pptx 차트 |
| 매핑 불가 노드 | — | **그 노드 bbox만 래스터 이미지로** 배치 + 경고 로그 |

- **색 스냅**: computedStyle의 rgb 값을 `design-tokens.md` 팔레트 중 최근접색으로 스냅해
  임의색 유입을 막는다(허용오차 내). 스냅 실패 시 원색 유지 + 경고.
- **그라데이션 헬퍼**: python-pptx 고수준 API에 그라데이션이 없으므로 `lxml`로 도형
  `spPr`에 `<a:gradFill>` 주입(작업지시서 §6 OOXML 보강 철학과 동일).
- **자동맞춤 끔**: 텍스트박스는 word_wrap/auto_size를 명시 제어하고 박스 크기를 bbox로 고정
  (작업지시서 §5.1-5).

---

## 4. 신뢰도 브리지 — 아키타입 `data-ppt` 역할 힌트

임의 CSS 추측을 줄이기 위해, 14개 섹션의 핵심 요소에 **가벼운 역할 속성**을 1회 부여한다:

```html
<div data-ppt="box"> … </div>
<span data-ppt="text"> … </span>
<div data-ppt="rule"></div>
<table data-ppt="table"> … </table>
<div data-ppt="raster"> …복합 장식… </div>
<div data-ppt="chart"> …차트… </div>
```

- `data-*`는 **이미지 경로가 무시**하므로 현행 동작에 **0 영향**(순수 추가, 안전).
- 측정 스크립트는 `data-ppt`가 있으면 그 역할을 신뢰하고, 없으면 휴리스틱(텍스트 리프 판정,
  thin-rule 판정 등)으로 분류한다 → 힌트는 정확도를 높이는 보조이지 필수 아님.
- 힌트 부여는 섹션 단위로 점진 적용 가능(표지·간지부터).

---

## 5. 코드 구조 (격리)

```
scripts/
├─ build_design_ppt.py     # 기존. --mode {image|native} 플래그 추가. image 로직 불변.
├─ native_render.py        # ★신규: 측정 JS 문자열, dump-dom 호출, layout JSON 파싱,
│                          #        노드→개체 매핑, lxml 그라데이션/표 헬퍼, 색 스냅, 빌드 리포트.
└─ tests/
   ├─ test_build_design_ppt.py   # 기존
   └─ test_native_render.py      # 신규: 환산·색스냅·매핑 단위 테스트, 표 스타일 규칙
```

- `build_design_ppt.py`:
  - `argparse`(또는 경량 인자 파싱)로 `--mode` 추가. 기본 `image`(하위호환).
  - `mode == native`면 `native_render.build_native(sections, css, out_path)` 위임.
  - 팀 규칙(author=IT전략팀, 파일명 검증, 대외비 배지, noProof 노트)은 **양 모드 공통 경로**로
    유지 — 네이티브 슬라이드에도 발표자 노트·배지·메타데이터 동일 적용.
- `native_render.py`는 이미지 경로 함수를 호출하지 않는다(역방향 의존 금지). 공통 유틸
  (`find_browser`, `split_sections`, `_set_notes`, 파일명/author 처리)은 그대로 재사용.

### 5.1 빌드 리포트
빌드 종료 시 슬라이드별 요약을 출력한다:
```
Slide 01 (cover):  native=7  raster=1(logo)   ok
Slide 09 (charts): native=4  raster=1(chart)  ok
…
Total: 12 slides, 134 native objects, 3 raster fallbacks.
```
→ 편집 가능성과 폴백 위치를 가시화한다(작업지시서 §10 체크리스트 정신).

---

## 6. 문서 (SKILL.md / design-tokens.md)

- `SKILL.md`:
  - "출력 모드" 절 추가: **이미지(기본)** = 픽셀 완벽·편집 불가 / **네이티브(`--mode native`)**
    = 편집 가능·일부 장식 래스터 폴백. 언제 무엇을 쓰는지 가이드.
  - Quick Reference에 `--mode native` 행, 빌드 리포트 읽는 법 추가.
  - Common Mistakes: "네이티브 모드 결과의 차트/로고는 래스터일 수 있음(의도된 폴백)".
- `design-tokens.md`: 색 스냅 팔레트가 단일 소스임을 명시(이미 토큰이 있으므로 표만 보강).
- 작업지시서 §10 정확성 체크리스트를 스킬의 검증 절로 흡수(네이티브 개체 여부·색 토큰·맑은 고딕·
  표 규칙·대외비 일관·미확인은 `(미정 — 추후 확정)`).

---

## 7. 검증 · 완료 기준 (DoD)

1. `python build_design_ppt.py deck.html "제목 v1.0.pptx" --mode native` 가 성공.
2. python-pptx로 재오픈 시 슬라이드의 텍스트/표/도형이 **picture가 아닌 실제 개체**
   (`shape.has_text_frame`, `shape.has_table`)로 존재.
3. PowerPoint에서 텍스트·색·표 행 편집이 실제로 가능(수동 1회 확인).
4. `core_properties.author == "IT전략팀"`, 대외비 대상 시 전 슬라이드 배지, 노트 `noProof`.
5. 색은 design-tokens 팔레트로 스냅됨(임의색 없음). 폰트 맑은 고딕.
6. 표지·간지·KPI·표 4종이 이미지 모드 대비 레이아웃이 시각적으로 일치(오버레이 육안 비교).
7. `--mode` 미지정 시 기존 이미지 동작과 100% 동일(회귀 없음).
8. `pytest scripts/tests/` 통과.

---

## 8. 마일스톤

| 단계 | 산출 | 완료 기준 |
|---|---|---|
| N1. 측정 파이프라인 | `native_render.py` 측정 JS + dump-dom 파싱 + 환산 헬퍼 | layout JSON이 정확한 bbox/스타일 반환(표지로 검증) |
| N2. 기본 개체 | 텍스트·박스·룰(그라데이션) 매핑 + 색 스냅 | 표지·간지가 네이티브·편집 가능·레이아웃 일치 |
| N3. 표 | 네이티브 표 + 스타일 규칙 | 섹션 07 표가 헤더 네이비·지브라·합계 블루·음수 빨강 |
| N4. 폴백·리포트 | 로고 워드마크/래스터, 차트 래스터, 빌드 리포트 | 차트·로고 슬라이드가 폴백으로 정상 산출, 리포트 출력 |
| N5. data-ppt 힌트·문서 | 14 섹션 힌트 부여 + SKILL.md/토큰 갱신 + 테스트 | 10종 모두 네이티브화, 회귀 없음, pytest 통과 |

---

## 9. 리스크 · 미확인

- **`--headless=new --dump-dom`로 스크립트 실행 후 DOM 확보 가능 여부** → N1 초기에 검증
  (실패 시 remote-debugging-port 폴백, 새 pip 의존성 없음). `TODO(추가 확인 필요)`.
- **python-pptx 그라데이션/표 셀 스타일의 lxml 주입 안정성** → 헬퍼로 캡슐화 + 단위 테스트.
- **차트 네이티브화**는 v1 비목표(데이터 역추론 불가) → 래스터 폴백, 2차 마일스톤에서 검토.
- **휴리스틱 오분류** → `data-ppt` 힌트로 보강, 빌드 리포트로 폴백 위치 가시화.
