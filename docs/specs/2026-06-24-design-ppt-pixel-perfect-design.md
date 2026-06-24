# design-ppt — Claude Design HTML → 픽셀 그대로 .pptx (설계 문서)

- 작성일: 2026-06-24
- 작성: IT전략팀
- 상태: 설계 확정 대기 (사용자 검토 중)
- 대상 저장소: `airforce1226/aps-doc-skills` (마켓플레이스)

## 1. 배경 / 목적

Claude Design(`claude.ai/design`)으로 만든 사내 기술 보고서
("사내 표준 백엔드 아키텍처와 테스트 LDAP 구축 보고 v1.0")의 **시각 디자인을
재사용 가능한 템플릿으로 굳히고**, 같은 디자인의 새 보고서를 **로컬 스킬/커맨드로
생성해 최종 산출물을 .pptx로 내보내는 것**이 목표다.

원본은 슬라이드 한 장 = `<section ... style="width:1920px;height:1080px;...">`
형태의 자체 완결형 HTML이며, CSS 그라데이션·SVG 로고·다양한 레이아웃을 사용한다.

### 확정된 의사결정 (사용자 합의)

| 항목 | 결정 |
|------|------|
| 재사용 방식 | 로컬 **Skill + Command** (기존 `/ppt-report` 계열과 동일 패턴) |
| 최종 포맷 | **.pptx** (HTML은 최종물이 아닌 디자인 매개체) |
| 변환 충실도 | **픽셀 그대로 (이미지 풀블리드)** — 편집가능성보다 디자인 보존 우선 |
| 패키징 | 기존 `ppt-report`와 **분리된 새 플러그인 `design-ppt`** |
| 작성 인터페이스 | 경직된 JSON 아님 — **Claude이 섹션 스니펫을 조립·내용 채움** |
| 폰트 | **Pretendard 시스템 설치** 후 family 이름으로 참조 (base64 임베드 제거) |

## 2. 핵심 원리

- **디자인은 HTML/CSS가 소스, .pptx는 그 렌더 결과.** 슬라이드는 1920×1080 PNG를
  16:9 슬라이드에 풀블리드로 박는다. PowerPoint에서 텍스트는 편집 불가(이미지)지만,
  **원문 텍스트는 발표자 노트(speaker notes)로 이관**해 검색성을 일부 회복한다.
- **결정적(스크립트) / 창의적(Claude) 분리**: 렌더 + pptx 조립은 스크립트가
  결정적으로 수행하고, 어떤 슬라이드 아키타입을 골라 무슨 내용을 채울지는 Claude이
  SKILL.md 지시에 따라 수행한다.
- **팀 규칙은 코드/스킬에 내장**: 저자 = IT전략팀, 파일명 규칙(공백·언더스코어 금지·
  `vN.0` 접미사), 한국어 노트 `lang=ko-KR`.

## 3. 추출한 디자인 시스템 (design-tokens)

원본 `<section>` 인라인 스타일에서 추출:

- **캔버스**: 1920 × 1080 px 고정, 패딩 96–130px
- **컬러**
  - 네이비(표지/강조 배경): `#0b1b3a`, 텍스트 `#ffffff`
  - 라이트(본문 배경): `#f5f7fa`, 텍스트 `#0b1b3a`
  - 액센트 블루: `#0b3fd1` (+ 표지 radial-gradient `rgba(11,63,209,.38)`)
  - APS 로고 그라데이션: `#BED600 → #2BA6CB` (SVG `linearGradient`)
- **타이포**: Pretendard (weight 500/600/700/800), 폴백 `Malgun Gothic`,`맑은 고딕`,sans-serif
- **슬라이드 아키타입 12종** (원본 `data-label`):
  표지 · 배경 · 기술스택 · 폴더구조 · core모듈 · 도메인패턴 · 설계철학 ·
  신규도메인 · 신규프로젝트 · 테스트LDAP · 적용가이드 · 마무리
- **메타**: 각 section은 `data-label`(슬라이드 식별)·`data-speaker-notes`(발표 노트) 보유

## 4. 산출물 구조 (새 플러그인)

```
plugins/design-ppt/
  .claude-plugin/plugin.json
  commands/design-ppt.md                 # /design-ppt 커맨드
  skills/creating-design-ppt/
    SKILL.md                             # 워크플로우 + 팀규칙 + 섹션 조립 가이드
    assets/
      design-tokens.md                   # §3의 디자인 시스템 문서
      base.css                           # 섹션 공통 래퍼/유틸 스타일
      sections/                          # 슬라이드 아키타입 스니펫(내용=placeholder)
        00-cover.html
        01-section-body.html
        02-arch-diagram.html
        ... (원본 12종에서 일반화한 재사용 단위로 정리)
    scripts/
      build_design_ppt.py                # HTML → 슬라이드별 PNG → .pptx 조립
```

> 참고: 스니펫은 원본 12장을 1:1 복제하지 않고, **재사용 가능한 레이아웃 단위로
> 일반화**한다(예: "표지", "섹션 본문", "다이어그램", "표/비교", "마무리"). 원본의
> 구체 다이어그램(아키텍처 레이어 등)은 예시 스니펫으로 포함하되 내용은 placeholder.

## 5. 작성/생성 워크플로우 (SKILL.md가 지시)

1. **내용 수집**: 보고서 종류·독자(임원/실무) 확인, 사용자 인터뷰 또는 기존 초안 읽기.
   미상 값은 지어내지 않고 `(미정 — 추후 확정)`로 표기.
2. **덱 조립**: `assets/sections/`의 스니펫을 골라 복제해 `deck.html` 한 파일을 구성.
   각 section의 placeholder를 실제 내용으로 채우고 `data-speaker-notes`도 작성.
   공통 스타일은 `base.css`, 디자인 토큰은 `design-tokens.md` 준수.
3. **빌드**: `python scripts/build_design_ppt.py deck.html "<제목> v1.0.pptx"`
4. **검증**: python-pptx로 재오픈해 슬라이드 수·노트·`core_properties.author=="IT전략팀"` 확인.
5. **완료 보고**: 파일 경로와 슬라이드 수를 보고.

## 6. 빌드 파이프라인 (build_design_ppt.py)

입력: 채워진 `deck.html` (N개 `<section>`), 출력 경로.

1. `deck.html`을 파싱해 `<section>` 단위로 분리. 각 섹션을 **1920×1080 단일 페이지
   HTML**로 래핑(공용 `<head>`/`base.css`/폰트 참조 포함).
2. 각 페이지를 **헤드리스 Chrome**으로 PNG 캡처(새 pip 의존성 없음):
   ```
   chrome --headless=new --disable-gpu --hide-scrollbars \
          --force-device-scale-factor=1 --window-size=1920,1080 \
          --screenshot="slide_NN.png" "file:///.../slide_NN.html"
   ```
   - Chrome 경로 자동 탐지(`Program Files\Google\Chrome\Application\chrome.exe`),
     없으면 Edge(`msedge.exe`)로 폴백.
3. **python-pptx 조립**: 16:9(13.333"×7.5") 빈 슬라이드에 각 PNG를 (0,0)부터
   슬라이드 전체 크기로 `add_picture` 풀블리드 배치.
4. **발표자 노트**: 각 section `data-speaker-notes` → 대응 슬라이드 notes_slide에
   텍스트로 기록(`lang=ko-KR`).
5. **메타**: `core_properties.author = last_modified_by = "IT전략팀"`. 저장.

### 의존성
- **있음**: Chrome(또는 Edge), Python 3.12, python-pptx 1.0.2
- **설치 필요**: Pretendard 폰트(시스템 설치) — 렌더 충실도용. 미설치 시 Malgun Gothic
  폴백으로 동작은 하나 자간/굵기가 원본과 달라질 수 있음.
- **pip 추가 의존성 없음** (Playwright 미사용 — 빌드 단계 비가시 렌더는 정상).

## 7. 팀 규칙 적용 (내장)

| 규칙 | 적용 위치 |
|------|-----------|
| 저자 = IT전략팀 | `build_design_ppt.py`가 `core_properties.author` 하드코딩 |
| 파일명: 공백·언더스코어 금지·`vN.0` 접미사 끝 | SKILL.md 지시 + 출력 경로는 호출자가 보장 |
| 맞춤법 squiggle 없음 | 슬라이드는 이미지라 해당 없음; 노트엔 `lang=ko-KR` 부여 |
| 미상 데이터 | 지어내지 않고 `(미정 — 추후 확정)` 표기 |

## 8. 범위 밖 (YAGNI)

- 네이티브 편집가능 텍스트 슬라이드(=Approach A) — 이번엔 안 함.
- `.docx` 원본 문서 생성 — 이번 산출물은 .pptx 단일.
- Claude Design 웹 연동/임포트 자동화.
- 임의 `.dc.html` 역(逆)변환기 — 본 스킬은 *자체 템플릿*으로 생성하는 경로만.

## 9. 오픈 이슈 / 리스크

- **폰트 충실도**: Pretendard 설치 여부에 결과가 좌우됨 → 빌드 스크립트가 설치 여부를
  점검해 경고하도록.
- **Chrome 헤드리스 렌더 일관성**: `--force-device-scale-factor=1`로 1:1 픽셀 보장,
  폰트 로딩 완료 대기 필요(필요시 `--virtual-time-budget` 또는 짧은 지연).
- **스니펫 일반화 범위**: 원본 12종을 몇 개의 재사용 단위로 묶을지는 구현 시
  스니펫별 마크업을 보며 확정.
