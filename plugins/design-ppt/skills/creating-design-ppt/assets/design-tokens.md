# design-ppt 디자인 토큰 (APS Brand Presentation System)

원본: **"APS 슬라이드 템플릿" (Turn on the APS ON)** 브랜드 프레젠테이션 시스템에서 추출.
이 토큰과 `sections/` 아키타입이 디자인 기준이며, 새 레이아웃도 여기에 맞춘다.

## 캔버스
- 슬라이드: **1920 × 1080 px** 고정 (`<section style="width:1920px;height:1080px;position:relative;...">`)
- 패딩: 표지/간지/마무리(네이비) **120px**, 본문 110px, 요약·WBS·조직·비교 **96px**, 좌우 130px
- 모든 `<section>`은 `position:relative` (보안 배지·워터마크 기준점)

## 컬러
| 용도 | HEX |
|------|-----|
| Navy (표지/간지/마무리 배경, 다크 박스) | `#0b1b3a` |
| APS Blue (액센트/강조/숫자/링크) | `#0b3fd1` |
| Brand Gradient (제목 하단 액센트 바, 로고) | `#BED600` → `#2BA6CB` |
| Paper (본문 배경) | `#f5f7fa` |
| 본문 텍스트 | `#0b1b3a` |
| 보조 텍스트(슬레이트) | `#5b6b85` |
| 옅은 보조/푸터 텍스트 | `#9aa6b8` |
| 보더/구분선/그리드 갭 | `#e1e7f0` |
| Soft 박스 배경(옅은 블루) | `#e7edf8` |
| 네이비 위 라이트블루 라벨 | `#7fa3ff` |
| 보안 분류(대외비) 적색 | `#c0392b` |
| Alert 적색 | `#e2231a` |

## 타이포
- 운영 폰트: **`'Malgun Gothic','맑은 고딕',sans-serif`** — 전 아키타입 본문·제목 기본값(시스템 보장, 표준 렌더 일치).
- 브랜드 권장(표지/대제목): **Noto Serif KR**(제목) · **Pretendard**(소제목/본문) — 시스템 설치 시 적용 가능(미설치 시 Malgun Gothic 폴백).
- 섹션 라벨: 대문자 + `letter-spacing:.26em`, APS Blue `#0b3fd1`, 600 (예: `EXECUTIVE SUMMARY`, `03 — PLAN`).
- 제목 하단 액센트 바: `width:60px; height:4px; background:linear-gradient(90deg,#BED600,#2BA6CB);`.

## 슬라이드 아키타입 (스니펫)
| 파일 | 용도 |
|------|------|
| `sections/00-brand-guide.html` | 브랜드 가이드(슬로건·핵심가치·색상·타이포) |
| `sections/01-cover.html` | 표지 (네이비 + 글로우 + 로고 + 대형 제목) |
| `sections/02-executive-summary.html` | 임원 요약 1페이지 (핵심 포인트 + 수치 + 의사결정 요청) |
| `sections/03-section-divider.html` | 섹션 간지 (큰 번호 + 섹션명) |
| `sections/04-body-two-column.html` | 본문 기본형 (좌 서술 / 우 카드 리스트) |
| `sections/05-metrics.html` | 지표 강조형 (큰 숫자 4분할) |
| `sections/06-phase-steps.html` | 단계/표 (Phase 카드, 현재 단계 네이비 강조) |
| `sections/07-table.html` | 표 (헤더 네이비, 짝수행 음영, 합계행 강조) |
| `sections/08-components.html` | 컴포넌트 라이브러리 (칩·번호 원·프로세스·강조 박스) |
| `sections/09-charts.html` | 차트 샘플 (막대·도넛) |
| `sections/10-wbs-gantt.html` | WBS 간트형 일정 |
| `sections/11-org-rnr.html` | 조직/R&R |
| `sections/12-asis-tobe.html` | As-Is / To-Be 비교 |
| `sections/13-closing.html` | 마무리 (핵심 메시지 + 다음 단계) |
| `sections/_classification.html` | 보안 분류 배지 오버레이 (대외비 등, 각 section 맨 앞 자식) |

## 메타 규칙
- 각 `<section>`은 `data-label`(식별)·`data-speaker-notes`(발표 노트)를 가진다.
- 빌드 스크립트가 `data-speaker-notes`를 PowerPoint 발표자 노트로 옮기며, 노트 텍스트에 `noProof="1"`(맞춤법 검사 안 함)을 설정한다.
- 보안 분류(대외비 등) 해당 시 `sections/_classification.html` 배지를 **모든** 슬라이드에 일관 배치한다.
