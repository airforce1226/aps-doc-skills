# design-ppt 디자인 토큰

원본: "사내 표준 백엔드 아키텍처와 테스트 LDAP 구축 보고 v1.0" (Claude Design)에서 추출.

## 캔버스
- 슬라이드: **1920 × 1080 px** 고정 (`<section style="width:1920px;height:1080px">`)
- 패딩: 본문 96px, 표지 120px, 좌우 130px

## 컬러
| 용도 | HEX |
|------|-----|
| 네이비(표지/강조 배경) | `#0b1b3a` |
| 표지 텍스트 | `#ffffff` |
| 라이트(본문 배경) | `#f5f7fa` |
| 본문 텍스트 | `#0b1b3a` |
| 액센트 블루 | `#0b3fd1` |
| 표지 글로우(radial) | `rgba(11,63,209,.38)` |
| APS 로고 그라데이션 | `#BED600` → `#2BA6CB` |

## 타이포
- 폰트: **Pretendard** (weight 500/600/700/800), 폴백 `'Malgun Gothic','맑은 고딕',sans-serif`
- Pretendard는 **시스템 설치 전제** (미설치 시 폴백 렌더 → 자간/굵기 차이 가능)

## 슬라이드 아키타입 (스니펫)
| 파일 | 용도 |
|------|------|
| `sections/00-cover.html` | 표지 (네이비 + 글로우 + 제목/부제/작성정보) |
| `sections/01-section-body.html` | 섹션 본문 (제목 + 액센트 바 + 불릿/단락) |
| `sections/02-closing.html` | 마무리 (네이비 + 핵심 메시지) |

## 메타 규칙
- 각 `<section>`은 `data-label`(식별)·`data-speaker-notes`(발표 노트)를 가진다.
- 빌드 스크립트가 `data-speaker-notes`를 PowerPoint 발표자 노트로 옮긴다.
