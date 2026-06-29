# -*- coding: utf-8 -*-
"""Build a self-contained template gallery (assets/templates-gallery.html).

Renders every archetype in assets/sections/ at a scaled-down size so anyone can
open one file in a browser and see the current APS slide template forms — what
each <section> looks like and which file it lives in. Run after editing any
archetype to keep the gallery in sync:

    python scripts/build_gallery.py
"""
import re
from pathlib import Path

ASSETS = Path(__file__).resolve().parent.parent / "assets"
SECTIONS = ASSETS / "sections"
OUT = ASSETS / "templates-gallery.html"

SCALE = 0.46  # 1920x1080 -> ~883x497 preview

# filename -> (한글 이름, 용도) — design-tokens.md 아키타입 표와 동일하게 유지한다.
META = {
    "00-brand-guide.html":      ("브랜드 가이드", "슬로건·핵심가치·색상·타이포"),
    "01-cover.html":            ("표지", "네이비 + 글로우 + 로고 + 대형 제목"),
    "02-executive-summary.html":("임원 요약", "핵심 포인트 + 수치 + 의사결정 요청 (1페이지)"),
    "03-section-divider.html":  ("섹션 간지", "큰 번호 + 섹션명"),
    "04-body-two-column.html":  ("본문 2단", "좌 서술 / 우 카드 리스트"),
    "05-metrics.html":          ("지표 강조", "큰 숫자 4분할"),
    "06-phase-steps.html":      ("단계/Phase", "Phase 카드, 현재 단계 강조"),
    "07-table.html":            ("표", "헤더 Navy, 짝수행 음영, 합계행 강조"),
    "08-components.html":       ("컴포넌트", "칩·번호 원·프로세스·강조 박스"),
    "09-charts.html":           ("차트", "막대(네이티브)·도넛(래스터)"),
    "10-wbs-gantt.html":        ("WBS 간트", "일정 막대 타임라인"),
    "11-org-rnr.html":          ("조직/R&R", "조직도·역할 분담"),
    "12-asis-tobe.html":        ("As-Is / To-Be", "현행 vs 개선 비교"),
    "13-closing.html":          ("마무리", "핵심 메시지 + 다음 단계"),
}

GALLERY_CSS = """
:root { color-scheme: light; }
body { background:#eef1f6; color:#0b1b3a; padding:48px 56px 80px;
  font-family:'Pretendard','Malgun Gothic','맑은 고딕',sans-serif; }
.head { max-width:1100px; margin:0 auto 40px; }
.head h1 { font-size:34px; font-weight:800; margin:0 0 10px; letter-spacing:-.01em; }
.head .bar { width:60px; height:4px; background:linear-gradient(90deg,#BED600,#2BA6CB); margin:14px 0 18px; }
.head p { font-size:16px; line-height:1.7; color:#5b6b85; margin:6px 0; }
.head code { background:#dde3ed; padding:2px 7px; border-radius:4px; font-size:14px; }
.grid { display:flex; flex-direction:column; gap:40px; max-width:%(cardw)dpx; margin:0 auto; }
.card { background:#fff; border:1px solid #d8dfe9; border-radius:12px; overflow:hidden;
  box-shadow:0 6px 22px rgba(11,27,58,.07); }
.card .meta { display:flex; align-items:baseline; gap:14px; padding:18px 24px; border-bottom:1px solid #eef1f6; flex-wrap:wrap; }
.card .meta .name { font-size:20px; font-weight:700; }
.card .meta .file { font-size:14px; color:#0b3fd1; font-weight:600; }
.card .meta .use  { font-size:14px; color:#5b6b85; }
.frame { width:%(fw)dpx; height:%(fh)dpx; overflow:hidden; margin:0 auto; }
.frame > .scaler { transform:scale(%(scale)s); transform-origin:top left; width:1920px; height:1080px; }
.note { max-width:%(cardw)dpx; margin:36px auto 0; background:#fff; border:1px dashed #c0392b;
  border-radius:10px; padding:20px 24px; font-size:15px; line-height:1.7; color:#26354f; }
.note b { color:#c0392b; }
"""


def main():
    base_css = (ASSETS / "base.css").read_text(encoding="utf-8") if (ASSETS / "base.css").exists() else ""
    fw, fh = round(1920 * SCALE), round(1080 * SCALE)
    gallery_css = GALLERY_CSS % {"scale": SCALE, "fw": fw, "fh": fh, "cardw": fw}

    cards = []
    for path in sorted(SECTIONS.glob("*.html")):
        if path.name.startswith("_"):
            continue  # _classification.html is an overlay fragment, not a full slide
        name, use = META.get(path.name, (path.stem, ""))
        section = path.read_text(encoding="utf-8")
        cards.append(
            '<div class="card">'
            '<div class="meta"><span class="name">%s</span>'
            '<span class="file">%s</span><span class="use">%s</span></div>'
            '<div class="frame"><div class="scaler">%s</div></div></div>'
            % (name, path.name, use, section)
        )

    note = (
        '<div class="note"><b>보안 분류 배지</b> — '
        '<code>sections/_classification.html</code> 는 단독 슬라이드가 아니라 각 '
        '&lt;section&gt; 맨 앞에 붙이는 오버레이 조각이라 위 갤러리에는 별도 카드로 넣지 않았다. '
        '대외비/기밀 등 해당 시 모든 슬라이드에 일관 배치한다. '
        '<b>페이지 번호</b>는 빌드 시 자동 주입되므로(표지 제외) 위 미리보기에는 보이지 않는다.</div>'
    )

    html = (
        '<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">'
        '<title>APS design-ppt — 템플릿 갤러리</title>'
        "<style>%s\n%s</style></head><body>" % (base_css, gallery_css)
        + '<div class="head"><h1>APS design-ppt 템플릿 갤러리</h1>'
        '<div class="bar"></div>'
        "<p>APS Brand Presentation System 슬라이드 아키타입 미리보기. 각 카드는 실제 "
        "<code>assets/sections/*.html</code> 를 1920×1080 원본 그대로 축소 렌더한 것이다.</p>"
        "<p>덱을 만들 때 필요한 카드의 파일 내용을 <code>deck.html</code> 로 복사해 "
        "<code>[[...]]</code> 자리표시자만 실제 내용으로 교체한다.</p></div>"
        + '<div class="grid">' + "".join(cards) + "</div>"
        + note
        + "</body></html>"
    )
    OUT.write_text(html, encoding="utf-8")
    print("Wrote %s (%d archetypes)" % (OUT, len(cards)))


if __name__ == "__main__":
    main()
