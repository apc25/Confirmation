"""Build fill-in templates from the official 2020 HWP forms by replacing blanks with {{token}} markers.

Usage: python make_template.py [양식 탭 이름 ...]   (no argument = all forms)
"""
import re
import sys
from pathlib import Path

from hwpauto import Hwp

HERE = Path(__file__).resolve().parent
SRC_DIR = HERE.parent / "[별첨]2020 금융거래조회서 양식"
TPL_DIR = HERE / "templates"

DATE = r"20_+년 _+월 _+일"

# (regex over the document text, replacement (may use \1), expected number of replacements).
# Reply-section blanks (filled by the institution) are left alone: they are not preceded by 귀중 / followed by 현재.
RECIPIENT = [
    (r"_+ ?부.지점 귀중", "{{수신처}} 귀중", 2),
]
COMMON = [
    (rf"(귀중[\t ]+){DATE}", r"\1{{작성일}}", 2),
    (rf"{DATE} 현재", "{{조회기준일}} 현재", 2),
    (rf"유효기간 : {DATE}", "유효기간 : {{요구서유효기간}}", 1),
    # Before the 7-1 header rule: in some forms both blanks are the identical string.
    (r"제공받을 자 : _+ 회계법인 또는 _+ 감사반", "제공받을 자 : {{감사인명}}", 1),
    # Keep replacements within one character-style run: AllReplace maps styles by position.
    (r"_+ 회계법인 또는 _+ 감사반(?= \(담당공인회계사)", "{{감사인명}}", 1),
    (r"\(담당공인회계사 _+ \)", "(담당공인회계사 {{담당CPA}})", 1),
    (r"_+ 회계법인\(또는 _+ 감사반\) \(사업자등록번호 : +\)",
     "{{감사인명}} (사업자등록번호 : {{감사인사업자번호}})", 1),
    (r"담당공인회계사 _+", "담당공인회계사 {{담당CPA}}", 1),
    (r"\(전화 +: +fax : +, e-mail : *",
     "(전화 : {{CPA전화}},  fax : {{CPA팩스}},  e-mail : {{CPA이메일}}", 1),
    (r"수수료 환급계좌번호 : +은행명 : +예금주 : +",
     "수수료 환급계좌번호 : {{환급계좌}},  은행명 : {{환급은행}},  예금주 : {{환급예금주}} ", 1),
    (r"회 사 명 _+", "회 사 명 {{회사명}}", 1),
    (r"대표이사 _+ \(인\)", "대표이사 {{대표자}}     (인)", 1),
    # One tab only: the original tabs push 조회 기준일 past the right margin once a name is filled in.
    (r"조회 대상 회사 : *\t+", "조회 대상 회사 : {{회사명}}\t", 1),
    (r"○ 조회 대상 회사 :", "○ 조회 대상 회사 : {{회사명}}", 1),
    (r"\(2이상인 경우 목록첨부\):", "(2이상인 경우 목록첨부): {{법인등록번호}} / {{사업자등록번호}}", 1),
    (r"회 사 명 {5,}", "회 사 명 {{회사명}}", 1),
    (r"대표이사 {5,}\(인\)", "대표이사 {{대표자}}          (인)", 1),
]
SEAL = [(r"여__ 부__", "{{첨부여부}}", 1)]

FORMS = {
    "은행": ("1.은행조회서_수정.hwp", [
        (r"_+은행 _+ 귀중", "{{수신처}} 귀중", 2),
        (rf"당 은행이 {DATE}부터 {DATE}까지", "당 은행이 {{교부시작일}}부터 {{교부종료일}}까지", 1),
        (rf"{DATE}부터 {DATE}까지 당 은행과", "{{당좌시작일}}부터 {{당좌종료일}}까지 당 은행과", 1),
    ]),
    "비은행": ("2.금융거래조회서(비은행)_수정.hwp", RECIPIENT),
    "여신": ("3.금융거래조회서(여신)_수정.hwp", RECIPIENT + SEAL),
    "보험": ("4.보험거래조회서_수정.hwp", RECIPIENT + SEAL),
    "증권": ("5.금융거래조회서(증권)_수정.hwp", RECIPIENT + [
        (rf"{DATE}부터 {DATE}까지 당 사와", "{{거래시작일}}부터 {{거래종료일}}까지 당 사와", 1),
    ]),
    "보증기관": ("6.금융거래조회서(보증기관용).hwp", RECIPIENT + SEAL),
}
FINANCIAL = set(FORMS)

# 채권채무조회서: 3 identical copies (감사인보관용/회신용/회사보관용), so every rule hits 3 places.
# Blanks here are underlined spaces. The matched text (not just lookaround context) must include a label,
# otherwise AllReplace would hit every run of that many spaces in the document.
# The date and seal lines are pushed right with spaces in the original; they become right-aligned paragraphs
# (see RIGHT_ALIGN) so a long company name cannot wrap and push the page onto a second sheet.
FORMS["채권채무"] = ("7.채권채무조회서_법률명수정.hwp", [
    (r"(?m)^ {30,}년 +월 +일(?=\r?\n)", "{{작성일}}", 3),
    (r"(?m)^ +귀하", "{{거래처명}}  귀하", 3),
    (r"20 +년 +월 +일(?= 현재의)", "{{조회기준일}}", 3),
    (r" +회계법인 또는 +감사반(?= 앞으로)", " {{감사인명}}", 3),
    (r" +회계법인 또는 +감사반(?= 귀중)", " {{감사인명}}", 3),
    (r"(?<=감사인명 :) +회계법인 또는 +감사반", " {{감사인명}}", 3),
    (r"담당공인회계사 : +(?=\(전화)", "담당공인회계사 : {{담당CPA}}  ", 3),
    (r"\(전화 : +\)", "(전화 : {{CPA전화}})", 3),
    (r"(?m)^ {40,}\(인\)(?=\r?\n)", "{{회사명}}  대표이사 {{대표자}}  (인)", 3),
])
RIGHT_ALIGN = ["{{작성일}}", "대표이사 {{대표자}}"]
AR_ROWS = 3


def fill_ar_tables(h):
    """Put row tokens into the empty cells of the two amount tables on each copy."""
    h.top()
    for n in range(6):
        sec = "채권" if n % 2 == 0 else "채무"
        h.find("계 정 과 목")
        for r in range(1, AR_ROWS + 1):
            h.run("TableLowerCell")
            for c, (col, align) in enumerate((("계정", "Center"), ("적요", "Left"), ("금액", "Right"), ("비고", "Left"))):
                if c:
                    h.run("TableRightCell")
                h.run("MoveLineEnd", f"ParagraphShapeAlign{align}")
                h.insert_text(f"{{{{{sec}{r}_{col}}}}}")
            h.run("TableLeftCell", "TableLeftCell", "TableLeftCell")
        h.run("TableLowerCell", "TableRightCell", "MoveLineEnd", "ParagraphShapeAlignRight")
        h.insert_text(f"{{{{{sec}합계}}}}")


def build(h, form):
    src_name, rules = FORMS[form]
    h.open(SRC_DIR / src_name)
    for pattern, repl, expected in rules + (COMMON if form in FINANCIAL else []):
        pairs = {m.group(0): m.expand(repl) for m in re.finditer(pattern, h.text())}
        done = sum(h.replace(lit, new) for lit, new in sorted(pairs.items(), key=lambda p: len(p[0]), reverse=True))
        if done != expected:
            raise SystemExit(f"[{form}] {pattern!r}: {expected}곳 예상, {done}곳 바뀜 — 원본 양식 확인 필요")
    if form == "채권채무":
        for marker in RIGHT_ALIGN:
            h.top()
            for _ in range(3):
                h.find(marker)
                h.run("ParagraphShapeAlignRight")
        fill_ar_tables(h)
        if h.replace_everywhere("조회서번호 :", "조회서번호 : {{조회서번호}}") != 3:
            raise SystemExit("[채권채무] 꼬리말 '조회서번호 :' 3곳 예상")
        counts = {t: h.text().count("{{" + t + "}}") for t in re.findall(r"\{\{([^}]+)\}\}", h.text())}
        if any(c % 3 for c in counts.values()) or len(counts) != 8 * AR_ROWS + 2 + 8:
            raise SystemExit(f"[채권채무] 표시어 개수 이상: {counts}")
    text = h.text()
    dst = TPL_DIR / f"{form}.hwp"
    h.save(dst, "HWP")
    h.close_doc()
    tokens = sorted(set(re.findall(r"\{\{([^}]+)\}\}", text)))
    print(f"[{form}] 저장: {dst.name}  표시어 {len(tokens)}개: {', '.join(tokens)}")
    for line in re.findall(r"[^\n]*20\d?_+년[^\n]*", text):
        print(f"  (그대로 둠) {line.strip()}")


if __name__ == "__main__":
    with Hwp() as h:
        for form in sys.argv[1:] or FORMS:
            build(h, form)
