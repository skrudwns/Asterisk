#!/usr/bin/env python3
"""
논문을 문장 단위로 쪼개고 번호를 붙여 라벨링용 CSV를 만듭니다.

사용법:
    python prepare_sheet.py paper.pdf sheet.csv
    python prepare_sheet.py paper.txt sheet.csv

PDF를 주면 직접 텍스트를 뽑습니다.  pip install pymupdf
  - 한국어 PDF에는 공백 문자가 없는 경우가 많아, 글자 사이 간격으로 띄어쓰기를 되살립니다.
  - 제목·저자(요약 앞), 머리말·쪽 번호, 표 안의 글자, 그림 안의 작은 글자,
    각주, 위첨자, 주제어 줄은 뺍니다.
  - 문장 중간에 끼어든 캡션은 그 문단이 끝난 뒤로 옮겨, 문장이 끊기지 않게 합니다.
  - 문단 첫머리의 굵은 소제목("1단계: 문서 구조 기획")은 한글 굵기를 PDF에서 알 수
    없어 문장 앞에 그대로 남습니다.

텍스트 파일을 주면 빈 줄을 문단 경계로 읽습니다.

줄바꿈에서 단어가 잘린 것인지(정|보와) 띄어 쓴 것인지(이러한|정보의)는 PDF만 보고
알 수 없어서, 같은 문서 안의 다른 용례로 추측합니다. 확신이 없는 곳은
sheet.joins.txt 에 남기니 CSV와 대조해 눈으로 확인하세요.

참고문헌과 감사의 글은 문장 배열에 넣지 않습니다.

문장 분할은 kss가 설치돼 있으면 쓰고, 없으면 간단한 규칙으로 대체합니다.
  pip install kss
"""

import csv
import re
import sys
from collections import Counter
from pathlib import Path

# 문장 끝으로 착각하기 쉬운 마침표들
ABBR = [
    "et al.", "Fig.", "Figs.", "Eq.", "Eqs.", "Tab.",
    "cf.", "e.g.", "i.e.", "vs.", "etc.", "No.", "Dr.", "Prof.",
]

# 섹션 제목으로 볼 줄
SECTION_RE = re.compile(
    r"^\s*(?:"
    r"초\s*록|요\s*약|Abstract"
    r"|\d+(?:\.\d+)*\.?\s+\S.{0,38}"          # "3.1 실험 설정"
    r"|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.?\s+\S.{0,38}"       # "Ⅲ. 실험"
    r"|참\s*고\s*문\s*헌|References"
    r"|감사의\s*글|Acknowledge?ments?"
    r")\s*$"
)
ABSTRACT_RE = re.compile(r"^\s*(?:초\s*록|요\s*약|Abstract)\s*$", re.IGNORECASE)
NUMBERED_RE = re.compile(r"^\s*(?:\d+(?:\.\d+)*\.?|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.?)\s+\S")

# 문장 배열에 넣지 않는 섹션과 줄
SKIP_SECTION_RE = re.compile(
    r"^\s*(?:참\s*고\s*문\s*헌|References|감사의\s*글|Acknowledge?ments?)\s*$", re.IGNORECASE
)
KEYWORDS_RE = re.compile(r"^\s*(?:주제어|핵심어|키워드|Keywords?|Index Terms)\s*[:：]", re.IGNORECASE)

# 표·그림 캡션으로 볼 줄
CAPTION_RE = re.compile(r"^\s*(?:표|그림|Table|Figure|Fig\.)\s*\d+", re.IGNORECASE)
# 분할기가 캡션 머리("표 1.")만 따로 떼어 낸 조각
CAPTION_LABEL_RE = re.compile(r"^\s*(?:표|그림|Table|Figure|Fig\.)\s*\d+\s*[.:]?\s*$", re.IGNORECASE)

BULLET_RE = re.compile(r"^\s*[•∙◦▪●]")
PAGE_NO_RE = re.compile(r"^\s*[-–—]?\s*\d{1,4}\s*[-–—]?\s*$")
SENTENCE_END_RE = re.compile(r"[.?!:。]['’”)]?$")
HANGUL_RE = re.compile(r"[가-힣]")

SENTINEL = "\x00"


def _protect(text: str) -> str:
    """문장 끝이 아닌 마침표를 잠시 치환한다."""
    for abbr in ABBR:
        text = text.replace(abbr, abbr.replace(".", SENTINEL))
    text = re.sub(r"(?<=\d)\.(?=\d)", SENTINEL, text)       # 3.2
    text = re.sub(r"(?<=[A-Za-z])\.(?=[A-Za-z]\.)", SENTINEL, text)  # U.S.A.
    return text


def _restore(text: str) -> str:
    return text.replace(SENTINEL, ".")


def split_sentences(text: str) -> list[str]:
    try:
        import kss
        parts = [s.strip() for s in kss.split_sentences(text) if s.strip()]
    except Exception:
        protected = _protect(text)
        parts = [_restore(p).strip() for p in re.split(r"(?<=[.?!])\s+", protected) if p.strip()]

    # "표 1." 뒤에서 끊긴 캡션을 다시 붙인다
    merged: list[str] = []
    for p in parts:
        if merged and CAPTION_LABEL_RE.match(merged[-1]):
            merged[-1] += " " + p
        else:
            merged.append(p)
    return merged


# ---------------------------------------------------------------- 줄 잇기

# 홀로 써도 단어인 한 글자
STANDALONE_SYLLABLES = set("수두본각및더그이또즉약총등것때중간")
# 어절이 흔히 끝나는 글자 (조사·어미)
WORD_FINAL = set("은는이가을를에의와과도로서고며면다지만한인된될할던적게요")
PARTICLES = ("의", "을", "를", "은", "는", "이", "가", "에", "와", "과", "로", "으로", "도", "만", "서", "히")
LONG_PARTICLES = ("이나", "에서", "에게", "으로", "부터", "까지", "처럼")


class Joiner:
    """줄바꿈 자리에 공백을 넣을지 문서 안의 다른 용례로 정한다."""

    def __init__(self, lines: list[str]):
        self.lines = lines
        self.uncertain: list[str] = []
        self._cache: dict[str, int] = {}

    def _inline(self, s: str) -> int:
        # 줄 안에서만 센다. 지금 판단하는 줄바꿈 자리는 세지 않게 된다.
        if s not in self._cache:
            self._cache[s] = sum(line.count(s) for line in self.lines)
        return self._cache[s]

    def _is_word(self, tok: str) -> bool:
        # 줄 안에서 앞에 공백을 두고 따로 쓰인 적이 있는가
        pat = re.compile(r"(?<=\s)" + re.escape(tok) + r"(?=[\s,.)\]])")
        return any(pat.search(line) for line in self.lines)

    def glue(self, left: str, right: str) -> tuple[str, bool]:
        """(사이에 넣을 문자열, 확신 여부)"""
        a = left.split(" ")[-1]
        m = re.match(r"[^\s.,;:!?)\]’”]+", right)
        b = m.group(0) if m else right.split(" ")[0]

        if a.endswith("-"):                       # Gov-|Report, Qwen2.5-7B-|Instruct
            word = re.match(r"[A-Za-z0-9.]*", b).group(0)
            if word and self._inline(a[:-1] + word) > self._inline(a + word):
                return "\b", True                  # 하이픈을 지운다
            return "", True
        if not HANGUL_RE.match(b[:1]):
            return " ", True
        if not HANGUL_RE.search(a[-1:]):          # [5]|의, (PLOS)|과, 17.41|에서
            if re.search(r"[)\]A-Za-z0-9]$", a) and b.startswith(PARTICLES) and not self._is_word(b):
                return "", True
            return " ", True

        together, apart = self._inline(a + b), self._inline(a + " " + b)
        if together != apart:
            return ("" if together > apart else " "), True
        a_word, b_word = self._is_word(a), self._is_word(b)
        if a_word and b_word:
            return " ", True
        if not b_word and len(b) == 1 and b not in STANDALONE_SYLLABLES:
            return "", True                        # 활발|히, 분석하|여
        if not a_word and len(a) == 1 and a not in STANDALONE_SYLLABLES:
            return "", True                        # 정|보와, 논|문
        if not b_word and b.startswith(LONG_PARTICLES):
            return "", True                        # 왜곡|이나, 그룹|에서

        if a[-1] in WORD_FINAL:
            return " ", True                       # 어려움을|겪고

        # 여기부터는 추측
        if not b_word and b.startswith(("하", "했", "화", "으")):
            return "", False                       # 상승|했다, 구조|화되어
        if self._inline(a + b[:1]) > 0 and self._inline(a + " " + b[:1]) == 0:
            return "", False
        return " ", False

    def join(self, lines: list[str]) -> str:
        text = lines[0]
        for line in lines[1:]:
            sep, sure = self.glue(text, line)
            if sep == "\b":
                text = text[:-1] + line
                continue
            if not sure:
                a, b = text[-12:], line[:12]
                shown = a + sep + b
                self.uncertain.append(f"{'띄움' if sep else '붙임'}\t…{shown}…")
            text += sep + line
        return text


# ---------------------------------------------------------------- PDF

def _pdf_items(path: str) -> list[tuple[str, str]]:
    """PDF를 ("heading"|"line"|"break", 텍스트) 목록으로 만든다."""
    try:
        import pymupdf
    except ImportError:
        import fitz as pymupdf

    doc = pymupdf.open(path)
    pages, sizes = [], Counter()

    for pno, page in enumerate(doc):
        w, h = page.rect.width, page.rect.height
        tables = [pymupdf.Rect(t.bbox) for t in page.find_tables().tables]
        lines = []
        for block in page.get_text("rawdict")["blocks"]:
            for line in block.get("lines", []):
                out, prev_x1, size = [], None, Counter()
                spans = line["spans"]
                for span in spans:
                    if span["flags"] & 1:          # 위첨자: 각주 번호, 소속 표시
                        prev_x1 = span["bbox"][2]
                        continue
                    for c in span["chars"]:
                        ch = c["c"]
                        if (prev_x1 is not None and ch != " " and out and out[-1] != " "
                                and c["bbox"][0] - prev_x1 > span["size"] * 0.18):
                            out.append(" ")
                        out.append(ch)
                        prev_x1 = c["bbox"][2]
                        if not ch.isspace():
                            size[round(span["size"], 1)] += 1
                text = re.sub(r"\s+", " ", "".join(out)).strip()
                if not text:
                    continue
                first = next((s for s in spans if any(not c["c"].isspace() for c in s["chars"])), spans[0])
                x0, y0, x1, y1 = line["bbox"]
                rec = {
                    "page": pno, "text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                    "size": size.most_common(1)[0][0] if size else first["size"],
                    "sup_start": bool(first["flags"] & 1),
                    "first_size": first["size"],
                    "first_text": "".join(c["c"] for c in first["chars"]).strip(),
                    "bold_start": bool(first["flags"] & 16),
                    "side": "F" if x1 - x0 > w * 0.6 else ("L" if x0 < w / 2 else "R"),
                    "edge": y0 < h * 0.1 or y1 > h * 0.9,
                    "in_table": any(t.contains(pymupdf.Point((x0 + x1) / 2, (y0 + y1) / 2)) for t in tables),
                }
                sizes.update(size)
                lines.append(rec)
        pages.append(lines)

    body = sizes.most_common(1)[0][0]
    edge_texts = Counter(ln["text"] for lines in pages for ln in lines if ln["edge"])

    kept = []
    for lines in pages:
        footnotes: list[tuple[str, float]] = []
        page_kept: list[dict] = []
        for ln in lines:
            if ln["in_table"] or ln["size"] < body * 0.8:          # 표 안, 그림 안 글자
                continue
            if ln["edge"] and (edge_texts[ln["text"]] >= 2 or PAGE_NO_RE.match(ln["text"])):
                continue                                           # 머리말, 쪽 번호
            small = ln["size"] < body * 0.95
            marked = ln["sup_start"] or (
                ln["first_size"] < body * 0.8 and re.fullmatch(r"[\d*∗†‡§]+", ln["first_text"]))
            if marked and small:                                   # 각주 시작
                footnotes.append((ln["side"], ln["y0"]))
                continue
            if small and any(s == ln["side"] and ln["y0"] >= y for s, y in footnotes):
                continue                                           # 각주 이어지는 줄
            if KEYWORDS_RE.match(ln["text"]):
                continue
            prev = page_kept[-1] if page_kept else None
            if (prev and prev["side"] == "F" and ln["side"] != "F" and abs(ln["x0"] - prev["x0"]) < 3
                    and 0 < ln["y0"] - prev["y0"] < 2 * (prev["y1"] - prev["y0"])):
                ln["side"] = "F"                                   # 전폭 문단의 짧은 마지막 줄
            if (prev and prev["side"] == ln["side"] and abs(prev["y0"] - ln["y0"]) < 2
                    and ln["x0"] >= prev["x1"] - 1):               # "1." + "서론"
                prev["text"] += " " + ln["text"]
                prev["x1"] = ln["x1"]
                continue
            page_kept.append(ln)
        kept.extend(page_kept)

    # 단마다 왼쪽 끝. 들여쓰기를 재는 기준
    lefts: dict[tuple[int, str], float] = {}
    for key in {(ln["page"], ln["side"]) for ln in kept}:
        xs = Counter(round(ln["x0"]) for ln in kept if (ln["page"], ln["side"]) == key)
        lefts[key] = min(x for x, n in xs.items() if n == max(xs.values()))

    def is_heading(ln: dict) -> bool:
        t = ln["text"]
        if ABSTRACT_RE.match(t) or SKIP_SECTION_RE.match(t):
            return True
        return bool(NUMBERED_RE.match(t)) and len(t) <= 60 and (ln["bold_start"] or ln["size"] > body * 1.05)

    items: list[tuple[str, str]] = []
    deferred: list[dict] = []
    last = None
    started = not any(ABSTRACT_RE.match(ln["text"]) for ln in kept)

    def flush_deferred() -> None:
        for d in deferred:
            if CAPTION_RE.match(d["text"]):
                items.append(("break", ""))
            items.append(("line", d["text"]))
        if deferred:
            items.append(("break", ""))
        deferred.clear()

    for ln in kept:
        if not started:                       # 요약 앞의 제목·저자·소속
            if not ABSTRACT_RE.match(ln["text"]):
                continue
            started = True
        if is_heading(ln):
            flush_deferred()
            items.append(("heading", ln["text"]))
            last = None
            continue
        # 문장 중간에 끼어든 캡션·전폭 줄은 문단이 끝난 뒤로 미룬다
        mid_sentence = last is not None and last["side"] != "F" and not SENTENCE_END_RE.search(last["text"])
        if deferred:
            d = deferred[-1]
            follows = ln["side"] == "F" or (
                abs(ln["x0"] - d["x0"]) < 3 and 0 < ln["y0"] - d["y0"] < 2 * (d["y1"] - d["y0"]))
            if follows:
                deferred.append(ln)
                continue
        elif mid_sentence and (ln["side"] == "F" or CAPTION_RE.match(ln["text"])):
            deferred.append(ln)
            continue

        brk = last is None or BULLET_RE.match(ln["text"]) or CAPTION_RE.match(ln["text"])
        if not brk and (ln["side"] == "F") != (last["side"] == "F"):
            brk = True
        if not brk and SENTENCE_END_RE.search(last["text"]):
            indent = ln["x0"] - lefts[(ln["page"], ln["side"])]
            brk = 3 < indent < 20
        if brk:
            items.append(("break", ""))
            flush_deferred()
        items.append(("line", ln["text"]))
        last = ln
    flush_deferred()
    return items


def _text_items(path: str) -> list[tuple[str, str]]:
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines()
    items = []
    for raw in lines:
        line = raw.strip()
        if not line:
            items.append(("break", ""))
        elif SECTION_RE.match(line):
            items.append(("heading", line))
        else:
            items.append(("line", line))
    return items


# ---------------------------------------------------------------- 문단

def read_blocks(items: list[tuple[str, str]], joiner: Joiner) -> list[tuple[str, str, bool]]:
    """(섹션, 문단, 캡션여부) 목록으로 만든다."""
    blocks: list[tuple[str, str, bool]] = []
    section = "(머리말)"
    skip = False
    buf: list[str] = []
    buf_is_caption = False

    def flush() -> None:
        if buf and not skip:
            blocks.append((section, joiner.join(buf), buf_is_caption))
        buf.clear()

    for kind, line in items:
        if kind == "break":                   # 문단 경계
            flush()
            buf_is_caption = False
            continue

        if kind == "heading":                 # 섹션 제목
            flush()
            buf_is_caption = False
            section = re.sub(r"\s+", " ", line).strip()
            skip = bool(SKIP_SECTION_RE.match(line))
            continue

        if KEYWORDS_RE.match(line):
            continue

        if CAPTION_RE.match(line):            # 캡션 시작
            flush()
            buf_is_caption = True

        buf.append(line)

    flush()
    return blocks


def main(src: str, out: str) -> None:
    items = _pdf_items(src) if src.lower().endswith(".pdf") else _text_items(src)
    joiner = Joiner([t for k, t in items if k == "line"])

    rows = []
    idx = 0                               # 0부터 시작. 바꾸려면 이 값만 1로.

    for section, paragraph, is_caption in read_blocks(items, joiner):
        for sentence in split_sentences(paragraph):
            note = "캡션" if is_caption else ""
            rows.append([idx, section, sentence, "", "", note])
            idx += 1

    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["번호", "위치", "문장", "라벨", "연결", "메모"])
        writer.writerows(rows)

    print(f"문장 {idx}개를 {out} 에 저장했습니다.")
    if joiner.uncertain:
        joins = Path(out).with_suffix(".joins.txt")
        joins.write_text(
            "# 줄바꿈 자리를 추측으로 처리한 곳. 틀렸으면 CSV에서 고치세요.\n"
            + "\n".join(joiner.uncertain) + "\n",
            encoding="utf-8",
        )
        print(f"줄바꿈 {len(joiner.uncertain)}곳은 추측입니다. {joins} 를 확인하세요.")
    print("CSV를 열어 문장 경계가 맞는지 눈으로 확인하세요.")
    print("특히 2단 조판에서는 문장이 뒤섞일 수 있습니다.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
