#!/usr/bin/env python3
"""팀원들이 각자 채운 분해 시트를 합쳐, 판단이 갈린 문장을 회의 안건으로 뽑는다.

    python merge_sheets.py 홍길동.xlsx 김민수.xlsx 이서연.csv -o 비교.xlsx
    python merge_sheets.py sheets/*.xlsx -o 비교.xlsx --anon     # 이름 대신 A, B, C

입력은 prepare_sheet.py로 만들어 나눠준 시트(xlsx 또는 csv)를 각자 채운 것이다.
사람 이름은 파일 이름에서 공통 앞뒤를 뺀 부분을 쓴다 (--names로 직접 줄 수 있다).
문장 번호와 문장이 시트마다 같아야 한다. 다르면 다른 버전의 시트이므로 멈춘다.

출력 시트
  안건       판단이 갈린 문장만. 판단 문장(시트의 판단 열 ●)과 주장·근거가 걸린 문장 먼저,
             많이 갈린 순서로. 판단 문장이 아닌 곳의 빈칸은 "고르지 않음"으로 읽는다.
  주장 묶음  연결로 이어진 주장 문장들을 한 묶음으로 보고, 사람별로 고른 번호·세기·근거를 나란히.
             메모에 '대표'라고 적은 주장은 번호 옆에 '대표'를 붙인다.
  전체       모든 문장과 사람별 라벨.
  요약       상태별 문장 수, 자주 갈린 라벨 쌍, 입력 경고.
  규칙       회의에서 정한 규칙을 적는 표.

상태
  일치       라벨이 모두 같고 연결·세기도 같음 (모두 빈칸 포함)
  세부 다름  라벨은 같은데 연결한 주장이나 세기가 다름
  다수 일치  과반이 같은 라벨
  한 명만    한 명만 라벨을 붙이고 나머지는 빈칸 (3명 이상일 때)
  갈림       과반이 없음

유형
  경계  서로 다른 라벨을 붙임 (판단 문장에서는 빈칸도 라벨로 본다)
  누락  누구는 붙이고 누구는 빈칸 (판단 문장이 아닌 곳)
  연결  다른 주장에 연결함           세기  주장의 세기가 다름
  대표  같은 주장인데 다른 문장을 주장으로 고름 (초록 #2 vs 결론 #97)

사람별 일치율은 계산하지 않는다. 문장 단위로만 비교한다.
"""

import argparse
import csv
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from itertools import combinations

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

LABELS = ["주장", "근거", "조건:데이터셋", "조건:비교", "조건:지표", "조건:반복", "조건:자원", "한계"]
SPRINT1 = {"주장", "근거"}
BLANK = "(빈칸)"
COLUMNS = ("번호", "위치", "문장", "판단", "라벨", "연결", "메모")

LABEL_SPLIT_RE = re.compile(r"[,，/;·\n]+")
BLANK_TOKENS = {"_", "-", "—", "–", "x", "X", "없음", "빈칸", "(빈칸)", "n/a", "N/A"}   # 빈칸 대신 적는 표시
FIGURE_REF_RE = re.compile(r"(표|그림|Table|Figure|Fig\.?)\s*\d+", re.I)          # '표 1'은 문장 번호가 아니다
RANGE_RE = re.compile(r"(\d+)\s*[~～\-–]\s*(\d+)")
STRENGTH_RE = re.compile(r"세기\s*[=:：]\s*([A-Za-z]+)")
REP_RE = re.compile(r"(?<![가-힣])대표(?![가-힣])")   # "대표적인"은 제외

KIND_ORDER = {"판단": 0, "주장·근거": 1, "": 2}
STATUS_ORDER = {"갈림": 0, "다수 일치": 1, "한 명만": 2, "세부 다름": 3, "일치": 4}
STATUS_FILL = {"갈림": "FEE2E2", "다수 일치": "FEF3C7", "한 명만": "E0E7FF", "세부 다름": "F3F4F6"}


# ── 읽기 ──────────────────────────────────────────────────────────

def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():   # 엑셀이 97을 97.0으로 줄 때
        v = int(v)
    return str(v)


def read_table(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb["분해"] if "분해" in wb.sheetnames else wb.active
        table = [[_cell(v) for v in row] for row in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            table = list(csv.reader(f))
    for i, row in enumerate(table):
        cells = [c.strip() for c in row]
        if "번호" in cells and "라벨" in cells:
            col = {name: cells.index(name) for name in COLUMNS if name in cells}
            return col, table[i + 1:]
    sys.exit(f"{path}: '번호'와 '라벨' 열이 있는 머리글 행을 찾지 못했습니다")


def parse_labels(s, warn):
    out = set()
    for tok in LABEL_SPLIT_RE.split(s):
        tok = re.sub(r"\s+", "", tok).replace("：", ":")
        if not tok or tok in BLANK_TOKENS:
            continue
        if tok not in LABELS:
            warn(f"알 수 없는 라벨 '{tok}'")
        out.add(tok)
    return frozenset(out)


def parse_links(s, warn):
    """'97, 2', '75~81', '36~53(방법론 설명)' 같은 연결 칸을 문장 번호 목록으로."""
    s = s.strip()
    if s in BLANK_TOKENS:
        return []
    if FIGURE_REF_RE.search(s):
        warn(f"연결 '{s}'에 표·그림 이름이 있음 — 캡션 문장 번호로 적어야 셉니다")
        s = FIGURE_REF_RE.sub(" ", s)
    s = re.sub(r"\([^)]*\)", " ", s)
    out = []
    for m in RANGE_RE.finditer(s):
        a, b = int(m.group(1)), int(m.group(2))
        out += range(a, b + 1) if a <= b else [a, b]
    s = RANGE_RE.sub(" ", s)
    out += [int(x) for x in re.findall(r"\d+", s)]
    rest = re.sub(r"[\d\s,，、;·/#]+", "", s)
    if rest:
        warn(f"연결에 번호가 아닌 글자 '{rest}' — 무시함")
    return list(dict.fromkeys(out))


def label_text(labels):
    known = [l for l in LABELS if l in labels]
    return ", ".join(known + sorted(labels - set(LABELS)))


def load_sheet(path, name, warnings):
    col, body = read_table(path)
    get = lambda row, key: row[col[key]].strip() if key in col and col[key] < len(row) else ""
    rows = {}
    for row in body:
        num = get(row, "번호")
        if not num.isdigit():
            continue
        n = int(num)
        warn = lambda msg, n=n: warnings.append(f"{name} #{n}: {msg}")
        memo = get(row, "메모")
        m = STRENGTH_RE.search(memo)
        rows[n] = {
            "pos": get(row, "위치"),
            "text": get(row, "문장"),
            "labels": parse_labels(get(row, "라벨"), warn),
            "links": parse_links(get(row, "연결"), warn),
            "claim_ev": [],   # 주장 칸의 연결에 적은 주장 아닌 문장 (claim_groups가 채운다)
            "memo": re.sub(r"^캡션[\s,·:]*", "", memo),   # 캡션 행은 메모에 '캡션'이 미리 적혀 있다
            "strength": m.group(1).lower() if m else None,
            "judge": bool(get(row, "판단")),
        }
    return rows


def person_names(paths, given, anon):
    if anon:
        return [chr(ord("A") + i) if i < 26 else f"P{i + 1}" for i in range(len(paths))]
    if given:
        names = [s.strip() for s in given.split(",")]
        if len(names) != len(paths):
            sys.exit(f"--names에 {len(names)}명, 파일은 {len(paths)}개입니다")
        return names
    stems = [os.path.splitext(os.path.basename(p))[0] for p in paths]
    names = stems
    if len(stems) > 1:
        pre = len(os.path.commonprefix(stems))
        suf = len(os.path.commonprefix([s[::-1] for s in stems]))
        cut = [s[pre:len(s) - suf] for s in stems]
        if all(cut):
            names = [c.strip(" _-") or c for c in cut]
    else:
        names = [re.sub(r"^분해시트[\s_-]*", "", s) or s for s in stems]
    seen = Counter()
    out = []
    for n in names:
        seen[n] += 1
        out.append(n if seen[n] == 1 else f"{n}{seen[n]}")
    return out


def check_same_sheet(sheets, names):
    norm = lambda s: re.sub(r"\s+", " ", s).strip()
    base = sheets[0]
    problems = []
    for name, sh in zip(names[1:], sheets[1:]):
        only = sorted(base.keys() ^ sh.keys())
        if only:
            problems.append(f"{names[0]}와 {name}: 한쪽에만 있는 번호 {only[:10]}")
        diff = sorted(n for n in base.keys() & sh.keys() if norm(base[n]["text"]) != norm(sh[n]["text"]))
        if diff:
            problems.append(f"{names[0]}와 {name}: 문장이 다른 번호 {diff[:10]}")
    if problems:
        sys.exit("시트 버전이 다릅니다. 같은 번호가 같은 문장을 가리키지 않으면 비교할 수 없습니다.\n  "
                 + "\n  ".join(problems))


# ── 주장 묶음 ─────────────────────────────────────────────────────

def claim_groups(sheets, names, warnings):
    """주장 문장과, 주장끼리 연결한 것을 이어 묶음을 만든다. {문장 번호: 'C1'}"""
    claimed = {n for sh in sheets for n, r in sh.items() if "주장" in r["labels"]}
    parent = {n: n for n in claimed}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for name, sh in zip(names, sheets):
        noted = []
        for n, r in sh.items():
            if "주장" not in r["labels"]:
                continue
            for t in r["links"]:
                if t == n:
                    continue
                if t in claimed:
                    parent[find(n)] = find(t)
                elif t in sh:
                    # 주장 칸에 근거 번호를 적은 경우. 근거로 함께 보여 준다
                    r["claim_ev"].append(t)
            if r["claim_ev"]:
                noted.append(f"#{n}→{','.join(map(str, r['claim_ev']))}")
        if noted:
            warnings.append(f"{name}: 주장의 연결에 주장이 아닌 문장을 적음 ({' '.join(noted)}) — "
                            "주장 묶음 시트에서 근거로 함께 표시(*)")

    members = defaultdict(list)
    for n in sorted(claimed):
        members[find(n)].append(n)
    groups = sorted(members.values(), key=lambda m: m[0])
    return {n: f"C{i + 1}" for i, m in enumerate(groups) for n in m}, groups


def link_targets(r, group_of):
    """주장이 아닌 문장의 연결을 묶음 이름으로. 묶음에 없는 번호는 '#n'."""
    return frozenset(group_of.get(t, f"#{t}") for t in r["links"])


def link_text(r, group_of):
    if "주장" in r["labels"] and not (r["labels"] - {"주장"}):
        return ", ".join(map(str, r["links"]))
    return ", ".join(f"{t} ({group_of[t]})" if t in group_of else str(t) for t in r["links"])


# ── 비교 ──────────────────────────────────────────────────────────

def compare(n, sheets, group_of):
    recs = [sh[n] for sh in sheets]
    labels = [r["labels"] for r in recs]
    counts = Counter(labels)
    top, top_n = counts.most_common(1)[0]
    nonblank = [l for l in labels if l]
    types = []
    judge = recs[0]["judge"]
    if judge:
        # 판단 문장은 모두가 정해야 하는 문장이라 빈칸도 하나의 판단이다
        if len(counts) > 1:
            types.append("경계")
    elif len(set(nonblank)) > 1:
        types.append("경계")
    if not judge and nonblank and len(nonblank) < len(labels):
        # 빈칸인 사람이 같은 묶음의 다른 문장을 주장으로 골랐다면 누락이 아니라 대표 문장 선택 차이
        group = [m for m, g in group_of.items() if g == group_of.get(n)] if n in group_of else []
        blank = [sh for sh, l in zip(sheets, labels) if not l]
        other = [sh for sh in blank if any("주장" in sh[m]["labels"] for m in group)]
        if other:
            types.append("대표")
        if len(other) < len(blank):
            types.append("누락")

    if len(counts) == 1:
        if top - {"주장"} and len({link_targets(r, group_of) for r in recs}) > 1:
            types.append("연결")
        if "주장" in top and len({r["strength"] for r in recs} - {None}) > 1:
            types.append("세기")
        status = "세부 다름" if types else "일치"
    elif len(nonblank) == 1 and len(labels) >= 3:
        status = "한 명만"
    elif top_n * 2 > len(labels):
        status = "다수 일치"
    else:
        status = "갈림"
    if judge:
        kind = "판단"
    elif any(l & SPRINT1 for l in labels):
        kind = "주장·근거"
    else:
        kind = ""
    return status, types, kind


# ── 쓰기 ──────────────────────────────────────────────────────────

HEAD_FILL = PatternFill("solid", fgColor="1F2937")
WORK_FILL = PatternFill("solid", fgColor="FFFBEB")
THIN = Border(bottom=Side(style="thin", color="E5E7EB"))


def write_table(ws, header, rows, widths, work_cols=(), status_col=None, freeze="D2"):
    ws.append(header)
    for row in rows:
        ws.append(row)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for i, c in enumerate(row):
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = THIN
            if i in work_cols:
                c.fill = WORK_FILL
        if status_col is not None and row[status_col].value in STATUS_FILL:
            row[status_col].fill = PatternFill("solid", fgColor=STATUS_FILL[row[status_col].value])
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = freeze
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{len(rows) + 1}"


@dataclass
class Analysis:
    names: list
    sheets: list
    warnings: list
    group_of: dict      # 문장 번호 → 주장 묶음 이름 ('C1')
    groups: list        # 묶음마다 문장 번호 목록
    results: dict       # 문장 번호 → (상태, 유형 목록, 구분)

    @property
    def base(self):
        return self.sheets[0]

    @property
    def nums(self):
        return sorted(self.sheets[0])

    def agenda(self):
        rs = self.results
        ns = [n for n in self.nums if rs[n][0] != "일치"]
        return sorted(ns, key=lambda n: (KIND_ORDER[rs[n][2]], STATUS_ORDER[rs[n][0]], n))


def check_inputs(sheets, names, warnings):
    for name, sh in zip(names, sheets):
        unlinked = [n for n, r in sorted(sh.items()) if r["labels"] - {"주장"} and not r["links"]]
        if unlinked:
            warnings.append(f"{name}: 주장이 아닌 라벨인데 연결이 빈 문장 {len(unlinked)}개 {unlinked[:12]}")
        for n, r in sh.items():
            for t in r["links"]:
                if t not in sh:
                    warnings.append(f"{name} #{n}: 연결 #{t}는 없는 번호")


def analyze(paths, names):
    warnings = []
    sheets = [load_sheet(p, name, warnings) for p, name in zip(paths, names)]
    check_same_sheet(sheets, names)
    group_of, groups = claim_groups(sheets, names, warnings)
    check_inputs(sheets, names, warnings)
    results = {n: compare(n, sheets, group_of) for n in sorted(sheets[0])}
    return Analysis(names, sheets, warnings, group_of, groups, results)


def write_workbook(a, out):
    names, sheets, warnings = a.names, a.sheets, a.warnings
    group_of, groups, results = a.group_of, a.groups, a.results
    base, nums, P = a.base, a.nums, len(sheets)

    def person_cols(n):
        recs = [sh[n] for sh in sheets]
        return ([label_text(r["labels"]) for r in recs]
                + [link_text(r, group_of) for r in recs]
                + [r["memo"] for r in recs])

    person_head = ([f"{nm}\n라벨" for nm in names] + [f"{nm}\n연결" for nm in names]
                   + [f"{nm}\n메모" for nm in names])
    person_widths = [13] * P + [11] * P + [24] * P

    wb = Workbook()

    # 안건
    ws = wb.active
    ws.title = "안건"
    agenda = a.agenda()
    rows = []
    for n in agenda:
        status, types, kind = results[n]
        rows.append([n, base[n]["pos"], base[n]["text"], status, ", ".join(types), kind]
                    + person_cols(n) + ["", "", ""])
    first_work = 6 + 3 * P
    write_table(ws, ["번호", "위치", "문장", "상태", "유형", "구분"] + person_head
                + ["결정\n라벨", "결정\n연결", "근거가 된 규칙"], rows,
                [6, 14, 60, 10, 10, 9] + person_widths + [13, 10, 30],
                work_cols=range(first_work, first_work + 3), status_col=3)

    # 주장 묶음
    ws = wb.create_sheet("주장 묶음")
    rows = []
    for i, members in enumerate(groups):
        gid = f"C{i + 1}"
        picks = Counter(n for sh in sheets for n in members if "주장" in sh[n]["labels"])
        rep = max(members, key=lambda n: (picks[n], -n))
        chose = sum(any("주장" in sh[n]["labels"] for n in members) for sh in sheets)
        claim_cells, strength_cells, evidence_cells = [], [], []
        for sh in sheets:
            mine = [n for n in members if "주장" in sh[n]["labels"]]
            claim_cells.append(", ".join(f"{n} 대표" if REP_RE.search(sh[n]["memo"]) else str(n) for n in mine))
            strength_cells.append(", ".join(dict.fromkeys(sh[n]["strength"] or "?" for n in mine)))
            ev = {n for n, r in sh.items()
                  if "근거" in r["labels"] and gid in link_targets(r, group_of)}
            from_claim = {t for n in mine for t in sh[n]["claim_ev"]} - ev
            evidence_cells.append(", ".join(str(n) if n in ev else f"{n}*"
                                            for n in sorted(ev | from_claim)))
        rows.append([gid, ", ".join(map(str, members)), base[rep]["text"], f"{chose}/{P}"]
                    + claim_cells + strength_cells + evidence_cells + ["", "", ""])
    first_work = 4 + 3 * P
    write_table(ws, ["묶음", "문장 번호", "가장 많이 고른 주장 문장", "고른\n사람"]
                + [f"{nm}\n주장" for nm in names] + [f"{nm}\n세기" for nm in names]
                + [f"{nm}\n근거" for nm in names] + ["결정\n대표 번호", "결정\n세기", "결정\n근거"],
                rows, [6, 12, 60, 7] + [10] * P + [12] * P + [14] * P + [10, 12, 14],
                work_cols=range(first_work, first_work + 3), freeze="D2")

    # 전체
    ws = wb.create_sheet("전체")
    rows = []
    for n in nums:
        status, types, kind = results[n]
        rows.append([n, base[n]["pos"], base[n]["text"], status, ", ".join(types)] + person_cols(n))
    write_table(ws, ["번호", "위치", "문장", "상태", "유형"] + person_head, rows,
                [6, 14, 60, 10, 10] + person_widths, status_col=3)

    # 요약
    ws = wb.create_sheet("요약")
    status_count = Counter(r[0] for r in results.values())
    type_count = Counter(t for r in results.values() for t in r[1])
    pairs = defaultdict(list)
    for n in nums:
        labels = [label_text(sh[n]["labels"]) or BLANK for sh in sheets]
        for a, b in combinations(labels, 2):
            if a != b:
                key = tuple(sorted((a, b)))
                if not pairs[key] or pairs[key][-1] != n:
                    pairs[key].append(n)
    lines = [["사람", P], ["문장", len(nums)], ["주장 묶음", len(groups)], [],
             ["상태", "문장 수"]]
    lines += [[s, status_count.get(s, 0)] for s in STATUS_ORDER]
    lines += [[], ["유형", "문장 수"]]
    lines += [[t, type_count.get(t, 0)] for t in ("경계", "누락", "대표", "연결", "세기")]
    lines += [[], ["자주 갈린 라벨 쌍", "문장 수", "문장 번호"]]
    for (a, b), ns in sorted(pairs.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        lines.append([f"{a} ↔ {b}", len(ns), ", ".join(map(str, ns))])
    lines += [[], ["입력 경고"]] + [[w] for w in warnings or ["없음"]]
    lines += [[], ["주장 묶음 시트의 근거 번호에 붙은 *는 근거 라벨이 아니라 주장 칸의 연결에 적힌 번호입니다."]]
    lines += [["사람별 일치율은 일부러 계산하지 않았습니다. 문장 단위로만 비교합니다."]]
    for line in lines:
        ws.append(line)
    for row in ws.iter_rows():
        if row[0].value in ("상태", "유형", "자주 갈린 라벨 쌍", "입력 경고"):
            for c in row:
                c.font = Font(bold=True)
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABC", [44, 10, 60]):
        ws.column_dimensions[col].width = w

    # 규칙
    ws = wb.create_sheet("규칙")
    example = ["R1 (예시)", "실험 결과 수치를 보고하는 문장은 근거. 수치를 얻은 데이터셋·설정만 말하는 문장은 "
               "조건:데이터셋.", "76 근거, 56 조건:데이터셋", "57", "확정", ""]
    write_table(ws, ["번호", "규칙", "예시 번호", "경계 사례", "확정/미결", "관련 안건"], [example],
                [10, 60, 22, 14, 10, 16], freeze="A2")
    for c in ws[2]:
        c.font = Font(italic=True, color="6B7280")

    wb.save(out)


def print_summary(a, out):
    status_count = Counter(r[0] for r in a.results.values())
    agenda = a.agenda()
    print(f"{len(a.sheets)}명 · {len(a.nums)}문장 · 주장 묶음 {len(a.groups)}개 → {out}")
    print("  " + " · ".join(f"{s} {status_count.get(s, 0)}" for s in STATUS_ORDER))
    kinds = Counter(a.results[n][2] for n in agenda)
    print(f"  안건 {len(agenda)}문장 (판단 {kinds['판단']} · 주장·근거 {kinds['주장·근거']} · 기타 {kinds['']})")
    if a.warnings:
        print(f"  입력 경고 {len(a.warnings)}건 — 요약 시트 참고")


def main():
    ap = argparse.ArgumentParser(description="팀원 분해 시트를 합쳐 갈린 문장을 찾는다")
    ap.add_argument("sheets", nargs="+", help="각자 채운 시트 (xlsx 또는 csv), 2개 이상")
    ap.add_argument("-o", "--out", default="비교.xlsx")
    ap.add_argument("--names", help="쉼표로 구분한 이름. 파일 순서대로")
    ap.add_argument("--anon", action="store_true", help="이름 대신 A, B, C로 표시")
    args = ap.parse_args()
    if len(args.sheets) < 2:
        sys.exit("시트가 2개 이상 있어야 비교할 수 있습니다")

    names = person_names(args.sheets, args.names, args.anon)
    a = analyze(args.sheets, names)
    write_workbook(a, args.out)
    print_summary(a, args.out)


if __name__ == "__main__":
    main()
