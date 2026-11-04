#!/usr/bin/env python3
"""회의 자리에서 시트를 받자마자 돌려, 회의 순서대로 정리된 자료를 만든다.

    python meeting_prep.py data/sheets/ --llm data/분해시트_Claude.xlsx -o data/meeting/
    python meeting_prep.py 분해시트_홍길동.xlsx 분해시트_김민수.xlsx -o data/meeting/

입력은 각자 채운 시트 파일들 또는 그 파일들이 든 폴더다. 엑셀이 만드는 '~$' 임시 파일은 건너뛴다.
--llm으로 준 시트는 사람 시트와 따로 다룬다. 사람끼리 비교한 결과를 먼저 만들고,
LLM이 다르게 본 곳은 자료 맨 끝에 둔다. 회의 초반에 LLM의 답이 보이면 거기에 맞추게 되기 때문이다.

출력 (-o 폴더)
  회의자료.md          회의 순서대로 정리한 자료. Notion 회의록에 붙여 넣는다.
  비교.xlsx            사람 시트만 비교한 표 (merge_sheets.py와 같은 형식)
  비교_LLM포함.xlsx    --llm을 준 경우에만

AGENDA는 회차마다 바꾼다. 판단 문장과 질문 번호는 가이드의 "각자 판단할 것"과 맞춘다.
"""

import argparse
import datetime
import glob
import os
import re
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from merge_sheets import (REP_RE, analyze, label_text, link_targets, link_text,  # noqa: E402
                          person_names, write_workbook)

# (질문, 쟁점, 판단 문장, 같이 읽을 문장)
AGENDA = [
    ("Q1", "수치와 결론이 한 문장에 있을 때", [75], []),
    ("Q2", "\"시사한다\", \"~할 수 있다\"로 끝나는 해석 문장을 주장으로 칠지", [4, 78, 88], []),
    ("Q3", "주장이 말하는 비교 대상과 표가 실제로 비교하는 대상", [72, 73], [97]),
    ("Q4", "설계 의도와 실험으로 확인한 결과", [52, 54], []),
    ("Q5", "데이터셋·모델을 고른 이유는 조건인가 빈칸인가", [57, 58, 59, 63], []),
]

ON_HOLD = [
    "조건 목록에 \"언어\"를 넣을지",
    "논문에 없는 조건(\"보고 없음\", 베이스라인 설명 없음)을 어떻게 기록할지",
    "저자 표현과 표 숫자가 다를 때 화면에 표시할지",
]

TIME_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(시간|h|hr|분|m|min)\b", re.I)


# ── 입력 ──────────────────────────────────────────────────────────

def collect(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            files += sorted(glob.glob(os.path.join(p, "*.xlsx")) + glob.glob(os.path.join(p, "*.csv")))
        else:
            files.append(p)
    return [f for f in files if not os.path.basename(f).startswith("~$")]


def minutes(memo):
    total = 0.0
    for num, unit in TIME_RE.findall(memo):
        total += float(num) * (60 if unit.lower() in ("시간", "h", "hr") else 1)
    return round(total) if total else None


def hm(m):
    return f"{m // 60}시간 {m % 60}분" if m >= 60 else f"{m}분"


# ── 마크다운 ──────────────────────────────────────────────────────

def cell(s):
    return str(s).replace("|", "\\|").replace("\n", " ").strip() or " "


def table(header, rows):
    out = ["| " + " | ".join(header) + " |", "|" + " --- |" * len(header)]
    out += ["| " + " | ".join(cell(c) for c in r) + " |" for r in rows]
    return out


def quote(a, n):
    r = a.base[n]
    return [f"> **#{n}** · {r['pos']}", ">", f"> {r['text']}"]


def sentence_block(a, n):
    status, types, _ = a.results[n]
    recs = [sh[n] for sh in a.sheets]
    dist = Counter(label_text(r["labels"]) or "(빈칸)" for r in recs)
    lines = quote(a, n)
    lines += ["", f"상태: **{status}**" + (f" ({', '.join(types)})" if types else "")
              + " · " + " · ".join(f"{k} {v}" for k, v in dist.most_common()), ""]
    lines += table(["이름", "라벨", "연결", "메모"],
                   [[nm, label_text(r["labels"]) or "(빈칸)", link_text(r, a.group_of), r["memo"]]
                    for nm, r in zip(a.names, recs)])
    return lines + [""]


def claim_cell(sh, members):
    mine = [n for n in members if "주장" in sh[n]["labels"]]
    if not mine:
        return "—"
    return ", ".join(f"#{n}" + (" 대표" if REP_RE.search(sh[n]["memo"]) else "")
                     + f" ({sh[n]['strength'] or '?'})" for n in mine)


def evidence_of(a, sh, gid, members):
    ev = {n for n, r in sh.items() if "근거" in r["labels"] and gid in link_targets(r, a.group_of)}
    mine = [n for n in members if "주장" in sh[n]["labels"]]
    from_claim = {t for n in mine for t in sh[n]["claim_ev"]} - ev
    return ev, from_claim


def report(a, llm=None, llm_name=None):
    P = len(a.sheets)
    L = [f"# 분해 연습 비교 결과", "",
         f"{datetime.date.today()} · 사람 {P}명 ({', '.join(a.names)})"
         + (f" · LLM 1개 ({llm_name}, 맨 끝 7절에서만 다룸)" if llm else ""), ""]

    # 0. 입력 확인
    L += ["## 0. 입력 확인", ""]
    times = [minutes(sh[0]["memo"]) if 0 in sh else None for sh in a.sheets]
    got = sorted(t for t in times if t)
    if got:
        mid = got[len(got) // 2]
        L.append(f"- 걸린 시간: {hm(got[0])} ~ {hm(got[-1])} (가운데 {hm(mid)})"
                 + (f", 적지 않음 {times.count(None)}명" if None in times else ""))
    else:
        L.append("- 걸린 시간: 아무도 #0 메모에 적지 않음")
    splits = sorted({n for sh in a.sheets for n, r in sh.items() if "분할" in r["memo"]})
    L.append(f"- 분할 오류 메모: {', '.join(f'#{n}' for n in splits) if splits else '없음'}")
    L.append(f"- 입력 경고 {len(a.warnings)}건" + (":" if a.warnings else ""))
    L += [f"  - {w}" for w in a.warnings]
    L.append("")

    # 1. 주장 연결의 뜻
    L += ["## 1. 찾기 ① 주장 연결의 뜻", "",
          "가이드 기준: 주장의 연결 = 같은 내용을 말하는 다른 주장. 사람마다 다르게 쓰면 서로 다른 주장이 한 묶음으로 뭉친다.", ""]
    rows = []
    for nm, sh in zip(a.names, a.sheets):
        claims = [n for n, r in sh.items() if "주장" in r["labels"]]
        to_claim = sum(1 for n in claims for t in sh[n]["links"] if t in a.group_of and t != n)
        to_other = sum(len(sh[n]["claim_ev"]) for n in claims)
        rows.append([nm, len(claims), to_claim, to_other])
    L += table(["이름", "주장 문장 수", "주장→주장 연결", "주장→다른 문장 연결"], rows) + [""]
    L += ["묶음마다 누구의 연결로 문장들이 한데 묶였는지. 같은 내용이 아닌 문장이 들어 있으면 뭉친 것이다.", ""]
    rows = []
    for i, members in enumerate(a.groups):
        if len(members) < 2:
            continue
        edges = [f"{nm}: " + ", ".join(f"#{n}→#{t}" for n in members for t in sh[n]["links"]
                                       if "주장" in sh[n]["labels"] and t in members and t != n)
                 for nm, sh in zip(a.names, a.sheets)]
        edges = [e for e in edges if not e.endswith(": ")]
        rows.append([f"C{i + 1}", len(members), ", ".join(f"#{n}" for n in members), " / ".join(edges)])
    L += (table(["묶음", "문장 수", "문장", "연결"], rows) if rows else ["- 둘 이상 묶인 주장 없음"]) + [""]

    # 2. 판단 문장
    L += ["## 2. 판단 문장", ""]
    for q, issue, judge, context in AGENDA:
        L += [f"### {q}. {issue}", ""]
        for n in context:
            L += quote(a, n) + [""]
        for n in judge:
            if n in a.base:
                L += sentence_block(a, n)
        L += ["결정: ", ""]

    # 3. 대표 주장
    L += ["## 3. 찾기 ② 대표 주장", "", "칸: 주장으로 고른 문장, (세기). '대표'는 메모에 대표라고 적은 문장.", ""]
    rows = []
    for i, members in enumerate(a.groups):
        rows.append([f"C{i + 1}", ", ".join(f"#{n}" for n in members)]
                    + [claim_cell(sh, members) for sh in a.sheets])
    L += table(["묶음", "문장"] + a.names, rows) + [""]

    # 4. 근거
    L += ["## 4. 찾기 ③ 근거", "", "*는 근거 라벨이 아니라 주장 칸의 연결에 적힌 번호.", ""]
    rows = []
    for i, members in enumerate(a.groups):
        gid = f"C{i + 1}"
        per = [evidence_of(a, sh, gid, members) for sh in a.sheets]
        if not any(ev or fc for ev, fc in per):
            continue
        sets = [ev | fc for ev, fc in per]
        common = set.intersection(*sets) if all(sets) else set()
        rows.append([gid]
                    + [", ".join(f"#{n}" + ("*" if n in fc else "") for n in sorted(ev | fc)) or "—"
                       for ev, fc in per]
                    + [", ".join(f"#{n}" for n in sorted(common)) or "—"])
    L += table(["묶음"] + a.names + ["모두 고른 근거"], rows) + [""]

    # 5. 세기
    L += ["## 5. 세기가 갈린 주장", ""]
    rows = []
    for i, members in enumerate(a.groups):
        vals = [{sh[n]["strength"] for n in members if "주장" in sh[n]["labels"]} - {None} for sh in a.sheets]
        if len(set().union(*vals)) > 1:
            rows.append([f"C{i + 1}", ", ".join(f"#{n}" for n in members)]
                        + [", ".join(sorted(v)) or "—" for v in vals])
    L += (table(["묶음", "문장"] + a.names, rows) if rows else ["- 없음"]) + [""]

    # 6. 그 밖의 갈림
    judge_nums = {n for _, _, js, _ in AGENDA for n in js}
    rest = [n for n in a.agenda() if n not in judge_nums and a.results[n][2] == "주장·근거"]
    L += ["## 6. 그 밖에 주장·근거가 갈린 문장", ""]
    rows = [[f"#{n}", a.results[n][0], ", ".join(a.results[n][1])]
            + [label_text(sh[n]["labels"]) or "(빈칸)" for sh in a.sheets] for n in rest]
    L += (table(["문장", "상태", "유형"] + a.names, rows) if rows else ["- 없음"]) + [""]

    # 7. LLM
    if llm:
        L += [f"## 7. {llm_name}만 다르게 본 곳", "", "사람끼리는 같은 라벨인데 LLM만 다른 문장. 회의 후반에 본다.", ""]
        rows = []
        for n in a.nums:
            human = {sh[n]["labels"] for sh in a.sheets}
            if len(human) == 1 and llm[n]["labels"] not in human:
                h = next(iter(human))
                rows.append([f"#{n}", label_text(h) or "(빈칸)", label_text(llm[n]["labels"]) or "(빈칸)",
                             llm[n]["memo"]])
        L += (table(["문장", "사람", llm_name, f"{llm_name} 메모"], rows) if rows else ["- 없음"]) + [""]

    # 8. 보류
    L += ["## 8. 보류 (결정하지 않고 목록에만)", ""] + [f"- {x}" for x in ON_HOLD] + [""]
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="회의용 비교 자료를 한 번에 만든다")
    ap.add_argument("sheets", nargs="+", help="사람 시트 파일들 또는 폴더")
    ap.add_argument("--llm", help="LLM이 채운 시트. 사람 비교와 따로 다룬다")
    ap.add_argument("--llm-name", default="Claude")
    ap.add_argument("--anon", action="store_true", help="이름 대신 A, B, C로 표시")
    ap.add_argument("-o", "--out", default="meeting", help="출력 폴더")
    args = ap.parse_args()

    files = [f for f in collect(args.sheets)
             if not args.llm or os.path.abspath(f) != os.path.abspath(args.llm)]
    if not files:
        sys.exit("사람 시트를 찾지 못했습니다")
    os.makedirs(args.out, exist_ok=True)

    names = person_names(files, None, args.anon)
    a = analyze(files, names)
    write_workbook(a, os.path.join(args.out, "비교.xlsx"))

    llm = None
    if args.llm:
        both = analyze(files + [args.llm], names + [args.llm_name])
        write_workbook(both, os.path.join(args.out, "비교_LLM포함.xlsx"))
        llm = both.sheets[-1]

    md = os.path.join(args.out, "회의자료.md")
    with open(md, "w", encoding="utf-8") as f:
        f.write(report(a, llm, args.llm_name))

    print(f"사람 {len(files)}명: {', '.join(names)}" + (f" + {args.llm_name}" if llm else ""))
    print(f"  안건 {len(a.agenda())}문장 · 주장 묶음 {len(a.groups)}개 · 입력 경고 {len(a.warnings)}건")
    print(f"  → {md}")
    print(f"  → {os.path.join(args.out, '비교.xlsx')}" + ("  +  비교_LLM포함.xlsx" if llm else ""))


if __name__ == "__main__":
    main()
