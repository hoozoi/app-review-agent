"""
보고서 정량 평가 (LLM 판정) — outputs/<버전>/report.md 를 읽어 평가하고 결과를 저장

품질 기준별 평가 지표
  [기준 1] 분석이 앱과 관련된 리뷰에만 근거하는가
    관련성: 개선안 중 앱의 기능·화면·정책으로 해결할 수 있는 개선안의 비율
    → 보고서에서 개선안을 하나씩 뽑아 개선안마다 O/X 판정 (이때 개선안에 직접 적힌 리뷰 번호도 함께 뽑음)
  [기준 2] 개선안마다 근거가 있는가
    근거 있는 개선안 비율: 실제로 그 문제를 말하는 근거 리뷰가 1개 이상 연결된 개선안의 비율
    → (개선안, 리뷰) 쌍마다 리뷰 원문을 보고 O/X 판정. 근거가 없거나 틀린 개선안은 모두 X
  [기준 3] 읽는 사람이 한눈에 파악할 수 있는가
    질문 응답률: 의사결정 질문(eval_prompts/questions.txt) 중 보고서만 보고 답할 수 있는 질문의 비율
    쌍대 비교 선택률: 두 보고서를 이름을 가리고 나란히 보여 주고, 핵심을 더 빨리 파악할 수 있는 쪽을 고르게 함
    → A/B 순서를 번갈아 바꿔 PAIRWISE_ROUNDS번 반복, 버전별로 선택된 비율
  모든 버전을 같은 방식(같은 프롬프트)으로 추출·판정해 비교가 공정하도록 했습니다.

판정 기준(프롬프트)은 eval_prompts/ 폴더의 텍스트 파일에 있으니, 기준을 바꾸려면 그 파일만 고치면 됩니다.

실행
  python evaluate.py              → v1, v4 평가
  python evaluate.py v1 v3 v4     → 원하는 버전만
  python evaluate.py v1 v2 v3 v4 --pair v3 v4   → 쌍대 비교할 두 버전 지정
  (--pair를 생략하면, 평가 대상에 v3와 v4가 모두 있을 때 v3 vs v4로 비교)

결과 (outputs/eval/)
  relevance_<버전>.xlsx   추출된 개선안, 관련성 판정과 이유, 연결된 근거 리뷰 번호
  grounding_<버전>.xlsx   [개선안별] 근거 있음 여부  [리뷰별] (개선안, 리뷰) 쌍별 판정과 이유
  questions_<버전>.xlsx   질문별 응답 가능 여부, 보고서에서 찾은 답, 이유
  pairwise_<A>_vs_<B>.xlsx  회차별 제시 순서, 선택한 보고서, 이유
  summary.md              품질 기준별·버전별 점수 비교 표 (포트폴리오에 바로 사용)
  eval_log.csv            실행할 때마다 지표가 누적되는 기록
"""
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ───────── 설정 ─────────
MODEL = "claude-sonnet-5-5"
MAX_TOKENS = 16000
TEMPERATURE = 0
REVIEW_FILE = "reviews.xlsx"
PAIRWISE_ROUNDS = 10                 # 쌍대 비교 반복 횟수 (A/B 순서를 번갈아 바꿈)
# ────────────────────────

BASE = Path(__file__).parent
OUT_DIR = BASE / "outputs"
EVAL_DIR = OUT_DIR / "eval"
PROMPT_DIR = BASE / "eval_prompts"


# ───────── LLM 호출 (도구 형식으로 답을 받아 파싱 오류를 없앰) ─────────
_client = None
_use_temp = True


_use_tool_choice = True


def _call(kwargs):
    """모델이 지원하지 않는 설정(temperature, 도구 강제 선택)은 빼고 다시 호출"""
    global _use_temp, _use_tool_choice
    while True:
        try:
            with _client.messages.stream(**kwargs) as s:
                return s.get_final_message()
        except Exception as e:
            msg = str(e)
            if "temperature" in msg and "temperature" in kwargs:
                _use_temp = False
                kwargs.pop("temperature")
            elif "tool_choice" in msg and "tool_choice" in kwargs:
                _use_tool_choice = False
                kwargs.pop("tool_choice")
            else:
                raise


def _check(data, schema, path=""):
    """도구 스키마의 필수 항목과 허용 값(enum)을 지켰는지 확인해, 문제 목록을 반환"""
    problems = []
    if schema.get("type") == "object":
        if not isinstance(data, dict):
            return [f"{path or '결과'}는 객체여야 합니다."]
        for k in schema.get("required", []):
            if k not in data:
                problems.append(f"{path}{k} 항목이 빠졌습니다.")
        for k, sub in schema.get("properties", {}).items():
            if k in data:
                problems += _check(data[k], sub, f"{path}{k}.")
    elif schema.get("type") == "array":
        if not isinstance(data, list):
            return [f"{path.rstrip('.')}는 목록이어야 합니다."]
        for i, x in enumerate(data):
            problems += _check(x, schema.get("items", {}), f"{path.rstrip('.')}[{i}].")
    elif "enum" in schema and data not in schema["enum"]:
        problems.append(f"{path.rstrip('.')} 값은 {schema['enum']} 중 하나여야 합니다. (받은 값: {data!r})")
    return problems


def ask(system, user, tool, max_tries=4):
    """tool 스키마에 맞춘 답을 받아 dict로 반환.
    도구를 호출하지 않았거나 형식이 틀리면 이유를 알려주고 다시 요청함"""
    global _client
    if _client is None:
        import anthropic
        _client = anthropic.Anthropic()
    system = system + f"\n\n결과는 반드시 {tool['name']} 도구를 호출해 제출하세요. 도구의 항목 이름과 허용 값을 그대로 지키세요."
    messages = [{"role": "user", "content": user}]
    for _ in range(max_tries):
        kwargs = dict(model=MODEL, max_tokens=MAX_TOKENS, system=system, tools=[tool], messages=messages)
        if _use_temp:
            kwargs["temperature"] = TEMPERATURE
        if _use_tool_choice:
            kwargs["tool_choice"] = {"type": "tool", "name": tool["name"]}
        resp = _call(kwargs)
        call = next((b for b in resp.content if b.type == "tool_use" and b.name == tool["name"]), None)
        if call is None:
            messages = messages + [
                {"role": "assistant", "content": resp.content},
                {"role": "user", "content": f"{tool['name']} 도구를 호출해 결과를 제출하세요."},
            ]
            continue
        problems = _check(call.input, tool["input_schema"])
        if not problems:
            return call.input
        print(f"    (형식 오류로 다시 요청: {problems[0]})")
        messages = messages + [
            {"role": "assistant", "content": resp.content},
            {"role": "user", "content": [{"type": "tool_result", "tool_use_id": call.id, "is_error": True,
                                          "content": "형식 오류로 반려되었습니다. 고쳐서 다시 제출하세요.\n- "
                                                     + "\n- ".join(problems[:10])}]},
        ]
    raise RuntimeError("모델이 형식에 맞는 답을 주지 않았습니다.")


def load_prompt(name):
    return (PROMPT_DIR / f"{name}.txt").read_text(encoding="utf-8").strip()


# ───────── 1) 관련성 ─────────
RELEVANCE_TOOL = {
    "name": "submit_relevance",
    "description": "보고서에서 뽑아낸 개선안 목록과 개선안별 판정을 제출합니다.",
    "input_schema": {
        "type": "object",
        "properties": {
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "improvement": {"type": "string", "description": "개선안. 보고서에 적힌 개선안을 원문에 가깝게 한 줄로"},
                        "verdict": {"type": "string", "enum": ["O", "X"], "description": "판정"},
                        "reason": {"type": "string", "description": "판정 이유 한 줄"},
                        "cited_reviews": {"type": "array", "items": {"type": "integer"},
                                   "description": "보고서에서 이 개선안에 직접 연결해 적은 리뷰 번호. 없으면 빈 목록"},
                    },
                    "required": ["improvement", "verdict", "reason", "cited_reviews"],
                },
            }
        },
        "required": ["items"],
    },
}


def eval_relevance(version, report):
    print("  [관련성] 개선안과 연결된 리뷰 번호 추출, 관련성 판정 중...")
    res = ask(load_prompt("relevance"), f"[평가할 보고서]\n{report}", RELEVANCE_TOOL)
    items = [{"개선안": x.get("improvement", ""), "판정": x.get("verdict", "누락"), "이유": x.get("reason", ""),
              "근거리뷰": sorted({int(n) for n in (x.get("cited_reviews") or []) if str(n).isdigit()})}
             for x in res["items"]]
    df = pd.DataFrame(items)
    df.insert(0, "번호", range(1, len(df) + 1))
    df["근거리뷰"] = df["근거리뷰"].map(lambda r: ", ".join(map(str, r)))
    df.to_excel(EVAL_DIR / f"relevance_{version}.xlsx", index=False)
    o = int((df["판정"] == "O").sum())
    cited = sum(1 for it in items if it["근거리뷰"])
    print(f"    → 개선안 {len(df)}개 중 앱 관련 {o}개, 근거 리뷰가 연결된 개선안 {cited}개")
    return (o, len(df)), (cited, len(df)), items


# ───────── 2) 근거 정확도 ─────────
GROUNDING_TOOL = {
    "name": "submit_grounding",
    "description": "근거 리뷰별 판정을 제출합니다.",
    "input_schema": {
        "type": "object",
        "properties": {
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "review_no": {"type": "integer", "description": "리뷰 번호"},
                        "verdict": {"type": "string", "enum": ["O", "X"], "description": "판정"},
                        "reason": {"type": "string", "description": "판정 이유 한 줄"},
                    },
                    "required": ["review_no", "verdict", "reason"],
                },
            }
        },
        "required": ["items"],
    },
}


def eval_grounding(version, items, reviews):
    """개선안마다 연결된 근거 리뷰를 판정해,
    '맞는 근거 리뷰가 1개 이상 연결된 개선안' 수를 반환 (분모는 전체 개선안 수)"""
    system = load_prompt("grounding")
    rows, summary = [], []
    cited = [(k, it) for k, it in enumerate(items, 1) if it["근거리뷰"]]
    for i, (k, it) in enumerate(cited, 1):
        lines = []
        for n in it["근거리뷰"]:
            text = reviews.get(n)
            lines.append(f"{n}. {text}" if text else f"{n}. (리뷰 파일에 없는 번호)")
        user = f"[개선안]\n{it['개선안']}\n\n[근거로 연결된 리뷰]\n" + "\n".join(lines)
        print(f"  [근거] {i}/{len(cited)} {it['개선안'][:30]} (리뷰 {len(it['근거리뷰'])}개)")
        res = ask(system, user, GROUNDING_TOOL)
        judged = {int(x["review_no"]): {"판정": x.get("verdict", "누락"), "이유": x.get("reason", "")}
                  for x in res["items"] if str(x.get("review_no", "")).isdigit()}
        for n in it["근거리뷰"]:
            j = judged.get(n, {"판정": "누락", "이유": "판정 결과에 없음"})
            rows.append({"개선안번호": k, "개선안": it["개선안"], "리뷰번호": n, "리뷰": reviews.get(n, ""),
                         "판정": j.get("판정", "누락"), "이유": j.get("이유", "")})
    pairs = pd.DataFrame(rows, columns=["개선안번호", "개선안", "리뷰번호", "리뷰", "판정", "이유"])
    for k, it in enumerate(items, 1):
        mine = pairs[pairs["개선안번호"] == k]
        ok = int((mine["판정"] == "O").sum())
        summary.append({"개선안번호": k, "개선안": it["개선안"], "연결된 근거 수": len(mine),
                        "맞는 근거 수": ok, "근거 있음": "O" if ok else "X"})
    per_issue = pd.DataFrame(summary)
    with pd.ExcelWriter(EVAL_DIR / f"grounding_{version}.xlsx") as w:
        per_issue.to_excel(w, sheet_name="개선안별", index=False)
        pairs.to_excel(w, sheet_name="리뷰별", index=False)
    grounded = int((per_issue["근거 있음"] == "O").sum())
    acc = f"{int((pairs['판정'] == 'O').sum())}/{len(pairs)}" if len(pairs) else "연결된 근거 없음"
    print(f"    → 개선안 {len(items)}개 중 근거 있는 개선안 {grounded}개 (참고: 연결된 근거 중 맞는 근거 {acc})")
    return grounded, len(items)


# ───────── 3) 질문 응답률 ─────────
QUESTION_TOOL = {
    "name": "submit_answers",
    "description": "질문별로 보고서만 보고 답할 수 있는지 판정해 제출합니다.",
    "input_schema": {
        "type": "object",
        "properties": {
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "question_no": {"type": "integer", "description": "질문 번호"},
                        "verdict": {"type": "string", "enum": ["O", "X"], "description": "판정"},
                        "answer": {"type": "string", "description": "보고서에서 찾은 답. 답할 수 없으면 빈 문자열"},
                        "reason": {"type": "string", "description": "판정 이유 한 줄"},
                    },
                    "required": ["question_no", "verdict", "answer", "reason"],
                },
            }
        },
        "required": ["items"],
    },
}


def load_questions():
    lines = (PROMPT_DIR / "questions.txt").read_text(encoding="utf-8").splitlines()
    return [re.sub(r"^\s*\d+[.)]\s*", "", l).strip() for l in lines
            if l.strip() and not l.strip().startswith("#")]


def eval_questions(version, report):
    questions = load_questions()
    q_text = "\n".join(f"{i}. {q}" for i, q in enumerate(questions, 1))
    print(f"  [질문 응답률] 질문 {len(questions)}개 판정 중...")
    res = ask(load_prompt("answerability"),
              f"[질문]\n{q_text}\n\n[보고서]\n{report}", QUESTION_TOOL)
    judged = {int(x["question_no"]): {"판정": x.get("verdict", "누락"), "답": x.get("answer", ""),
                                      "이유": x.get("reason", "")}
              for x in res["items"] if str(x.get("question_no", "")).isdigit()}
    rows = []
    for i, q in enumerate(questions, 1):
        j = judged.get(i, {"판정": "누락", "답": "", "이유": "판정 결과에 없음"})
        rows.append({"번호": i, "질문": q, "판정": j.get("판정", "누락"), "보고서에서 찾은 답": j.get("답", ""),
                     "이유": j.get("이유", "")})
    df = pd.DataFrame(rows)
    df.to_excel(EVAL_DIR / f"questions_{version}.xlsx", index=False)
    o = int((df["판정"] == "O").sum())
    print(f"    → 질문 {len(df)}개 중 답할 수 있는 질문 {o}개")
    return o, len(df)


# ───────── 4) 쌍대 비교 ─────────
PAIRWISE_TOOL = {
    "name": "submit_choice",
    "description": "두 보고서 중 하나를 선택해 제출합니다.",
    "input_schema": {
        "type": "object",
        "properties": {
            "choice": {"type": "string", "enum": ["A", "B"], "description": "선택한 보고서"},
            "reason": {"type": "string", "description": "선택 이유 한두 줄"},
        },
        "required": ["choice", "reason"],
    },
}


def eval_pairwise(a, b, reports, rounds=PAIRWISE_ROUNDS):
    """a, b 보고서를 이름을 가리고 비교. 짝수 회차는 a가 A, 홀수 회차는 b가 A"""
    system = load_prompt("pairwise")
    rows = []
    for i in range(rounds):
        first, second = (a, b) if i % 2 == 0 else (b, a)
        user = f"[보고서 A]\n{reports[first]}\n\n[보고서 B]\n{reports[second]}"
        res = ask(system, user, PAIRWISE_TOOL)
        winner = first if res["choice"] == "A" else second
        rows.append({"회차": i + 1, "보고서 A": first, "보고서 B": second,
                     "선택": res["choice"], "선택된 버전": winner, "이유": res.get("reason", "")})
        print(f"  [쌍대 비교] {i + 1}/{rounds}  A={first}, B={second} → {winner}")
    df = pd.DataFrame(rows)
    df.to_excel(EVAL_DIR / f"pairwise_{a}_vs_{b}.xlsx", index=False)
    wins = {v: int((df["선택된 버전"] == v).sum()) for v in (a, b)}
    print(f"    → {a} {wins[a]}회, {b} {wins[b]}회 선택")
    return {v: (wins[v], rounds) for v in (a, b)}


# ───────── 실행 ─────────
def pct(r):
    o, n = r
    return f"{o / n:.0%} ({o}/{n})" if n else "-"


def main():
    args = sys.argv[1:]
    pair = None
    if "--pair" in args:
        i = args.index("--pair")
        pair = tuple(args[i + 1:i + 3])
        args = args[:i] + args[i + 3:]
        if len(pair) != 2:
            print("사용법: --pair v3 v4")
            sys.exit(1)
    versions = args or ["v1", "v4"]
    if pair is None and "v3" in versions and "v4" in versions:
        pair = ("v3", "v4")
    if not os.environ.get("ANTHROPIC_API_KEY") and _client is None:
        print("API 키가 없습니다. ANTHROPIC_API_KEY를 설정하세요.")
        sys.exit(1)
    EVAL_DIR.mkdir(parents=True, exist_ok=True)
    rv = pd.read_excel(BASE / REVIEW_FILE)
    reviews = {int(r["번호"]): str(r["리뷰"]).strip() for _, r in rv.iterrows()}

    results, reports = {}, {}
    for v in versions:
        path = OUT_DIR / v / "report.md"
        if not path.exists():
            print(f"[{v}] 보고서가 없습니다: {path}")
            continue
        print(f"\n[{v}] 평가 시작")
        report = path.read_text(encoding="utf-8")
        reports[v] = report
        rel, _, items = eval_relevance(v, report)
        results[v] = {"관련성": rel, "근거": eval_grounding(v, items, reviews),
                      "질문": eval_questions(v, report)}

    pair_result = {}
    if pair:
        if all(v in reports for v in pair):
            print(f"\n[쌍대 비교] {pair[0]} vs {pair[1]} ({PAIRWISE_ROUNDS}회)")
            pair_result = eval_pairwise(pair[0], pair[1], reports)
        else:
            print(f"\n[쌍대 비교] {pair} 중 평가 대상에 없는 버전이 있어 건너뜁니다.")

    rows = [
        ("1. 분석이 앱과 관련된 리뷰에만 근거하는가", "관련성", "앱의 기능·화면·정책으로 해결할 수 있는 개선안 비율", "관련성"),
        ("2. 개선안마다 근거가 있는가", "근거 있는 개선안 비율", "실제로 그 문제를 말하는 근거 리뷰가 1개 이상 연결된 개선안 비율", "근거"),
        ("3. 읽는 사람이 한눈에 파악할 수 있는가", "질문 응답률", "의사결정 질문 중 보고서만 보고 답할 수 있는 비율", "질문"),
    ]
    lines = ["| 품질 기준 | 지표 | 정의 | " + " | ".join(results) + " |",
             "|---|---|---|" + "---|" * len(results)]
    for crit, name, desc, key in rows:
        lines.append(f"| {crit} | {name} | {desc} | " + " | ".join(pct(r[key]) for r in results.values()) + " |")
    if pair_result:
        lines.append(f"|  | 쌍대 비교 선택률 ({pair[0]} vs {pair[1]}) | 핵심 문제와 우선순위를 더 빨리 파악할 수 있는 보고서로 선택된 비율 | "
                     + " | ".join(pct(pair_result[v]) if v in pair_result else "-" for v in results) + " |")
    lines += ["", "**질문 목록** (eval_prompts/questions.txt)"]
    lines += [f"{i}. {q}" for i, q in enumerate(load_questions(), 1)]
    lines += ["", f"- 평가: LLM 판정 ({MODEL}), 판정 기준은 eval_prompts/ 참고",
              f"- 실행: {datetime.now():%Y-%m-%d %H:%M}"]
    (EVAL_DIR / "summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    log = EVAL_DIR / "eval_log.csv"
    pd.DataFrame([{"실행시각": datetime.now().strftime("%Y-%m-%d %H:%M"), "버전": v,
                   "관련성": pct(r["관련성"]), "근거있는개선안": pct(r["근거"]),
                   "질문응답률": pct(r["질문"]),
                   "쌍대비교": pct(pair_result[v]) if v in pair_result else ""} for v, r in results.items()]
                 ).to_csv(log, mode="a", header=not log.exists(), index=False, encoding="utf-8-sig")

    print("\n" + "\n".join(lines[:5 + (1 if pair_result else 0)]))
    print(f"\n결과 저장: outputs/eval/ (summary.md, relevance_*.xlsx, grounding_*.xlsx, questions_*.xlsx, pairwise_*.xlsx)")


if __name__ == "__main__":
    main()