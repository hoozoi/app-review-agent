"""
앱 리뷰 분석 에이전트 — 지침 버전만 바꿔 끼우며 실행·검증·기록

구조
  [1단계 분류 에이전트]  리뷰 전체를 읽고, 앱 자체에 대한 리뷰인지 O/X로 판단
      도구: get_reviews, save_judgments, finish_filter
      → 모든 리뷰를 판단하지 않고 끝내려 하면 반려
  [2단계 분석 에이전트]  1단계에서 O로 분류된 리뷰만 받아 분석하고 보고서 제출
      도구: get_reviews(분류된 리뷰만 보임), submit_report
      → 제출하면 하네스가 자동 검증, 기준 미달이면 이유와 함께 반려 → 고쳐서 재제출
  지침 파일은 '무엇을 어떻게 할지'만 담고, 도구 사용법과 검증은 이 코드(하네스)가 제공합니다.

지침 폴더 구조 (버전마다 폴더 하나)
  prompts/v1/analyze.txt                               ← 분류 없이 바로 분석
  prompts/v2/filter.txt + analyze.txt                  ← 분류 단계 추가
  prompts/v3/filter.txt + analyze.txt + harness.json   ← 분류 기준 강화, 분석 지침 구조화, 자동 검증
  prompts/v4/filter.txt + analyze.txt + harness.json   ← 카테고리 구조, 개선안은 검증된 데이터로 작성
  harness.json: {"structured": true} 자동 검증 / {"render_issues": true} 개선안 부분을 검증된 데이터로 작성

설치
  pip install anthropic google-play-scraper pandas openpyxl
  API 키 (PowerShell):  $env:ANTHROPIC_API_KEY="키값"
         (Mac/Linux):   export ANTHROPIC_API_KEY="키값"

1) 리뷰 수집 (최신 100개, 한 번만)
  python run_prompt.py collect "앱 이름 또는 플레이스토어 링크"
  → 이미 reviews.xlsx가 있으면 멈춤. 새로 받으려면 뒤에 --force

2) 에이전트 실행
  python run_prompt.py v1
  python run_prompt.py v2
  python run_prompt.py v3
  python run_prompt.py v4 --filter-from v3   ← 분류는 v3 결과를 재사용하고 분석만 실행

3) 분류 일관성 비교 (같은 지침을 두 번 돌린 뒤 판단이 갈린 리뷰 확인)
  python run_prompt.py diff <결과폴더1> <결과폴더2>
"""
import copy
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ───────── 설정 ─────────
REVIEW_FILE = "reviews.xlsx"
MODEL = "claude-sonnet-5-5"
MAX_TOKENS = 16000
TEMPERATURE = 0                      # 버전 간 비교를 위해 결과 흔들림 최소화 (지원하는 모델에서만 적용)
MAX_TURNS = 30                       # 단계별로 에이전트가 응답할 수 있는 최대 횟수
MAX_SUBMITS = 3                      # 최대 제출 횟수 (반려는 최대 2번, 3번째 제출은 남은 문제를 기록하고 받음)
REVIEW_COUNT = 100                   # 수집할 최신 리뷰 개수
MIN_LENGTH = 10                      # 이보다 짧은 리뷰("굿", "별로" 등)는 제외
# ────────────────────────

BASE = Path(__file__).parent
PROMPT_DIR = BASE / "prompts"
OUT_DIR = BASE / "outputs"
LEVELS = {"하": 0, "중": 1, "상": 2}

FILTER_SYSTEM = """\
당신은 앱 리뷰를 분류하는 에이전트입니다. 아래 도구를 사용해 작업하세요.
- get_reviews: 리뷰를 읽습니다. 한 번에 최대 50개씩 읽을 수 있으니 전체 리뷰를 모두 읽으세요.
- save_judgments: 리뷰별 판단(O/X)과 이유를 저장합니다. 여러 번 나눠 저장해도 됩니다.
- finish_filter: 모든 리뷰의 판단을 저장했으면 호출해 작업을 마칩니다. 빠진 리뷰가 있으면 반려됩니다.

아래는 이번 작업의 지침입니다.
"""

ANALYZE_SYSTEM = """\
당신은 앱 리뷰를 분석하는 에이전트입니다. 아래 도구를 사용해 작업하세요.
- get_reviews: 분석 대상 리뷰를 읽습니다. 한 번에 최대 50개씩 읽을 수 있으니 전체를 모두 읽은 뒤 분석하세요.
- submit_report: 작업을 마치면 최종 결과를 제출하세요.{extra}
반드시 submit_report로 제출해야 작업이 끝납니다.

아래는 이번 작업의 지침입니다.
"""
# 구조화 제출(harness.json의 "structured": true)일 때만 붙는 안내
STRUCTURED_EXTRA = """ 개선안은 issues에도 항목별로 담아 제출하세요.
제출한 결과가 검증을 통과하지 못하면 이유가 돌아옵니다. 그 이유를 바탕으로 고쳐서 다시 제출하세요."""

GET_REVIEWS = {
    "name": "get_reviews",
    "description": "리뷰를 순서대로 읽습니다. offset번째부터 count개(최대 50)를 반환합니다. 각 줄 앞의 숫자가 리뷰 번호입니다.",
    "input_schema": {
        "type": "object",
        "properties": {
            "offset": {"type": "integer", "description": "몇 번째 리뷰부터 읽을지 (0부터)"},
            "count": {"type": "integer", "description": "읽을 개수 (최대 50)"},
        },
        "required": ["offset", "count"],
    },
}

FILTER_TOOLS = [
    GET_REVIEWS,
    {
        "name": "save_judgments",
        "description": "리뷰별 판단을 저장합니다. 같은 번호는 덮어씁니다.",
        "input_schema": {
            "type": "object",
            "properties": {
                "items": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "번호": {"type": "integer"},
                            "판단": {"type": "string", "enum": ["O", "X"], "description": "O: 앱 관련, X: 무관"},
                            "이유": {"type": "string"},
                        },
                        "required": ["번호", "판단", "이유"],
                    },
                }
            },
            "required": ["items"],
        },
    },
    {
        "name": "finish_filter",
        "description": "모든 리뷰의 판단을 저장했으면 호출해 분류 작업을 마칩니다.",
        "input_schema": {"type": "object", "properties": {}},
    },
]

SUBMIT_SIMPLE = {
    "name": "submit_report",
    "description": "최종 분석 결과를 제출합니다.",
    "input_schema": {
        "type": "object",
        "properties": {"report": {"type": "string", "description": "최종 분석 결과 전문 (마크다운)"}},
        "required": ["report"],
    },
}

SUBMIT_STRUCTURED = {
    "name": "submit_report",
    "description": "최종 분석 결과를 제출합니다. report에는 사람이 읽을 전체 결과를, issues에는 개선안을 항목별로 담습니다.",
    "input_schema": {
        "type": "object",
        "properties": {
            "report": {"type": "string", "description": "최종 분석 결과 전문 (마크다운)"},
            "issues": {
                "type": "array",
                "description": "우선순위에 들어가는 개선안 목록 (지침에서 정한 순서대로). 기타 요청은 report에만 적음",
                "items": {
                    "type": "object",
                    "properties": {
                        "개선안": {"type": "string", "description": "무엇을 바꾸는지 한 줄 제목"},
                        "카테고리": {"type": "string"},
                        "근거리뷰": {"type": "array", "items": {"type": "integer"}},
                        "개선방법": {"type": "string"},
                        "기대효과": {"type": "string"},
                        "난이도": {"type": "string", "enum": ["하", "중", "상"]},
                        "난이도이유": {"type": "string"},
                    },
                    "required": ["개선안", "카테고리", "근거리뷰", "개선방법", "기대효과", "난이도", "난이도이유"],
                },
            },
        },
        "required": ["report", "issues"],
    },
}


SUBMIT_RENDER = copy.deepcopy(SUBMIT_STRUCTURED)
SUBMIT_RENDER["description"] = ("최종 분석 결과를 제출합니다. report에는 요약과 카테고리별 문제점만 쓰고, "
                                "개선안은 issues에, 근거 리뷰가 1개뿐인 요청은 others에 담습니다. "
                                "개선안 표와 상세 설명은 하네스가 issues로 보고서에 자동으로 붙입니다.")
SUBMIT_RENDER["input_schema"]["properties"]["report"]["description"] = "요약과 카테고리별 문제점만 (개선안은 쓰지 않음)"
SUBMIT_RENDER["input_schema"]["properties"]["others"] = {
    "type": "array",
    "description": "근거 리뷰가 1개뿐인 기타 요청",
    "items": {"type": "object",
              "properties": {"요청": {"type": "string"}, "근거리뷰": {"type": "integer"}},
              "required": ["요청", "근거리뷰"]},
}
SUBMIT_RENDER["input_schema"]["required"] = ["report", "issues", "others"]

RENDER_EXTRA = """ report에는 요약과 카테고리별 문제점만 쓰세요. 개선안은 issues에, 근거 리뷰가 1개뿐인 요청은 others에 담으세요.
개선안 표와 상세 설명, 기타 요청 목록은 검증을 통과한 issues와 others로 보고서에 자동으로 붙습니다.
제출한 결과가 검증을 통과하지 못하면 이유가 돌아옵니다. 그 이유를 바탕으로 고쳐서 다시 제출하세요."""


def render_report(report, issues, others):
    """검증을 통과한 데이터로 개선안 부분을 직접 작성 → 검증한 내용이 보고서에서 빠질 수 없게 함"""
    lines = [report.rstrip(), "", "## 3. 개선안 우선순위", "",
             "해결 난이도가 낮은 개선안부터 순위를 매겼고, 난이도가 같으면 근거 리뷰가 많은 개선안을 앞에 두었습니다. "
             "적은 노력으로 바로 해결할 수 있는 문제부터 처리하기 위한 기준입니다.", "",
             "| 순위 | 개선안 | 카테고리 | 근거 리뷰 수 | 난이도 |", "|---|---|---|---|---|"]
    for i, it in enumerate(issues, 1):
        lines.append(f"| {i} | {it['개선안']} | {it['카테고리']} | {len(set(it['근거리뷰']))} | {it['난이도']} |")
    for i, it in enumerate(issues, 1):
        refs = ", ".join(str(n) for n in sorted(set(it["근거리뷰"])))
        lines += ["", f"### {i}. {it['개선안']}",
                  f"- **개선 방법**: {it['개선방법']}",
                  f"- **기대 효과**: {it['기대효과']}",
                  f"- **난이도**: {it['난이도']} ({it['난이도이유']})",
                  f"- **근거 리뷰**: {refs}"]
    lines += ["", "## 4. 기타 요청 (근거 리뷰 1개)", ""]
    lines += [f"- {o['요청']} ({o['근거리뷰']})" for o in others] or ["- 없음"]
    return "\n".join(lines) + "\n"


def load_harness_options(vdir):
    """버전 폴더의 harness.json으로 하네스 동작을 켬. 없으면 기본(단순 제출, 구조 검증 없음).
    지침(prompt)에 없는 기준이 도구 설명이나 검증 메시지로 새어 들어가지 않게 하기 위함."""
    path = vdir / "harness.json"
    opts = {"structured": False, "render_issues": False}
    if path.exists():
        text = path.read_text(encoding="utf-8-sig").strip()
        try:
            opts.update(json.loads(text))
        except json.JSONDecodeError:
            print(f"{path} 내용이 올바르지 않습니다. (현재 내용: {text[:50]!r})")
            print('파일을 열어 다음 한 줄만 들어 있게 저장하세요:  {"structured": true}')
            sys.exit(1)
    return opts


# ───────── 리뷰 수집 ─────────
def resolve_app_id(app):
    """앱 이름, 플레이스토어 링크, 패키지명 중 무엇을 넣어도 패키지명으로 변환"""
    m = re.search(r"[?&]id=([\w.]+)", app)
    if m:
        return m.group(1)
    if re.fullmatch(r"[a-zA-Z]\w*(\.\w+)+", app):
        return app
    from google_play_scraper import search
    found = search(app, lang="ko", country="kr", n_hits=1)
    if not found:
        print(f"'{app}'에 해당하는 앱을 찾지 못했습니다. 플레이스토어 링크를 넣어 보세요.")
        sys.exit(1)
    print(f"검색된 앱: {found[0]['title']} ({found[0]['appId']})")
    return found[0]["appId"]


def collect(app, force=False):
    path = BASE / REVIEW_FILE
    if path.exists() and not force:
        print(f"{REVIEW_FILE}이 이미 있습니다. 버전 간 같은 리뷰로 비교해야 하므로 중단합니다.")
        print("새로 수집하려면 뒤에 --force 를 붙이세요.")
        sys.exit(1)
    from google_play_scraper import reviews, Sort
    app_id = resolve_app_id(app)
    result, _ = reviews(app_id, lang="ko", country="kr", sort=Sort.NEWEST, count=REVIEW_COUNT * 2)
    df = pd.DataFrame(result)
    if df.empty:
        print("리뷰를 가져오지 못했습니다.")
        sys.exit(1)
    df = df[["score", "at", "content"]]
    df.columns = ["별점", "작성일", "리뷰"]
    df["리뷰"] = df["리뷰"].fillna("").str.strip()
    df = df[df["리뷰"].str.len() >= MIN_LENGTH].drop_duplicates("리뷰").head(REVIEW_COUNT)
    df["작성일"] = df["작성일"].dt.strftime("%Y-%m-%d")
    df.insert(0, "번호", range(1, len(df) + 1))
    df.to_excel(path, index=False)
    print(f"리뷰 {len(df)}개 저장: {REVIEW_FILE}")
    print(df["별점"].value_counts().sort_index().rename("개수").to_string())
    print("\n다음 단계: python run_prompt.py v1")


# ───────── 하네스: 도구 실행과 검증 ─────────
def format_reviews(df, offset, count):
    count = max(1, min(int(count), 50))
    rows = df.iloc[int(offset): int(offset) + count]
    if rows.empty:
        return "더 읽을 리뷰가 없습니다."
    lines = [f"{int(r['번호'])}. [별점 {int(r['별점'])}] {str(r['리뷰']).strip()}" for _, r in rows.iterrows()]
    remain = len(df) - (int(offset) + len(rows))
    tail = f"\n(이후 남은 리뷰 {remain}개, 다음 offset={int(offset) + len(rows)})" if remain > 0 else "\n(마지막 리뷰까지 읽었습니다)"
    return f"(전체 {len(df)}개 중)\n" + "\n".join(lines) + tail


class FilterStage:
    def __init__(self, df):
        self.df = df
        self.ids = set(int(n) for n in df["번호"])
        self.judgments = {}
        self.finish_tries = 0
        self.done = False

    def run_tool(self, name, args):
        if name == "get_reviews":
            return format_reviews(self.df, args.get("offset", 0), args.get("count", 50))
        if name == "save_judgments":
            bad = []
            for it in args.get("items", []):
                n = int(it.get("번호", -1))
                if n not in self.ids or it.get("판단") not in ("O", "X"):
                    bad.append(n)
                    continue
                self.judgments[n] = {"판단": it["판단"], "이유": it.get("이유", "")}
            msg = f"저장 완료: 누적 {len(self.judgments)}/{len(self.ids)}개"
            return msg + (f" (잘못된 항목 무시: {bad})" if bad else "")
        if name == "finish_filter":
            self.finish_tries += 1
            missing = sorted(self.ids - set(self.judgments))
            if missing and self.finish_tries < MAX_SUBMITS:
                return f"반려: 판단하지 않은 리뷰가 {len(missing)}개 있습니다. 번호: {missing[:30]}"
            self.done = True
            return "분류 완료." + (f" (미판단 {len(missing)}개는 제외 처리)" if missing else "")
        return f"알 수 없는 도구: {name}"

    def summarize(self, name, args, out):
        if name == "get_reviews":
            return f"리뷰 {args.get('offset')}번째부터 {args.get('count')}개 읽음"
        if name == "save_judgments":
            items = args.get("items", [])
            o = sum(1 for i in items if i.get("판단") == "O")
            return f"{len(items)}개 판단 저장 (O {o} / X {len(items) - o}) → {out}"
        return out


class AnalyzeStage:
    def __init__(self, df, structured=False, render=False):
        self.structured = structured
        self.render = render
        self.df = df.reset_index(drop=True)
        self.ids = set(int(n) for n in df["번호"])
        self.submits = 0
        self.rejections = []
        self.final = None

    def verify(self, report, issues):
        problems = []
        if not report or not report.strip():
            problems.append("report가 비어 있습니다.")
        if not self.structured:
            return problems   # 구조화 제출을 켜지 않은 버전은 내용 검증 없이 그대로 받음
        if not issues:
            problems.append("issues가 비어 있습니다. 개선안을 항목별로 담아 주세요.")
            return problems
        for i, it in enumerate(issues, 1):
            name = it.get("개선안") or f"{i}번 개선안"
            for field in ["개선안", "카테고리", "개선방법", "기대효과"]:
                if not str(it.get(field, "")).strip():
                    problems.append(f"[{name}] '{field}'이 비어 있습니다.")
            refs = it.get("근거리뷰") or []
            if not refs:
                problems.append(f"[{name}] 근거 리뷰가 없습니다.")
            outside = [n for n in refs if n not in self.ids]
            if outside:
                problems.append(f"[{name}] 분석 대상이 아닌 리뷰 번호를 근거로 들었습니다: {outside}")
            if it.get("난이도") not in LEVELS:
                problems.append(f"[{name}] 난이도는 하/중/상 중 하나여야 합니다.")
            if not str(it.get("난이도이유", "")).strip():
                problems.append(f"[{name}] 난이도 판단 이유가 비어 있습니다.")
            if len(set(refs)) == 1:
                problems.append(f"[{name}] 근거 리뷰가 1개뿐입니다. 우선순위에서 빼고 '기타 요청'으로 옮기세요.")
        if self.render:
            if re.search(r"^#+\s*\d*\.?\s*개선안", report or "", re.M):
                problems.append("report에 개선안 섹션을 쓰지 마세요. 개선안은 issues로만 제출하면 자동으로 붙습니다.")
            cats = {str(it.get("카테고리", "")).strip() for it in issues}
            if not 3 <= len(cats) <= 7:
                problems.append(f"개선안의 카테고리가 {len(cats)}개입니다. 비슷한 문제를 묶어 7개 이하의 카테고리로 다루세요.")
        keys = [(LEVELS.get(it.get("난이도"), 9), -len(set(it.get("근거리뷰") or []))) for it in issues]
        if keys != sorted(keys):
            problems.append("우선순위 정렬이 맞지 않습니다. 난이도가 낮은 순, 같으면 근거 리뷰가 많은 순이어야 합니다.")
        return problems

    def run_tool(self, name, args):
        if name == "get_reviews":
            return format_reviews(self.df, args.get("offset", 0), args.get("count", 50))
        if name == "submit_report":
            self.submits += 1
            report, issues = args.get("report", ""), args.get("issues")
            others = args.get("others") or []
            problems = self.verify(report, issues)
            if self.render:
                for o in others:
                    if o.get("근거리뷰") not in self.ids:
                        problems.append(f"[기타 요청: {o.get('요청')}] 분석 대상이 아닌 리뷰 번호입니다: {o.get('근거리뷰')}")
            if problems and self.submits < MAX_SUBMITS:
                self.rejections.append(problems)
                return "검증 실패로 반려되었습니다. 아래 문제를 고쳐 다시 제출하세요.\n- " + "\n- ".join(problems)
            if self.render and issues:
                report = render_report(report, issues, others)
            self.final = {"report": report, "issues": issues or [], "others": others, "unresolved": problems}
            return "제출이 완료되었습니다." + (" (재제출 한도 초과로 남은 문제가 기록되었습니다)" if problems else "")
        return f"알 수 없는 도구: {name}"

    @property
    def done(self):
        return self.final is not None

    def summarize(self, name, args, out):
        if name == "get_reviews":
            return f"리뷰 {args.get('offset')}번째부터 {args.get('count')}개 읽음"
        if name == "submit_report":
            return f"보고서 제출 (개선안 {len(args.get('issues') or [])}개) → {out}"
        return out


# ───────── 에이전트 루프 ─────────
def check_api_key():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("API 키가 없습니다. ANTHROPIC_API_KEY를 설정하세요.")
        sys.exit(1)


def run_stage(label, system, tools, stage, trace, client=None):
    trace.append(f"\n## {label}\n")
    print(f"\n[{label}]")
    _loop(client, system, tools, stage, trace, nudge="작업을 마쳤다면 지정된 도구를 호출해 끝내세요.")


def _record(stage, trace, turn, calls):
    results = []
    for cid, name, args in calls:
        out = stage.run_tool(name, args)
        trace.append(f"- [{turn}] `{name}` {stage.summarize(name, args, out)}")
        results.append((cid, out))
    print(f"  턴 {turn}: " + ", ".join(c[1] for c in calls))
    return results


def _loop(client, system, tools, stage, trace, nudge):
    if client is None:
        import anthropic
        client = anthropic.Anthropic()
    messages = [{"role": "user", "content": "지침에 따라 작업을 시작하세요."}]
    use_temp = True
    for turn in range(1, MAX_TURNS + 1):
        kwargs = dict(model=MODEL, max_tokens=MAX_TOKENS, system=system, tools=tools, messages=messages)
        if use_temp:
            kwargs["temperature"] = TEMPERATURE
        try:
            with client.messages.stream(**kwargs) as stream:
                resp = stream.get_final_message()
        except Exception as e:
            # SDK나 모델이 temperature를 받지 않으면 빼고 다시 호출
            if use_temp and "temperature" in str(e):
                use_temp = False
                kwargs.pop("temperature")
                with client.messages.stream(**kwargs) as stream:
                    resp = stream.get_final_message()
            else:
                raise
        messages.append({"role": "assistant", "content": resp.content})
        thought = "".join(b.text for b in resp.content if b.type == "text").strip()
        if thought:
            trace.append(f"- **[{turn}] 에이전트**: {thought[:300]}")
        calls = [(b.id, b.name, b.input) for b in resp.content if b.type == "tool_use"]
        if not calls:
            messages.append({"role": "user", "content": nudge})
            trace.append(f"- [{turn}] 도구 없이 응답 → 마무리 요청")
            continue
        results = _record(stage, trace, turn, calls)
        messages.append({"role": "user", "content": [
            {"type": "tool_result", "tool_use_id": cid, "content": out} for cid, out in results]})
        if stage.done:
            return


# ───────── 실행 ─────────
def load_judgments(version):
    path = OUT_DIR / version / "filter.xlsx"
    if not path.exists():
        print(f"재사용할 분류 결과가 없습니다: {path}")
        sys.exit(1)
    f = pd.read_excel(path)
    return {int(r["번호"]): {"판단": r["AI 판단"], "이유": r["AI 이유"]} for _, r in f.iterrows()
            if r["AI 판단"] in ("O", "X")}


def run_version(version, filter_from=None, client=None):
    vdir = PROMPT_DIR / version
    analyze_path, filter_path = vdir / "analyze.txt", vdir / "filter.txt"
    if not analyze_path.exists():
        print(f"분석 지침이 없습니다: {analyze_path}")
        sys.exit(1)
    opts = load_harness_options(vdir)     # 설정 오류는 분류 전에 미리 확인
    if client is None:
        check_api_key()
    df = pd.read_excel(BASE / REVIEW_FILE)
    out = OUT_DIR / version
    out.mkdir(parents=True, exist_ok=True)
    trace = [f"# {version} 에이전트 작업 기록", f"- 모델: {MODEL}",
             f"- 실행: {datetime.now():%Y-%m-%d %H:%M}"]
    print(f"[{version}] 리뷰 {len(df)}개")

    # 1단계: 분류
    judgments = None
    if filter_from:
        judgments = load_judgments(filter_from)
        trace.append(f"\n## 1단계 분류\n\n- {filter_from}의 분류 결과 재사용")
        print(f"\n[1단계 분류] {filter_from} 결과 재사용")
    elif filter_path.exists():
        fs = FilterStage(df)
        run_stage("1단계 분류", FILTER_SYSTEM + filter_path.read_text(encoding="utf-8").strip(),
                  FILTER_TOOLS, fs, trace, client)
        judgments = fs.judgments
    else:
        trace.append("\n## 1단계 분류\n\n- 없음 (모든 리뷰를 그대로 분석)")
        print("\n[1단계 분류] 없음 → 모든 리뷰를 분석")

    if judgments is not None:
        sheet = df[["번호", "별점", "리뷰"]].copy()
        sheet["AI 판단"] = sheet["번호"].map(lambda n: judgments.get(int(n), {}).get("판단", "미판단"))
        sheet["AI 이유"] = sheet["번호"].map(lambda n: judgments.get(int(n), {}).get("이유", ""))
        sheet["확인 메모"] = ""
        sheet.to_excel(out / "filter.xlsx", index=False)
        kept = df[df["번호"].map(lambda n: judgments.get(int(n), {}).get("판단") == "O")]
        n_x = int((sheet["AI 판단"] == "X").sum())
        print(f"  → 앱 관련 {len(kept)}개 / 무관 {n_x}개")
    else:
        kept = df

    # 2단계: 분석
    render = opts["structured"] and opts["render_issues"]
    st = AnalyzeStage(kept, structured=opts["structured"], render=render)
    if render:
        system, submit = ANALYZE_SYSTEM.format(extra=RENDER_EXTRA), SUBMIT_RENDER
    elif opts["structured"]:
        system, submit = ANALYZE_SYSTEM.format(extra=STRUCTURED_EXTRA), SUBMIT_STRUCTURED
    else:
        system, submit = ANALYZE_SYSTEM.format(extra=""), SUBMIT_SIMPLE
    tools = [GET_REVIEWS, submit]
    label = "2단계 분석" + (" (구조화 제출·검증 켜짐)" if opts["structured"] else "")
    run_stage(label, system + analyze_path.read_text(encoding="utf-8").strip(),
              tools, st, trace, client)

    # 저장 (실행 중에 outputs 폴더가 지워졌어도 다시 만들어 결과를 잃지 않게 함)
    final = st.final or {"report": "(제출 없음: 최대 턴 수 초과)", "issues": [], "unresolved": []}
    out.mkdir(parents=True, exist_ok=True)
    (out / "report.md").write_text(final["report"], encoding="utf-8")
    (out / "trace.md").write_text("\n".join(trace), encoding="utf-8")
    (out / "result.json").write_text(json.dumps({
        "issues": final["issues"], "others": final.get("others", []), "rejections": st.rejections, "unresolved": final["unresolved"],
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    for p in [filter_path, analyze_path]:
        if p.exists():
            (out / f"prompt_{p.name}").write_text(p.read_text(encoding="utf-8"), encoding="utf-8")

    rejected = st.submits - 1 if st.final else st.submits
    log = OUT_DIR / "run_log.csv"
    pd.DataFrame([{
        "버전": version, "실행시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "분류": f"재사용({filter_from})" if filter_from else ("실행" if judgments is not None else "없음"),
        "분석대상": len(kept), "무관제외": len(df) - len(kept),
        "제출횟수": st.submits, "반려횟수": rejected, "개선안수": len(final["issues"]),
        "미해결문제": len(final["unresolved"]) if st.final else "미제출",
    }]).to_csv(log, mode="a", header=not log.exists(), index=False, encoding="utf-8-sig")

    print(f"\n제출 {st.submits}회, 검증 반려 {rejected}회")
    if final["unresolved"]:
        print("재제출 한도를 넘겨 남은 문제:", *final["unresolved"], sep="\n  - ")
    print(f"결과 저장: outputs/{version}/ (report.md, trace.md" + (", filter.xlsx" if judgments is not None else "") + ")")
    return st, judgments


def diff_filters(a, b):
    """두 실행의 분류 결과를 비교해 판단이 갈린 리뷰를 보여줌 (기준이 애매한 리뷰 찾기)"""
    fa, fb = [pd.read_excel(OUT_DIR / v / "filter.xlsx") for v in (a, b)]
    m = fa[["번호", "별점", "리뷰", "AI 판단", "AI 이유"]].merge(
        fb[["번호", "AI 판단", "AI 이유"]], on="번호", suffixes=(f"_{a}", f"_{b}"))
    diff = m[m[f"AI 판단_{a}"] != m[f"AI 판단_{b}"]]
    oa, ob = int((fa["AI 판단"] == "O").sum()), int((fb["AI 판단"] == "O").sum())
    print(f"앱 관련(O): {a} {oa}개 / {b} {ob}개")
    print(f"판단이 같은 리뷰: {len(m) - len(diff)}/{len(m)}   갈린 리뷰: {len(diff)}개\n")
    for _, r in diff.iterrows():
        print(f"{int(r['번호']):>3}. {str(r['리뷰'])[:70]}")
        print(f"     {a}: {r[f'AI 판단_{a}']} - {r[f'AI 이유_{a}']}")
        print(f"     {b}: {r[f'AI 판단_{b}']} - {r[f'AI 이유_{b}']}\n")
    path = OUT_DIR / f"diff_{a}_vs_{b}.xlsx"
    diff.to_excel(path, index=False)
    print(f"저장: {path.name}")
    log = OUT_DIR / "consistency_log.csv"
    pd.DataFrame([{"비교": f"{a} vs {b}", "실행시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
                   "O_첫번째": oa, "O_두번째": ob, "전체": len(m), "갈린리뷰": len(diff)}]
                 ).to_csv(log, mode="a", header=not log.exists(), index=False, encoding="utf-8-sig")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    if sys.argv[1] == "diff":
        if len(sys.argv) < 4:
            print("사용법: python run_prompt.py diff <결과폴더1> <결과폴더2>   예) diff v3 v4")
            sys.exit(1)
        diff_filters(sys.argv[2], sys.argv[3])
        return
    if sys.argv[1] == "collect":
        if len(sys.argv) < 3:
            print('사용법: python run_prompt.py collect "앱 이름 또는 플레이스토어 링크"')
            sys.exit(1)
        collect(sys.argv[2], force="--force" in sys.argv)
        return
    filter_from = None
    if "--filter-from" in sys.argv:
        i = sys.argv.index("--filter-from")
        filter_from = sys.argv[i + 1] if i + 1 < len(sys.argv) else None
    run_version(sys.argv[1], filter_from=filter_from)


if __name__ == "__main__":
    main()