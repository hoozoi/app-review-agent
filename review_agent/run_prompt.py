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
  prompts/v5/filter.txt + analyze.txt + harness.json + compare.txt
                                                       ← 우선순위를 영향도×난이도로, 경쟁 앱 비교 지침 추가
  harness.json: {"structured": true} 자동 검증 / {"render_issues": true} 개선안 부분을 검증된 데이터로 작성
                {"priority": "matrix"} 에이전트는 심각도·난이도만 판단, 영향도와 순서는 하네스가 계산
                {"categories": true} 카테고리별 관련 리뷰를 받아 report와 대조 (비교에 필요)

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

4) 경쟁 앱 비교 (앱 이름이나 플레이스토어 링크를 원하는 만큼)
  python run_prompt.py compare "우버 택시" "타다"
  python run_prompt.py compare "우버 택시" --from 2026-09-01 --to 2026-09-30 --base "카카오 T" --version v5
  → 기준 앱(기본 카카오 T)과 경쟁 앱의 같은 기간 리뷰를 수집·분석한 뒤, 카테고리를 공통으로 묶어 언급 비율 비교
  → 기간을 생략하면 최근 30일. 앱당 최대 200개(넘으면 기간 전체에서 고르게), 앱 관련 리뷰 30개 미만은 '표본 부족'
  → 결과: outputs/compare/<기간>/comparison_*.md, 앱별 분석은 outputs/compare/<기간>/<패키지명>/ (재실행 시 재사용)
"""
import copy
import json
import os
import re
import sys
from datetime import datetime, timedelta
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
BASE_APP = "com.kakao.taxi"          # compare에서 --base를 생략하면 쓰는 기준 앱 (카카오 T)
COMPARE_VERSION = "v5"               # compare에서 --version을 생략하면 쓰는 지침 버전
PERIOD_DAYS = 30                     # compare에서 기간을 생략하면 최근 30일
PERIOD_MAX = 200                     # 기간 안 리뷰가 이보다 많으면 기간 전체에서 고르게 이만큼만 사용 (비용 제한)
PERIOD_MIN = 30                      # 앱 관련 리뷰가 이보다 적으면 비교표에 '표본 부족'으로 표시
PERIOD_SKIP = 10                     # 기간 안 리뷰가 이보다 적으면 분석하지 않고 비교에서 뺌 (비율이 의미 없음)
# ────────────────────────

BASE = Path(__file__).parent
PROMPT_DIR = BASE / "prompts"
OUT_DIR = BASE / "outputs"
LEVELS = {"하": 0, "중": 1, "상": 2}


def rel(path):
    """출력용 경로 (프로젝트 폴더 기준)"""
    try:
        return path.relative_to(BASE)
    except ValueError:
        return path

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
# harness.json의 "priority": "matrix"일 때 붙는 안내 (순서는 하네스가 계산)
MATRIX_EXTRA = "\n개선안마다 심각도와 그 이유도 담으세요. 우선순위 순서는 지침의 기준대로 자동으로 매겨지므로 issues는 어떤 순서로 담아도 됩니다."
# harness.json의 "categories": true일 때 붙는 안내 (비교 단계에서 카테고리별 언급 비율 계산에 씀)
CATEGORIES_EXTRA = "\nreport에 쓴 카테고리는 categories에도 이름과 관련 리뷰 번호를 담으세요. 이름과 관련 리뷰 수가 report와 같아야 합니다."


def submit_tool(opts):
    """harness.json 옵션에 맞는 submit_report 도구. 옵션이 켜진 항목만 스키마에 들어감"""
    if not opts["structured"]:
        return SUBMIT_SIMPLE
    submit = copy.deepcopy(SUBMIT_RENDER if opts["render_issues"] else SUBMIT_STRUCTURED)
    props = submit["input_schema"]["properties"]
    if opts["priority"] == "matrix":
        item = props["issues"]["items"]
        item["properties"]["심각도"] = {"type": "string", "enum": ["하", "중", "상"]}
        item["properties"]["심각도이유"] = {"type": "string"}
        item["required"] += ["심각도", "심각도이유"]
        props["issues"]["description"] = "우선순위에 들어가는 개선안 목록 (순서 무관). 기타 요청은 others에"
    if opts["categories"]:
        props["categories"] = {
            "type": "array",
            "description": "report의 카테고리별 문제점에 쓴 카테고리와 관련 리뷰 번호",
            "items": {"type": "object",
                      "properties": {"카테고리": {"type": "string"},
                                     "관련리뷰": {"type": "array", "items": {"type": "integer"}}},
                      "required": ["카테고리", "관련리뷰"]},
        }
        submit["input_schema"]["required"].append("categories")
    return submit


# ───────── 우선순위 (영향도 × 난이도) ─────────
SEVERITY_WEIGHT = {"하": 1, "중": 2, "상": 3}
HIGH_IMPACT_SHARE = 0.1                # 분석 대상 리뷰의 10% 이상이 근거 리뷰면 영향도 높음
SEVERE_IMPACT_SHARE = 0.05             # 심각도 상이면 5% 이상이어도 영향도 높음 (리뷰 한두 개짜리가 맨 앞에 오는 것 방지)
QUADRANTS = {1: "① 바로 실행", 2: "② 중장기 과제", 3: "③ 여유 있을 때", 4: "④ 보류"}


def matrix_order(issues, total):
    """심각도·난이도(에이전트 판단)로 영향도와 구분을 계산해 정렬. 계산할 수 있는 건 하네스가 함"""
    for it in issues:
        n = len(set(it["근거리뷰"]))
        high = n >= total * HIGH_IMPACT_SHARE or (it["심각도"] == "상" and n >= total * SEVERE_IMPACT_SHARE)
        easy = it["난이도"] != "상"
        it["영향도"] = "높음" if high else "낮음"
        it["구분"] = (1 if easy else 2) if high else (3 if easy else 4)
        it["점수"] = n * SEVERITY_WEIGHT[it["심각도"]]
    return sorted(issues, key=lambda it: (it["구분"], -it["점수"], LEVELS[it["난이도"]]))


def render_matrix(issues, total):
    """우선순위 표 하나와 기준 한 줄. 세부 기준(비율·가중치)은 각주로 내려 읽을 양을 줄임"""
    lines = ["영향이 큰 문제 중 빨리 해결할 수 있는 것부터 순위를 매겼습니다. "
             "(① 바로 실행 → ② 중장기 과제 → ③ 여유 있을 때 → ④ 보류)", "",
             "| 순위 | 개선안 | 구분 | 근거 리뷰 수 |", "|---|---|---|---|"]
    for i, it in enumerate(issues, 1):
        lines.append(f"| {i} | {it['개선안']} | {QUADRANTS[it['구분']]} | {len(set(it['근거리뷰']))} |")
    lines += ["", f"<sub>영향 큼: 근거 리뷰가 분석 대상({total}개)의 {HIGH_IMPACT_SHARE:.0%} 이상, 또는 심각도 상이면서 "
                  f"{SEVERE_IMPACT_SHARE:.0%} 이상. 빨리 해결: 난이도 하·중. 같은 구분 안에서는 근거 리뷰 수 × 심각도(상 3, 중 2, 하 1) 순.</sub>"]
    return lines


def render_report(report, issues, others, total=None):
    """검증을 통과한 데이터로 개선안 부분을 직접 작성 → 검증한 내용이 보고서에서 빠질 수 없게 함.
    total을 주면 영향도×난이도 우선순위(matrix)로, 없으면 난이도 순으로 작성"""
    lines = [report.rstrip(), "", "## 3. 개선안 우선순위", ""]
    if total is not None:
        lines += render_matrix(issues, total)
    else:
        lines += ["해결 난이도가 낮은 개선안부터 순위를 매겼고, 난이도가 같으면 근거 리뷰가 많은 개선안을 앞에 두었습니다. "
                  "적은 노력으로 바로 해결할 수 있는 문제부터 처리하기 위한 기준입니다.", "",
                  "| 순위 | 개선안 | 카테고리 | 근거 리뷰 수 | 난이도 |", "|---|---|---|---|---|"]
        for i, it in enumerate(issues, 1):
            lines.append(f"| {i} | {it['개선안']} | {it['카테고리']} | {len(set(it['근거리뷰']))} | {it['난이도']} |")
    for i, it in enumerate(issues, 1):
        refs = ", ".join(str(n) for n in sorted(set(it["근거리뷰"])))
        if total is not None:   # 판단 결과는 제목 옆 태그로만, 판단 이유는 result.json에 남김 (보고서 분량 줄이기)
            lines += ["", f"### {i}. {it['개선안']}",
                      f"{QUADRANTS[it['구분']]} · 심각도 {it['심각도']} · 난이도 {it['난이도']}",
                      f"- **개선 방법**: {it['개선방법']}",
                      f"- **기대 효과**: {it['기대효과']}",
                      f"- **근거 리뷰**: {refs}"]
            continue
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
    opts = {"structured": False, "render_issues": False, "priority": "difficulty", "categories": False}
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
    ids = search_play(app)
    if not ids:
        print(f"'{app}'에 해당하는 앱을 찾지 못했습니다. 플레이스토어 링크를 넣어 보세요.")
        sys.exit(1)
    print(f"검색된 앱: {app_title(ids[0])} ({ids[0]})  ← 다른 앱이면 플레이스토어 링크를 넣으세요")
    return ids[0]


def search_play(query):
    """플레이스토어 검색 결과의 패키지명 목록.
    google_play_scraper의 search()는 첫 결과의 패키지명이 비어 오는 경우가 있어 검색 페이지에서 직접 읽음"""
    import urllib.parse
    import urllib.request
    url = "https://play.google.com/store/search?" + urllib.parse.urlencode(
        {"q": query, "c": "apps", "hl": "ko", "gl": "kr"})
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    html = urllib.request.urlopen(req, timeout=20).read().decode("utf-8")
    return list(dict.fromkeys(re.findall(r"/store/apps/details\?id=([\w.]+)", html)))


def app_title(app_id):
    from google_play_scraper import app as app_detail
    try:
        return app_detail(app_id, lang="ko", country="kr")["title"]
    except Exception:
        return app_id


def clean_reviews(result):
    """수집한 리뷰를 번호·별점·작성일·리뷰 표로 정리. 짧은 리뷰와 중복 리뷰는 제외"""
    df = pd.DataFrame(result)
    if df.empty:
        return df
    df = df[["score", "at", "content"]]
    df.columns = ["별점", "작성일", "리뷰"]
    df["리뷰"] = df["리뷰"].fillna("").str.strip()
    return df[df["리뷰"].str.len() >= MIN_LENGTH].drop_duplicates("리뷰")


def save_reviews(df, path):
    df = df.copy()
    df["작성일"] = pd.to_datetime(df["작성일"]).dt.strftime("%Y-%m-%d")
    df.insert(0, "번호", range(1, len(df) + 1))
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False)
    return df


def collect_period(app_id, start, end, path):
    """start~end 기간의 리뷰를 모두 가져와 저장. PERIOD_MAX개가 넘으면 기간 전체에서 고르게 뽑음"""
    from google_play_scraper import reviews, Sort
    rows, token = [], None
    for _ in range(50):   # 한 번에 200개씩, 최대 1만 개까지 거슬러 올라감
        batch, token = reviews(app_id, lang="ko", country="kr", sort=Sort.NEWEST, count=200,
                               continuation_token=token)
        rows += [r for r in batch if start <= r["at"].date() <= end]
        if not batch or token is None or batch[-1]["at"].date() < start:
            break
    df = clean_reviews(rows)
    found = len(df)
    if found > PERIOD_MAX:
        picks = sorted({round(i * (found - 1) / (PERIOD_MAX - 1)) for i in range(PERIOD_MAX)})
        df = df.iloc[picks]
    if found:
        save_reviews(df, path)
    return found, len(df)


def collect(app, force=False):
    path = BASE / REVIEW_FILE
    if path.exists() and not force:
        print(f"{REVIEW_FILE}이 이미 있습니다. 버전 간 같은 리뷰로 비교해야 하므로 중단합니다.")
        print("새로 수집하려면 뒤에 --force 를 붙이세요.")
        sys.exit(1)
    from google_play_scraper import reviews, Sort
    app_id = resolve_app_id(app)
    result, _ = reviews(app_id, lang="ko", country="kr", sort=Sort.NEWEST, count=REVIEW_COUNT * 2)
    df = clean_reviews(result)
    if df.empty:
        print("리뷰를 가져오지 못했습니다.")
        sys.exit(1)
    df = save_reviews(df.head(REVIEW_COUNT), path)
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
    def __init__(self, df, structured=False, render=False, matrix=False, categories=False):
        self.structured = structured
        self.render = render
        self.matrix = matrix            # 우선순위를 영향도×난이도로 하네스가 계산
        self.categories = categories    # 카테고리별 관련 리뷰를 구조화해 받음
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
            if self.matrix:
                if it.get("심각도") not in LEVELS:
                    problems.append(f"[{name}] 심각도는 하/중/상 중 하나여야 합니다.")
                reason = str(it.get("심각도이유", "")).strip()
                if not reason:
                    problems.append(f"[{name}] 심각도 판단 이유가 비어 있습니다.")
                elif it.get("심각도") == "상":
                    cited = {int(n) for n in re.findall(r"\d+", reason)} & set(refs)
                    if not cited:
                        problems.append(f"[{name}] 심각도를 상으로 판단했다면 이유에 이용 불가나 금전 손실을 말한 "
                                        "근거 리뷰 번호를 적으세요. 그런 리뷰가 없으면 중으로 판단하세요.")
        if self.render:
            if re.search(r"^#+\s*\d*\.?\s*개선안", report or "", re.M):
                problems.append("report에 개선안 섹션을 쓰지 마세요. 개선안은 issues로만 제출하면 자동으로 붙습니다.")
            cats = {str(it.get("카테고리", "")).strip() for it in issues}
            if not 3 <= len(cats) <= 7:
                problems.append(f"개선안의 카테고리가 {len(cats)}개입니다. 비슷한 문제를 묶어 7개 이하의 카테고리로 다루세요.")
        if self.matrix:
            return problems   # 순서는 하네스가 계산하므로 정렬 검증 없음
        keys = [(LEVELS.get(it.get("난이도"), 9), -len(set(it.get("근거리뷰") or []))) for it in issues]
        if keys != sorted(keys):
            problems.append("우선순위 정렬이 맞지 않습니다. 난이도가 낮은 순, 같으면 근거 리뷰가 많은 순이어야 합니다.")
        return problems

    def verify_categories(self, report, categories, issues):
        """report의 '### 카테고리 (관련 리뷰 N개)'와 categories가 일치하는지, 개선안의 카테고리가 그 안에 있는지"""
        if not categories:
            return ["categories가 비어 있습니다. report의 카테고리마다 이름과 관련 리뷰 번호를 담으세요."]
        problems = []
        heads = {m.group(1).strip(): int(m.group(2)) for m in
                 re.finditer(r"^###\s*(.+?)\s*\(관련 리뷰\s*(\d+)개\)", report or "", re.M)}
        names = set()
        for c in categories:
            name, refs = str(c.get("카테고리", "")).strip(), set(c.get("관련리뷰") or [])
            names.add(name)
            outside = sorted(refs - self.ids)
            if outside:
                problems.append(f"[카테고리: {name}] 분석 대상이 아닌 리뷰 번호입니다: {outside}")
            if name not in heads:
                problems.append(f"[카테고리: {name}] report에 '### {name} (관련 리뷰 N개)' 제목이 없습니다. 이름을 report와 같게 쓰세요.")
            elif heads[name] != len(refs):
                problems.append(f"[카테고리: {name}] report에는 관련 리뷰 {heads[name]}개, categories에는 {len(refs)}개입니다. 숫자를 맞추세요.")
        for it in issues or []:
            if str(it.get("카테고리", "")).strip() not in names:
                problems.append(f"[{it.get('개선안')}] 카테고리 '{it.get('카테고리')}'가 categories에 없습니다.")
        return problems

    def run_tool(self, name, args):
        if name == "get_reviews":
            return format_reviews(self.df, args.get("offset", 0), args.get("count", 50))
        if name == "submit_report":
            self.submits += 1
            report = args.get("report", "")
            issues, bad_issues = as_items(args.get("issues"), "issues")
            others, bad_others = as_items(args.get("others"), "others")
            categories, bad_cats = as_items(args.get("categories"), "categories")
            problems = bad_issues + bad_others + bad_cats + self.verify(report, issues)
            if self.categories:
                problems += self.verify_categories(report, categories, issues)
            if self.render:
                for o in others:
                    if o.get("근거리뷰") not in self.ids:
                        problems.append(f"[기타 요청: {o.get('요청')}] 분석 대상이 아닌 리뷰 번호입니다: {o.get('근거리뷰')}")
            if problems and self.submits < MAX_SUBMITS:
                self.rejections.append(problems)
                return "검증 실패로 반려되었습니다. 아래 문제를 고쳐 다시 제출하세요.\n- " + "\n- ".join(problems)
            if self.matrix and issues and all(it.get("심각도") in LEVELS and it.get("난이도") in LEVELS
                                              for it in issues):
                issues = matrix_order(issues, len(self.ids))
            if self.render and issues:
                report = render_report(report, issues, others,
                                       total=len(self.ids) if self.matrix and "구분" in issues[0] else None)
            self.final = {"report": report, "issues": issues or [], "others": others, "unresolved": problems,
                          "categories": categories}
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


def as_items(value, field):
    """도구 입력의 목록 항목을 [dict, ...]로 맞춤. 모델이 JSON 문자열로 보내면 풀고, 형식이 틀리면 반려 사유를 돌려줌"""
    if value is None:
        return [], []
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            return [], [f"{field}는 항목의 목록이어야 합니다. 문자열이 아니라 항목별 객체로 담아 다시 제출하세요."]
    if not isinstance(value, list) or not all(isinstance(v, dict) for v in value):
        return [], [f"{field}는 항목별 객체의 목록이어야 합니다. 형식을 맞춰 다시 제출하세요."]
    return value, []


# ───────── 에이전트 루프 ─────────
RETRIES = 3                           # 연결 끊김·서버 과부하 같은 일시적 오류일 때 다시 시도하는 횟수


def _call(client, kwargs):
    """응답 하나를 받음. 일시적 오류는 잠시 기다렸다 다시 시도하고, 요청 자체가 잘못된 오류는 그대로 올림"""
    import time
    import anthropic
    for attempt in range(RETRIES + 1):
        try:
            with client.messages.stream(**kwargs) as stream:
                return stream.get_final_message()
        except Exception as e:
            if isinstance(e, anthropic.APIStatusError):
                transient = e.status_code in (408, 429, 500, 502, 503, 529)
            else:   # 연결 끊김·시간 초과만 (SDK가 감싸지 않고 올리는 httpx 오류 포함). 인자 오류 등은 바로 올림
                transient = isinstance(e, anthropic.APIConnectionError) or any(
                    k in type(e).__name__ for k in ("Protocol", "Connect", "Timeout", "ReadError"))
            if not transient or attempt == RETRIES:
                raise
            print(f"  일시적 오류로 다시 시도합니다 ({attempt + 1}/{RETRIES}): {type(e).__name__}")
            time.sleep(5 * (attempt + 1))


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
            resp = _call(client, kwargs)
        except Exception as e:
            # SDK나 모델이 temperature를 받지 않으면 빼고 다시 호출
            if use_temp and "temperature" in str(e):
                use_temp = False
                kwargs.pop("temperature")
                resp = _call(client, kwargs)
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
def load_judgments(version, out_root=None):
    path = (out_root or OUT_DIR) / version / "filter.xlsx"
    if not path.exists():
        print(f"재사용할 분류 결과가 없습니다: {path}")
        sys.exit(1)
    f = pd.read_excel(path)
    return {int(r["번호"]): {"판단": r["AI 판단"], "이유": r["AI 이유"]} for _, r in f.iterrows()
            if r["AI 판단"] in ("O", "X")}


def run_version(version, filter_from=None, client=None, review_path=None, out_root=None):
    """review_path, out_root를 주면 그 리뷰를 분석해 out_root/<버전>/에 저장 (compare에서 앱마다 사용)"""
    vdir = PROMPT_DIR / version
    analyze_path, filter_path = vdir / "analyze.txt", vdir / "filter.txt"
    if not analyze_path.exists():
        print(f"분석 지침이 없습니다: {analyze_path}")
        sys.exit(1)
    opts = load_harness_options(vdir)     # 설정 오류는 분류 전에 미리 확인
    if client is None:
        check_api_key()
    out_root = out_root or OUT_DIR
    df = pd.read_excel(review_path or BASE / REVIEW_FILE)
    out = out_root / version
    out.mkdir(parents=True, exist_ok=True)
    trace = [f"# {version} 에이전트 작업 기록", f"- 모델: {MODEL}",
             f"- 실행: {datetime.now():%Y-%m-%d %H:%M}"]
    print(f"[{version}] 리뷰 {len(df)}개")

    # 1단계: 분류
    judgments = None
    if filter_from:
        judgments = load_judgments(filter_from, out_root)
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
    matrix = opts["structured"] and opts["priority"] == "matrix"
    categories = opts["structured"] and opts["categories"]
    st = AnalyzeStage(kept, structured=opts["structured"], render=render, matrix=matrix, categories=categories)
    if render:
        extra = RENDER_EXTRA
    elif opts["structured"]:
        extra = STRUCTURED_EXTRA
    else:
        extra = ""
    extra += (MATRIX_EXTRA if matrix else "") + (CATEGORIES_EXTRA if categories else "")
    system, submit = ANALYZE_SYSTEM.format(extra=extra), submit_tool(opts)
    tools = [GET_REVIEWS, submit]
    label = "2단계 분석" + (" (구조화 제출·검증 켜짐)" if opts["structured"] else "")
    run_stage(label, system + analyze_path.read_text(encoding="utf-8").strip(),
              tools, st, trace, client)

    # 저장 (실행 중에 outputs 폴더가 지워졌어도 다시 만들어 결과를 잃지 않게 함)
    final = st.final or {"report": "(제출 없음: 최대 턴 수 초과)", "issues": [], "unresolved": []}
    out.mkdir(parents=True, exist_ok=True)
    (out / "report.md").write_text(final["report"], encoding="utf-8")
    (out / "trace.md").write_text("\n".join(trace), encoding="utf-8")
    result = {"issues": final["issues"], "others": final.get("others", []), "rejections": st.rejections,
              "unresolved": final["unresolved"]}
    if categories:
        result.update(analyzed=len(kept), total=len(df), categories=final.get("categories", []),
                      submitted=st.final is not None)
    (out / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    for p in [filter_path, analyze_path]:
        if p.exists():
            (out / f"prompt_{p.name}").write_text(p.read_text(encoding="utf-8"), encoding="utf-8")

    rejected = st.submits - 1 if st.final else st.submits
    log = out_root / "run_log.csv"
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
    print(f"결과 저장: {rel(out)}/ (report.md, trace.md" + (", filter.xlsx" if judgments is not None else "") + ")")
    return st, judgments


# ───────── 경쟁 앱 비교 ─────────
COMPARE_SYSTEM = """\
당신은 여러 앱의 리뷰 분석 결과를 비교하는 에이전트입니다. 아래 도구를 사용해 작업하세요.
- submit_mapping: 앱마다 다른 이름으로 정리된 문제 카테고리를 공통 카테고리로 묶어 제출합니다.
  검증을 통과하면 공통 카테고리별 언급 비율 표가 돌아옵니다. 빠지거나 중복된 카테고리가 있으면 반려됩니다.
- submit_comparison: 비율 표를 근거로 기준 앱의 약점(weaknesses), 강점(strengths), 시사점(implications)을 제출합니다. 반드시 이 도구로 제출해야 작업이 끝납니다.

아래는 이번 작업의 지침입니다.
"""


def compare_tools(labels):
    insight = {"type": "array", "items": {"type": "object",
               "properties": {"공통카테고리": {"type": "string"}, "내용": {"type": "string"}},
               "required": ["공통카테고리", "내용"]}}
    return [{
        "name": "submit_mapping",
        "description": "앱별 카테고리를 공통 카테고리로 묶습니다. 모든 앱의 모든 카테고리가 정확히 한 번씩 들어가야 합니다.",
        "input_schema": {"type": "object", "properties": {"groups": {"type": "array", "items": {
            "type": "object",
            "properties": {"공통카테고리": {"type": "string"},
                           "항목": {"type": "array", "items": {"type": "object", "properties": {
                               "앱": {"type": "string", "enum": labels}, "카테고리": {"type": "string"}},
                               "required": ["앱", "카테고리"]}}},
            "required": ["공통카테고리", "항목"]}}}, "required": ["groups"]},
    }, {
        "name": "submit_comparison",
        "description": "비율 표를 근거로 기준 앱의 약점, 강점, 시사점을 제출합니다.",
        "input_schema": {"type": "object", "properties": {
            "weaknesses": {**insight, "description": "약점"}, "strengths": {**insight, "description": "강점"},
            "implications": {"type": "array", "items": {"type": "string"}, "description": "시사점"}},
            "required": ["weaknesses", "strengths", "implications"]},
    }]


class CompareStage:
    """apps: [{"label", "analyzed", "categories": {이름: set(리뷰번호)}}], 첫 번째가 기준 앱"""

    def __init__(self, apps):
        self.apps = apps
        self.base = apps[0]["label"]
        self.table = None      # [{"공통카테고리", "비율": {앱: 비율}, "차이"}]
        self.groups = None
        self.submits = 0
        self.rejections = []
        self.final = None

    def verify_mapping(self, groups):
        problems, seen = [], {}
        want = {(a["label"], c) for a in self.apps for c in a["categories"]}
        names = [str(g.get("공통카테고리", "")).strip() for g in groups]
        if not groups:
            return ["groups가 비어 있습니다."]
        if "" in names or len(set(names)) != len(names):
            problems.append("공통카테고리 이름이 비어 있거나 중복됩니다.")
        for g in groups:
            for it in g.get("항목") or []:
                key = (it.get("앱"), str(it.get("카테고리", "")).strip())
                if key not in want:
                    problems.append(f"없는 카테고리입니다: {key[0]}의 '{key[1]}'")
                seen[key] = seen.get(key, 0) + 1
        dup = [f"{a}의 '{c}'" for (a, c), n in seen.items() if n > 1 and (a, c) in want]
        if dup:
            problems.append(f"두 번 이상 묶인 카테고리가 있습니다: {dup}")
        miss = [f"{a}의 '{c}'" for a, c in sorted(want - set(seen))]
        if miss:
            problems.append(f"묶이지 않은 카테고리가 있습니다: {miss}")
        return problems

    def compute(self, groups):
        """공통 카테고리별로 앱마다 '그 문제를 언급한 앱 관련 리뷰 수 / 앱 관련 리뷰 수'. 한 리뷰는 한 번만 셈.
        그 앱의 카테고리가 하나도 묶이지 않았으면 비율은 None (불만이 없는 건지 분석에서 안 잡힌 건지 알 수 없음).
        차이는 기준 앱과 경쟁 앱 하나 이상에 모두 있는 카테고리만, 그 카테고리가 있는 경쟁 앱 평균과 비교"""
        rows = []
        for g in groups:
            share = {}
            for a in self.apps:
                mine = [str(it["카테고리"]).strip() for it in g["항목"] if it["앱"] == a["label"]]
                ids = set().union(*(a["categories"].get(c, set()) for c in mine))
                share[a["label"]] = len(ids) / a["analyzed"] if mine and a["analyzed"] else None
            others = [share[a["label"]] for a in self.apps[1:] if share[a["label"]] is not None]
            common = share[self.base] is not None and bool(others)
            rows.append({"공통카테고리": g["공통카테고리"].strip(), "비율": share,
                         "차이": share[self.base] - sum(others) / len(others) if common else None})
        return sorted(rows, key=lambda r: (r["차이"] is None, -(r["차이"] or 0)))

    def table_md(self):
        labels = [a["label"] for a in self.apps]
        pct = lambda v: "-" if v is None else f"{v:.0%}"
        lines = ["| 공통 카테고리 | " + " | ".join(labels) + " | 차이 (기준 − 경쟁 평균) |",
                 "|---" * (len(labels) + 2) + "|"]
        for r in self.table:
            gap = "한 앱에서만 나옴" if r["차이"] is None else f"{r['차이'] * 100:+.0f}%p"
            lines.append(f"| {r['공통카테고리']} | " + " | ".join(pct(r["비율"][l]) for l in labels) + f" | {gap} |")
        lines.append("\n'-'는 그 앱에서 해당 카테고리가 나오지 않은 경우입니다 (불만이 없는 건지 분석에서 잡히지 않은 건지 알 수 없음).")
        return "\n".join(lines)

    def run_tool(self, name, args):
        if name == "submit_mapping":
            groups, problems = as_items(args.get("groups"), "groups")
            for g in groups:
                g["항목"], bad = as_items(g.get("항목"), f"[{g.get('공통카테고리')}] 항목")
                problems += bad
            problems = problems or self.verify_mapping(groups)
            if problems:
                self.rejections.append(problems)
                return "검증 실패로 반려되었습니다. 아래 문제를 고쳐 다시 제출하세요.\n- " + "\n- ".join(problems)
            self.groups, self.table = groups, self.compute(groups)
            return ("묶은 결과를 받았습니다. 공통 카테고리별 언급 비율입니다. "
                    "이 표를 근거로 submit_comparison을 제출하세요.\n" + self.table_md())
        if name == "submit_comparison":
            if self.table is None:
                return "먼저 submit_mapping으로 공통 카테고리를 제출하세요."
            self.submits += 1
            weak, bad_w = as_items(args.get("weaknesses"), "weaknesses")
            strong, bad_s = as_items(args.get("strengths"), "strengths")
            notes = args.get("implications")
            notes = [notes] if isinstance(notes, str) else [str(n) for n in notes or []]
            args = {"약점": weak, "강점": strong, "시사점": notes}
            problems = bad_w + bad_s + self.verify_insights(args)
            if problems and self.submits < MAX_SUBMITS:
                self.rejections.append(problems)
                return "검증 실패로 반려되었습니다. 아래 문제를 고쳐 다시 제출하세요.\n- " + "\n- ".join(problems)
            self.final = {"약점": args.get("약점") or [], "강점": args.get("강점") or [],
                          "시사점": args.get("시사점") or [], "unresolved": problems}
            return "제출이 완료되었습니다."
        return f"알 수 없는 도구: {name}"

    def verify_insights(self, args):
        """약점은 기준 앱 비율이 더 높은, 강점은 더 낮은 공통 카테고리여야 함 (표와 어긋난 주장 방지)"""
        diff = {r["공통카테고리"]: r["차이"] for r in self.table}
        problems = []
        for kind, sign in (("약점", 1), ("강점", -1)):
            for it in args.get(kind) or []:
                name = str(it.get("공통카테고리", "")).strip()
                if name not in diff:
                    problems.append(f"[{kind}] 표에 없는 공통 카테고리입니다: {name}")
                elif diff[name] is None:
                    problems.append(f"[{kind}] '{name}'은 한 앱에서만 나온 카테고리라 약점·강점으로 판단할 수 없습니다.")
                elif diff[name] * sign <= 0:
                    problems.append(f"[{kind}] '{name}'은 기준 앱 비율이 경쟁 앱 평균보다 "
                                    f"{'높지' if sign > 0 else '낮지'} 않습니다 (차이 {diff[name] * 100:+.0f}%p).")
        if not args.get("시사점"):
            problems.append("시사점이 비어 있습니다.")
        return problems

    @property
    def done(self):
        return self.final is not None

    def summarize(self, name, args, out):
        if name == "submit_mapping":
            return f"공통 카테고리 {len(args.get('groups') or [])}개 제출 → {out.splitlines()[0]}"
        if name == "submit_comparison":
            return f"비교 결과 제출 → {out.splitlines()[0]}"
        return out


def short_label(title, app_id):
    """'타다 - 새로운 이동 기준 (...)' → '타다'. 표 머리글과 도구 입력에 쓰는 짧은 이름"""
    return re.split(r"\s+-\s+|\s*[(\[:|]", title or "")[0].strip() or app_id


def prepare_app(app, start, end, version, client):
    """앱 하나의 리뷰를 기간으로 수집하고 분석. 이미 있으면 재사용 (같은 기간·버전이면 비용이 다시 들지 않음)"""
    app_id = resolve_app_id(app)
    root = OUT_DIR / "compare" / f"{start}_{end}" / app_id
    info_path, review_path, result_path = root / "app.json", root / REVIEW_FILE, root / version / "result.json"
    if info_path.exists():
        info = json.loads(info_path.read_text(encoding="utf-8"))
    else:
        title = app_title(app_id)
        print(f"[{title}] {start} ~ {end} 리뷰 수집 중...")
        found, used = collect_period(app_id, start, end, review_path)
        info = {"appId": app_id, "title": title, "found": found, "used": used}
        root.mkdir(parents=True, exist_ok=True)
        info_path.write_text(json.dumps(info, ensure_ascii=False, indent=1), encoding="utf-8")
    info["label"] = short_label(info["title"], app_id)
    if info["used"] < PERIOD_SKIP:
        print(f"[{info['label']}] 기간 안 리뷰가 {info['used']}개뿐이라 비교에서 뺍니다. "
              f"--from으로 기간을 늘려 보세요. (최소 {PERIOD_SKIP}개)")
        return None
    note = f" (기간 안 {info['found']}개 중 고르게 {info['used']}개)" if info["found"] > info["used"] else ""
    print(f"[{info['label']}] 리뷰 {info['used']}개{note}")
    if not result_path.exists():
        run_version(version, client=client, review_path=review_path, out_root=root)
    result = json.loads(result_path.read_text(encoding="utf-8"))
    if not result.get("submitted") or not result.get("categories"):
        print(f"[{info['label']}] 분석 결과가 없어 비교에서 뺍니다. ({rel(result_path.parent)} 확인)")
        return None
    reviews = pd.read_excel(review_path)
    return {**info, "root": root, "analyzed": result["analyzed"], "total": result["total"],
            "categories": {str(c["카테고리"]).strip(): set(c["관련리뷰"]) for c in result["categories"]},
            "issues": result["issues"], "star_mean": reviews["별점"].mean(),
            "star1": (reviews["별점"] == 1).mean()}


def compare_context(apps):
    """비교 에이전트에게 주는 데이터: 앱별 카테고리와 관련 리뷰 수, 기준 앱의 개선안 우선순위"""
    lines = [f"기준 앱: {apps[0]['label']} / 경쟁 앱: {', '.join(a['label'] for a in apps[1:])}", ""]
    for a in apps:
        warn = " (표본 부족: 참고 수준)" if a["analyzed"] < PERIOD_MIN else ""
        lines.append(f"[{a['label']}] 앱 관련 리뷰 {a['analyzed']}개{warn}")
        lines += [f"- {c}: 관련 리뷰 {len(ids)}개" for c, ids in a["categories"].items()]
        lines.append("")
    lines.append(f"[{apps[0]['label']} 개선안 우선순위]")
    lines += [f"{i}. {it['개선안']} ({it['카테고리']}, {QUADRANTS.get(it.get('구분'), '-')})"
              for i, it in enumerate(apps[0]["issues"], 1)]
    return "\n".join(lines)


def render_comparison(apps, st, start, end, version):
    base, labels = apps[0]["label"], [a["label"] for a in apps]
    lines = [f"# 경쟁 앱 리뷰 비교: {base} vs {', '.join(labels[1:])}", "",
             f"- 기간: {start} ~ {end} / 지침: {version} / 모델: {MODEL}", "",
             "## 1. 앱별 개요", "",
             "| 앱 | 분석한 리뷰 | 앱 관련 리뷰 | 평균 별점 | 1점 비율 | 비고 |", "|---|---|---|---|---|---|"]
    for a in apps:
        notes = []
        if a["found"] > a["used"]:
            notes.append(f"기간 안 {a['found']}개 중 고르게 {a['used']}개 사용")
        if a["analyzed"] < PERIOD_MIN:
            notes.append(f"표본 부족 ({PERIOD_MIN}개 미만, 참고 수준)")
        lines.append(f"| {a['label']} | {a['total']} | {a['analyzed']} | {a['star_mean']:.1f} | "
                     f"{a['star1']:.0%} | {', '.join(notes) or '-'} |")
    diff = {r["공통카테고리"]: r for r in st.table}
    fmt = lambda name: (f"{base} {diff[name]['비율'][base]:.0%} / 경쟁 평균 "
                        f"{diff[name]['비율'][base] - diff[name]['차이']:.0%}") if name in diff and diff[name]["차이"] is not None else ""
    lines += ["", "## 2. 공통 카테고리별 언급 비율", "",
              "비율 = 그 문제를 언급한 앱 관련 리뷰 수 ÷ 앱 관련 리뷰 수. 앱마다 따로 분석한 카테고리를 공통 카테고리로 묶어 계산했습니다.", "",
              st.table_md()]
    for title, key in ((f"3. {base}의 약점 (경쟁 앱보다 불만이 많은 문제)", "약점"),
                       (f"4. {base}의 강점 (경쟁 앱보다 불만이 적은 문제)", "강점")):
        lines += ["", f"## {title}", ""]
        lines += [f"- **{it['공통카테고리']}** ({fmt(it['공통카테고리'].strip())}): {it['내용']}"
                  for it in st.final[key]] or ["- 없음"]
    lines += ["", "## 5. 시사점", ""] + [f"- {s}" for s in st.final["시사점"]]
    lines += ["", "## 부록 1: 카테고리 묶음", "", "| 공통 카테고리 | 묶인 카테고리 |", "|---|---|"]
    for g in st.groups:
        lines.append(f"| {g['공통카테고리']} | " + "<br>".join(f"{i['앱']}: {i['카테고리']}" for i in g["항목"]) + " |")
    lines += base_priority_lines(apps[0])
    return "\n".join(lines) + "\n"


def base_priority_lines(app):
    """시사점이 번호로 가리키는 기준 앱 개선안 목록 (같은 기간 리뷰로 분석한 결과)"""
    lines = ["", f"## 부록 2: {app['label']} 개선안 우선순위 (같은 기간 리뷰 분석)", "",
             f"전체 보고서: {rel(app['root'] / COMPARE_VERSION / 'report.md').as_posix()}", "",
             "| 순위 | 개선안 | 구분 | 근거 리뷰 수 |", "|---|---|---|---|"]
    for i, it in enumerate(app["issues"], 1):
        lines.append(f"| {i} | {it['개선안']} | {QUADRANTS.get(it.get('구분'), '-')} | {len(set(it['근거리뷰']))} |")
    return lines


def run_compare(apps_in, base=BASE_APP, start=None, end=None, version=COMPARE_VERSION, client=None):
    vdir = PROMPT_DIR / version
    opts = load_harness_options(vdir)
    if not opts["categories"] or not (vdir / "compare.txt").exists():
        print(f"{version}은 비교를 지원하지 않습니다. harness.json에 \"categories\": true, compare.txt가 있는 버전을 쓰세요.")
        sys.exit(1)
    if client is None:
        check_api_key()
    end = end or datetime.now().date()
    start = start or end - timedelta(days=PERIOD_DAYS - 1)
    base_app = prepare_app(base, start, end, version, client)
    rivals = [a for a in (prepare_app(x, start, end, version, client) for x in apps_in) if a]
    if not base_app or not rivals:
        print("기준 앱과 경쟁 앱이 하나 이상 있어야 비교할 수 있습니다.")
        sys.exit(1)
    apps = [base_app, *rivals]
    labels = [a["label"] for a in apps]
    if len(set(labels)) < len(labels):     # 짧은 이름이 겹치면 패키지명으로 구분
        for a in apps:
            a["label"] = a["appId"]
    trace = [f"# 경쟁 앱 비교 작업 기록", f"- 모델: {MODEL}", f"- 실행: {datetime.now():%Y-%m-%d %H:%M}"]
    st = CompareStage(apps)
    system = COMPARE_SYSTEM + (vdir / "compare.txt").read_text(encoding="utf-8").strip() + "\n\n" + compare_context(apps)
    run_stage("비교", system, compare_tools([a["label"] for a in apps]), st, trace, client)
    if not st.final:
        print("비교 결과를 제출받지 못했습니다 (최대 턴 수 초과).")
        sys.exit(1)
    out = OUT_DIR / "compare" / f"{start}_{end}"
    name = "comparison_" + "_vs_".join(a["appId"].split(".")[-1] for a in apps)
    (out / f"{name}.md").write_text(render_comparison(apps, st, start, end, version), encoding="utf-8")
    (out / f"{name}_trace.md").write_text("\n".join(trace), encoding="utf-8")
    (out / f"{name}.json").write_text(json.dumps({
        "apps": [{k: a[k] for k in ("appId", "label", "found", "used", "total", "analyzed")} for a in apps],
        "groups": st.groups, "table": st.table, "insights": st.final, "rejections": st.rejections,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n비교 결과 저장: {rel(out / name)}.md")
    return st


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
    if sys.argv[1] == "compare":
        args, opt = sys.argv[2:], {}
        for flag in ("--base", "--from", "--to", "--version"):
            if flag in args:
                i = args.index(flag)
                if i + 1 >= len(args):
                    print(f"{flag} 뒤에 값을 넣으세요.")
                    sys.exit(1)
                opt[flag] = args[i + 1]
                del args[i:i + 2]
        if not args:
            print('사용법: python run_prompt.py compare "경쟁 앱1" "경쟁 앱2" [--from 2026-09-01 --to 2026-09-30]')
            sys.exit(1)
        day = lambda s: datetime.strptime(s, "%Y-%m-%d").date() if s else None
        run_compare(args, base=opt.get("--base", BASE_APP), start=day(opt.get("--from")),
                    end=day(opt.get("--to")), version=opt.get("--version", COMPARE_VERSION))
        return
    filter_from = None
    if "--filter-from" in sys.argv:
        i = sys.argv.index("--filter-from")
        filter_from = sys.argv[i + 1] if i + 1 < len(sys.argv) else None
    run_version(sys.argv[1], filter_from=filter_from)


if __name__ == "__main__":
    main()