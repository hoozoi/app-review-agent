"""데모 페이지(index.html)에 들어갈 데이터를 실제 실행 결과에서 만듦

  python demo/build_data.py
  → review_agent/outputs/v5 (카카오 T 분석)와 outputs/compare (우버 비교) 결과를 읽어
    demo/index.html의 <script id="demo-data"> 안에 넣음 (파일 하나로 어디서든 열리게)
"""
import json
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "review_agent" / "outputs"
V5 = OUT / "v5"
COMPARE = OUT / "compare" / "2026-09-24_2026-10-23" / "comparison_taxi_vs_ubercab.json"
QUADRANTS = {1: "바로 실행", 2: "중장기 과제", 3: "여유 있을 때", 4: "보류"}


def summary_lines(report):
    part = report.split("## 1. 요약", 1)[1].split("\n## ", 1)[0]
    return [l[2:].strip() for l in part.splitlines() if l.startswith("- ")]


def main():
    result = json.loads((V5 / "result.json").read_text(encoding="utf-8"))
    report = (V5 / "report.md").read_text(encoding="utf-8")
    judged = pd.read_excel(V5 / "filter.xlsx")["AI 판단"].value_counts()
    trace = (V5 / "trace.md").read_text(encoding="utf-8")
    analyze_trace = trace.split("## 2단계 분석", 1)[1]
    submits = analyze_trace.count("`submit_report`")
    rejected = analyze_trace.count("반려되었습니다")

    issues = [{
        "rank": i, "title": it["개선안"], "category": it["카테고리"], "quadrant": it["구분"],
        "quadrantName": QUADRANTS[it["구분"]], "evidence": len(set(it["근거리뷰"])),
        "severity": it["심각도"], "difficulty": it["난이도"], "method": it["개선방법"],
        "effect": it["기대효과"], "refs": sorted(set(it["근거리뷰"])),
    } for i, it in enumerate(result["issues"], 1)]

    cmp = json.loads(COMPARE.read_text(encoding="utf-8"))
    labels = [a["label"] for a in cmp["apps"]]
    names = {"Uber": "우버"}
    rows = [{"name": r["공통카테고리"],
             "values": [r["비율"][l] for l in labels],
             "diff": r["차이"]} for r in cmp["table"]]

    data = {
        "app": {"name": "카카오 T", "store": "Google Play", "appId": "com.kakao.taxi"},
        "steps": {"collected": result["total"], "kept": int(judged.get("O", 0)), "excluded": int(judged.get("X", 0)),
                  "categories": len(result["categories"]), "issues": len(issues),
                  "submits": submits, "rejected": rejected},
        "summary": summary_lines(report),
        "categories": sorted(({"name": c["카테고리"], "count": len(set(c["관련리뷰"]))} for c in result["categories"]),
                             key=lambda c: -c["count"]),
        "issues": issues,
        "compare": {
            "period": re.sub(r"_", " ~ ", COMPARE.parent.name),
            "apps": [{"name": names.get(a["label"], a["label"]), "analyzed": a["analyzed"]} for a in cmp["apps"]],
            "rows": rows,
            "weaknesses": cmp["insights"]["약점"],
            "strengths": cmp["insights"]["강점"],
        },
    }
    path = Path(__file__).resolve().parent / "index.html"
    page = path.read_text(encoding="utf-8")
    block = ('<script id="demo-data">/* build_data.py가 실제 실행 결과로 채움. 직접 고치지 마세요. */\n'
             "window.DEMO_DATA = " + json.dumps(data, ensure_ascii=False) + ";</script>")
    page = re.sub(r'<script id="demo-data">.*?</script>', lambda m: block, page, count=1, flags=re.S)
    path.write_text(page, encoding="utf-8")
    print(f"저장: {path.relative_to(ROOT)} (개선안 {len(issues)}개, 비교 카테고리 {len(rows)}개)")


if __name__ == "__main__":
    main()
