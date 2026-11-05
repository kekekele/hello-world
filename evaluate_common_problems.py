"""Compare CSV issues with Markdown issue titles and export CSV reports."""

from __future__ import annotations

import argparse
import csv
from dataclasses import asdict, dataclass
import json
from pathlib import Path
import re
import sys


@dataclass(frozen=True)
class Problem:
    id: str
    text: str


@dataclass(frozen=True)
class Judgment:
    excel_id: str
    markdown_ids: list[str]
    reason: str


@dataclass
class GroupResult:
    name: str
    excel_problems: list[Problem]
    markdown_problems: list[Problem]
    judgments: list[Judgment]
    markdown_only: dict[str, str]


Groups = dict[str, list[Problem]]
HEADERS = ("序号", "分组", "共性特征", "问题")
GROUP_HEADING = re.compile(r"^ {0,3}##[ \t]+(.+?)\s*#*\s*$")
ISSUE_HEADING = re.compile(r"^ {0,3}###[ \t]+\d+[.、)][ \t]*(.+?)\s*#*\s*$")
FENCE = re.compile(r"^ {0,3}(`{3,}|~{3,})(.*)$")


def read_csv(path: Path) -> Groups:
    if path.suffix.lower() != ".csv":
        raise ValueError("参考问题请使用 .csv 格式。")
    with path.open(encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file, strict=True)
        first_row = next(reader, None)
        if first_row is None:
            raise ValueError("CSV 缺少表头。")
        headers = [value.strip() for value in first_row]
        for header in HEADERS:
            if headers.count(header) != 1:
                raise ValueError(f"第一行必须且只能有一列“{header}”。")
        columns = {header: headers.index(header) for header in HEADERS}
        groups: Groups = {}
        count = 0
        for row in reader:
            if not any(value.strip() for value in row):
                continue
            if len(row) != len(headers):
                raise ValueError(f"CSV 第{reader.line_num}行：字段数与表头不一致。")
            group, issue = row[columns["分组"]].strip(), row[columns["问题"]].strip()
            if not group or not issue:
                raise ValueError(f"CSV 第{reader.line_num}行：分组和问题不能为空。")
            count += 1
            groups.setdefault(group, []).append(Problem(f"E{count:06d}", issue))
        return groups


def read_markdown(path: Path) -> Groups:
    content = path.read_text(encoding="utf-8-sig")
    groups: Groups = {}
    group: str | None = None
    count = 0
    fence_char = ""
    fence_length = 0
    for line_number, line in enumerate(content.splitlines(), 1):
        fence = FENCE.match(line)
        if fence_char:
            if (fence and fence.group(1)[0] == fence_char
                    and len(fence.group(1)) >= fence_length
                    and not fence.group(2).strip()):
                fence_char = ""
            continue
        if fence:
            fence_char, fence_length = fence.group(1)[0], len(fence.group(1))
            continue
        group_match = GROUP_HEADING.match(line)
        if group_match:
            group = group_match.group(1).strip()
            if not group:
                raise ValueError(f"MD 第{line_number}行：分组名不能为空。")
            groups.setdefault(group, [])
            continue
        issue_match = ISSUE_HEADING.match(line)
        if issue_match:
            if group is None:
                raise ValueError(f"MD 第{line_number}行：问题标题必须位于 ## 分组名之下。")
            title = issue_match.group(1).strip()
            if not title:
                raise ValueError(f"MD 第{line_number}行：问题标题不能为空。")
            count += 1
            groups[group].append(Problem(f"M{count:06d}", title))
        elif re.match(r"^ {0,3}##(?:#)?(?:\s|$)", line):
            raise ValueError(f"MD 第{line_number}行：请使用 ## 分组名 或 ### 1. 问题标题。")
    if fence_char:
        raise ValueError("MD 中的代码围栏未闭合。")
    if content.strip() and not groups:
        raise ValueError("MD 中没有 ## 分组名。")
    return groups


def call_model(system_prompt: str, user_prompt: str) -> str:
    """接入已有模型函数，返回模型回复正文的 JSON 字符串。"""
    raise NotImplementedError("请在 call_model 中接入已有模型调用函数。")


def json_object(value: object) -> dict[str, object]:
    if not isinstance(value, dict) or not all(isinstance(key, str) for key in value):
        raise ValueError("模型响应中应为 JSON 对象的位置出现了其他类型。")
    return dict(value)


def json_array(value: object) -> list[object]:
    if not isinstance(value, list):
        raise ValueError("模型响应中应为数组的位置出现了其他类型。")
    return list(value)


def nonempty_string(value: object) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError("模型响应中的 ID、原因或回复正文必须为非空字符串。")
    return value


def parse_response(
    raw: str, excel: list[Problem], markdown: list[Problem],
) -> tuple[list[Judgment], dict[str, str]]:
    response = json_object(json.loads(nonempty_string(raw)))
    if set(response) != {"excel_results", "markdown_only"}:
        raise ValueError("模型响应必须包含 excel_results 和 markdown_only。")
    excel_ids = {problem.id for problem in excel}
    markdown_ids = {problem.id for problem in markdown}
    seen_excel: set[str] = set()
    matched_markdown: set[str] = set()
    judgments: list[Judgment] = []
    for raw_item in json_array(response["excel_results"]):
        item = json_object(raw_item)
        if set(item) != {"excel_id", "markdown_ids", "reason"}:
            raise ValueError("参考问题判定项字段不正确。")
        excel_id = nonempty_string(item["excel_id"])
        matched_ids = [nonempty_string(value) for value in json_array(item["markdown_ids"])]
        if excel_id not in excel_ids or excel_id in seen_excel:
            raise ValueError(f"参考问题 ID 不存在或重复：{excel_id}")
        if len(matched_ids) != len(set(matched_ids)) or not set(matched_ids) <= markdown_ids:
            raise ValueError(f"{excel_id} 的 Markdown ID 不存在或重复。")
        judgments.append(Judgment(excel_id, matched_ids, nonempty_string(item["reason"])))
        seen_excel.add(excel_id)
        matched_markdown.update(matched_ids)
    if seen_excel != excel_ids:
        raise ValueError("模型没有逐一判定全部参考问题。")
    markdown_only: dict[str, str] = {}
    for raw_item in json_array(response["markdown_only"]):
        item = json_object(raw_item)
        if set(item) != {"markdown_id", "reason"}:
            raise ValueError("MD 独有项字段不正确。")
        markdown_id = nonempty_string(item["markdown_id"])
        if markdown_id not in markdown_ids or markdown_id in markdown_only:
            raise ValueError(f"MD 独有项 ID 不存在或重复：{markdown_id}")
        markdown_only[markdown_id] = nonempty_string(item["reason"])
    if set(markdown_only) != markdown_ids - matched_markdown:
        raise ValueError("MD 独有项必须完整列出所有未命中参考问题的标题，且不能包含已命中项。")
    return judgments, markdown_only


def evaluate_group(
    name: str, excel: list[Problem], markdown: list[Problem], prompt: str,
) -> GroupResult:
    if not markdown:
        judgments = [Judgment(p.id, [], "该分组下没有 Markdown 问题标题。") for p in excel]
        return GroupResult(name, excel, markdown, judgments, {})
    if not excel:
        markdown_only = {p.id: "CSV 的同一分组下没有问题。" for p in markdown}
        return GroupResult(name, excel, markdown, [], markdown_only)
    user_prompt = json.dumps({
        "group_name": name,
        "excel_problems": [asdict(problem) for problem in excel],
        "markdown_problems": [asdict(problem) for problem in markdown],
    }, ensure_ascii=False)
    judgments, markdown_only = parse_response(call_model(prompt, user_prompt), excel, markdown)
    return GroupResult(name, excel, markdown, judgments, markdown_only)


def evaluate(excel: Groups, markdown: Groups, prompt: str) -> list[GroupResult]:
    results = []
    for name in dict.fromkeys([*excel, *markdown]):
        results.append(evaluate_group(name, excel.get(name, []), markdown.get(name, []), prompt))
    return results


def calculate_hit_rate(results: list[GroupResult]) -> tuple[int, int, float | None]:
    total = sum(len(result.excel_problems) for result in results)
    hits = sum(bool(judgment.markdown_ids) for result in results for judgment in result.judgments)
    return total, hits, hits / total if total else None


def csv_value(value: object) -> object:
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def write_csv(path: Path, rows: list[list[object]]) -> None:
    with path.open("x", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        for row in rows:
            writer.writerow([csv_value(value) for value in row])


def export_report(results: list[GroupResult], directory: Path) -> None:
    details: list[list[object]] = [["分组名", "CSV中问题", "Markdown有没有命中", "原因"]]
    extra: list[list[object]] = [["分组名", "Markdown问题", "原因"]]
    for result in results:
        judgments = {judgment.excel_id: judgment for judgment in result.judgments}
        for problem in result.excel_problems:
            judgment = judgments[problem.id]
            details.append([
                result.name, problem.text, "是" if judgment.markdown_ids else "否", judgment.reason,
            ])
        for problem in result.markdown_problems:
            if problem.id in result.markdown_only:
                extra.append([result.name, problem.text, result.markdown_only[problem.id]])
    total, hits, rate = calculate_hit_rate(results)
    summary: list[list[object]] = [
        ["CSV问题总数", "命中问题数", "命中率"],
        [total, hits, f"{rate:.2%}" if rate is not None else "N/A"],
    ]
    directory.mkdir(parents=True, exist_ok=False)
    write_csv(directory / "问题命中.csv", details)
    write_csv(directory / "MD独有问题.csv", extra)
    write_csv(directory / "命中指标.csv", summary)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", type=Path, required=True, help="参考问题 CSV 文件路径")
    parser.add_argument("--markdown", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True, help="输出三份 CSV 报告的新目录")
    args = parser.parse_args(argv)
    try:
        if args.output.exists():
            raise ValueError("output 必须是尚不存在的新目录，不覆盖已有文件或目录。")
        excel = read_csv(args.csv)
        markdown = read_markdown(args.markdown)
        prompt = Path(__file__).with_name("共性问题评估Prompt.txt").read_text(encoding="utf-8-sig")
        if not prompt.strip():
            raise ValueError("Prompt 文件不能为空。")
        results = evaluate(excel, markdown, prompt)
        export_report(results, args.output)
        total, hits, rate = calculate_hit_rate(results)
        rate_text = f"{rate:.2%}" if rate is not None else "N/A"
        print(f"命中率：{rate_text}（{hits}/{total}）；报告：{args.output.resolve()}")
        return 0
    except (OSError, ValueError, RuntimeError, csv.Error) as error:
        print(f"评估失败：{error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
