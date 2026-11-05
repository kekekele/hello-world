"""Evaluate grouped operator issues using an existing model-call function."""

from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass
import json
from pathlib import Path
import re
import sys
from typing import Callable
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet
from scipy.optimize import linear_sum_assignment


@dataclass(frozen=True)
class Problem:
    id: str
    source_text: str
    source_locator: str
    sequence: str = ""
    common_features: str = ""


@dataclass(frozen=True)
class Candidate:
    reference_id: str
    predicted_id: str
    label: str
    reference_evidence: str
    predicted_evidence: str
    reason: str
    differences: list[str]


Groups = dict[str, list[Problem]]
Judge = Callable[[dict[str, object], str], object]
HEADERS = ("序号", "分组", "共性特征", "问题")
GROUP_HEADING = re.compile(r"^ {0,3}##[ \t]+(.+?)\s*#*\s*$")
ISSUE_HEADING = re.compile(r"^ {0,3}###[ \t]+(\d+)[.、)][ \t]*(.+?)\s*#*\s*$")
FENCE = re.compile(r"^ {0,3}(`{3,}|~{3,})(.*)$")


def text(value: object) -> str:
    return "" if value is None else str(value).strip()


def read_excel(path: Path, sheet: str | None = None) -> Groups:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Excel 必须为 .xlsx/.xlsm；旧版 .xls 请先另存为 .xlsx。")
    workbook = load_workbook(path, data_only=False)
    try:
        if sheet is not None and sheet not in workbook.sheetnames:
            raise ValueError(f"找不到工作表 {sheet!r}；可用：{workbook.sheetnames}")
        worksheet = workbook[sheet] if sheet is not None else workbook.worksheets[0]
        header_values = [text(cell.value).lstrip("\ufeff") for cell in worksheet[1]]
        for header in HEADERS:
            if header_values.count(header) != 1:
                raise ValueError(f"{worksheet.title}: 第一行必须且只能有一列“{header}”。")
        columns = {header: header_values.index(header) + 1 for header in HEADERS}
        merged: dict[tuple[int, int], object] = {}
        for area in worksheet.merged_cells.ranges:
            relevant = [column for column in columns.values()
                        if area.min_col <= column <= area.max_col]
            if not relevant:
                continue
            if area.min_col != area.max_col:
                raise ValueError(f"{worksheet.title}: 不支持跨列合并 {area}。")
            header = next(name for name, column in columns.items()
                          if column == area.min_col)
            if area.min_row == 1 or header == "问题":
                raise ValueError(f"{worksheet.title}: 表头和问题列不允许合并：{area}。")
            value = worksheet.cell(area.min_row, area.min_col).value
            for row in range(area.min_row, area.max_row + 1):
                merged[row, area.min_col] = value
        groups: Groups = {}
        count = 0
        for row in range(2, worksheet.max_row + 1):
            raw = {name: worksheet.cell(row, column).value
                   for name, column in columns.items()}
            if all(not text(value) for value in raw.values()):
                continue
            values = {name: merged.get((row, column), raw[name])
                      for name, column in columns.items()}
            for name, value in values.items():
                if isinstance(value, str) and value.startswith("="):
                    raise ValueError(f"{worksheet.title}!第{row}行 {name}: 不允许公式。")
            group, issue = text(values["分组"]), text(values["问题"])
            if not group or not issue:
                raise ValueError(
                    f"{worksheet.title}!第{row}行: 分组和问题不能为空；"
                    "仅真实合并单元格会自动展开，普通空白分组不会向下填充。"
                )
            count += 1
            features = text(values["共性特征"])
            source = f"共性特征：{features}\n问题：{issue}" if features else f"问题：{issue}"
            groups.setdefault(group, []).append(Problem(
                f"E{count:06d}", source,
                f"{path.name}:{worksheet.title}!第{row}行",
                text(values["序号"]), features,
            ))
        if not groups:
            raise ValueError("Excel 没有问题记录。")
        return groups
    finally:
        workbook.close()


def read_markdown(path: Path) -> Groups:
    groups: Groups = {}
    group: str | None = None
    title: str | None = None
    sequence = ""
    start = 0
    body: list[str] = []
    count = 0
    fence_char = ""
    fence_length = 0

    def finish() -> None:
        nonlocal title, count
        if title is None:
            return
        if group is None:
            raise ValueError(f"{path.name}:{start}: 问题没有所属分组。")
        count += 1
        detail = "\n".join(body).strip()
        source = f"问题标题：{title}"
        if detail:
            source += f"\n详细解释：\n{detail}"
        groups[group].append(Problem(
            f"M{count:06d}", source, f"{path.name}:第{start}行", sequence,
        ))
        title = None
        body.clear()

    for line_number, line in enumerate(
        path.read_text(encoding="utf-8-sig").splitlines(), 1
    ):
        fence = FENCE.match(line)
        if fence_char:
            body.append(line)
            if (fence and fence.group(1)[0] == fence_char
                    and len(fence.group(1)) >= fence_length
                    and not fence.group(2).strip()):
                fence_char = ""
            continue
        if fence:
            if title is None:
                raise ValueError(f"{path.name}:{line_number}: 代码块必须位于问题解释内。")
            fence_char, fence_length = fence.group(1)[0], len(fence.group(1))
            body.append(line)
            continue
        group_match = GROUP_HEADING.match(line)
        if group_match:
            finish()
            group = group_match.group(1).strip()
            if not group:
                raise ValueError(f"{path.name}:{line_number}: 分组名不能为空。")
            groups.setdefault(group, [])
            continue
        issue_match = ISSUE_HEADING.match(line)
        if issue_match:
            finish()
            if group is None:
                raise ValueError(f"{path.name}:{line_number}: 问题必须放在 ## 分组名下面。")
            sequence, title = issue_match.groups()
            title = title.strip()
            start = line_number
            if not title:
                raise ValueError(f"{path.name}:{line_number}: 问题标题不能为空。")
            continue
        if re.match(r"^ {0,3}###(?:\s|$)", line):
            raise ValueError(f"{path.name}:{line_number}: 三级标题必须为 ### 1. 问题标题。")
        if re.match(r"^ {0,3}#(?:\s|$)", line):
            if group is not None:
                raise ValueError(f"{path.name}:{line_number}: 分组中不允许一级标题。")
            continue
        if title is not None:
            body.append(line)
        elif line.strip():
            raise ValueError(
                f"{path.name}:{line_number}: 文本不属于任何问题，"
                "请将适用范围放入问题标题或详细解释。"
            )
    if fence_char:
        raise ValueError(f"{path.name}: 问题解释中的代码围栏未闭合。")
    finish()
    if not groups:
        raise ValueError("MD 中没有 ## 分组名；零问题报告请保留对应的空分组标题。")
    return groups


def require_object(value: object, context: str) -> dict[str, object]:
    if not isinstance(value, dict) or not all(isinstance(key, str) for key in value):
        raise ValueError(f"{context}: 必须为 JSON 对象。")
    return {key: item for key, item in value.items()}


def require_string(value: object, context: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{context}: 必须为非空字符串。")
    return value


def string_list(value: object, context: str) -> list[str]:
    if not isinstance(value, list):
        raise ValueError(f"{context}: 必须为字符串数组。")
    return [require_string(item, context) for item in value]


def remap_groups(predictions: Groups, references: Groups, path: Path | None) -> Groups:
    if path is None:
        return predictions
    mapping = require_object(json.loads(path.read_text(encoding="utf-8-sig")), "分组映射")
    for name, target in mapping.items():
        target_name = require_string(target, f"分组映射 {name}")
        if name not in predictions or target_name not in references:
            raise ValueError(f"分组映射不存在的来源或目标：{name!r} -> {target_name!r}")
    result: Groups = {}
    for name, problems in predictions.items():
        target = require_string(mapping.get(name, name), "分组名称")
        result.setdefault(target, []).extend(problems)
    return result


def validate_response(value: object, group: dict[str, object]) -> list[Candidate]:
    response = require_object(value, "模型响应")
    expected_keys = {
        "group_id", "candidates",
        "reference_without_candidates", "prediction_without_candidates",
    }
    if set(response) != expected_keys or response["group_id"] != group["group_id"]:
        raise ValueError("模型响应字段或 group_id 不正确。")
    references = {
        require_string(item["id"], "参考 ID"): require_string(item["source_text"], "参考原文")
        for item in (
            require_object(entry, "参考问题")
            for entry in require_list(group["reference_problems"], "参考列表")
        )
    }
    predictions = {
        require_string(item["id"], "算法 ID"): require_string(item["source_text"], "算法原文")
        for item in (
            require_object(entry, "算法问题")
            for entry in require_list(group["predicted_problems"], "算法列表")
        )
    }
    candidates: list[Candidate] = []
    seen: set[tuple[str, str]] = set()
    candidate_keys = {
        "reference_id", "predicted_id", "label", "reference_evidence",
        "predicted_evidence", "reason", "differences",
    }
    for raw in require_list(response["candidates"], "candidates"):
        item = require_object(raw, "候选")
        if set(item) != candidate_keys:
            raise ValueError("候选字段不正确。")
        rid = require_string(item["reference_id"], "reference_id")
        pid = require_string(item["predicted_id"], "predicted_id")
        label = require_string(item["label"], "label")
        if rid not in references or pid not in predictions:
            raise ValueError(f"候选包含未知 ID：{rid}, {pid}")
        if (rid, pid) in seen or label not in {"full", "partial"}:
            raise ValueError(f"重复候选或非法标签：{rid}, {pid}, {label}")
        reference_evidence = require_string(item["reference_evidence"], "reference_evidence")
        predicted_evidence = require_string(item["predicted_evidence"], "predicted_evidence")
        if (reference_evidence not in references[rid]
                or predicted_evidence not in predictions[pid]):
            raise ValueError(f"证据不是对应问题的原文片段：{rid}, {pid}")
        differences = string_list(item["differences"], "differences")
        if label == "partial" and not differences:
            raise ValueError(f"partial 必须说明差异：{rid}, {pid}")
        candidates.append(Candidate(
            rid, pid, label, reference_evidence, predicted_evidence,
            require_string(item["reason"], "reason"), differences,
        ))
        seen.add((rid, pid))
    for field, universe, covered in (
        ("reference_without_candidates", set(references), {c.reference_id for c in candidates}),
        ("prediction_without_candidates", set(predictions), {c.predicted_id for c in candidates}),
    ):
        ids = string_list(response[field], field)
        if len(ids) != len(set(ids)) or set(ids) != universe - covered:
            raise ValueError(f"{field}: ID 未完整覆盖、重复或与候选冲突。")
    return candidates


def require_list(value: object, context: str) -> list[object]:
    if not isinstance(value, list):
        raise ValueError(f"{context}: 必须为数组。")
    return list(value)


def match_candidates(
    references: list[Problem], predictions: list[Problem], candidates: list[Candidate]
) -> list[Candidate]:
    if not references or not predictions:
        return []
    refs = sorted(references, key=lambda problem: problem.id)
    preds = sorted(predictions, key=lambda problem: problem.id)
    edges = {(candidate.reference_id, candidate.predicted_id): candidate
             for candidate in candidates}
    base = min(len(refs), len(preds)) + 1
    costs = []
    for reference in refs:
        row = []
        for prediction in preds:
            candidate = edges.get((reference.id, prediction.id))
            reward = 0 if candidate is None else (
                2 * base + 1 if candidate.label == "full" else base
            )
            row.append(-reward)
        costs.append(row + [0] * len(refs))
    row_indices, column_indices = linear_sum_assignment(costs)
    matches = []
    for row_index, column_index in zip(row_indices, column_indices):
        if column_index < len(preds):
            candidate = edges.get((refs[row_index].id, preds[column_index].id))
            if candidate is not None:
                matches.append(candidate)
    return matches


def metrics(n: int, m: int, full: int, partial: int) -> dict[str, object]:
    score = full + partial * 0.5
    return {
        "reference_count": n, "prediction_count": m,
        "full_count": full, "partial_count": partial, "weighted_score": score,
        "strict_precision": full / m if m else None,
        "strict_recall": full / n if n else None,
        "strict_f1": 2 * full / (n + m) if n + m else None,
        "weighted_precision": score / m if m else None,
        "weighted_recall": score / n if n else None,
        "weighted_f1": 2 * score / (n + m) if n + m else None,
        "unmatched_reference_count": n - full - partial,
        "unmatched_prediction_count": m - full - partial,
        "strict_incomplete_reference_count": n - full,
        "strict_not_fully_correct_prediction_count": m - full,
    }


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def call_model(system_prompt: str, user_prompt: str) -> str:
    """接入已有模型函数，传入系统提示词和用户内容，返回模型回复的 JSON 文本。"""
    raise NotImplementedError("请在 call_model 中接入已有模型调用函数，返回 JSON 文本。")


def make_judge(prompt: str, output: Path, responses_dir: Path | None) -> Judge:
    def judge(group: dict[str, object], call_id: str) -> object:
        if responses_dir is not None:
            value = json.loads(
                (responses_dir / f"{call_id}.json").read_text(encoding="utf-8-sig")
            )
        else:
            raw = require_string(
                call_model(prompt, json.dumps(group, ensure_ascii=False)),
                "模型函数返回值",
            )
            (output / "calls" / f"{call_id}.response.txt").write_text(raw, encoding="utf-8")
            value = json.loads(raw)
        write_json(output / "responses" / f"{call_id}.json", value)
        return value

    return judge


def evaluate(
    references: Groups, predictions: Groups, judge: Judge, batch_size: int,
    output: Path,
) -> dict[str, object]:
    if batch_size < 1:
        raise ValueError("batch_size 必须大于零。")
    group_results: list[dict[str, object]] = []
    total_n = total_m = total_full = total_partial = 0
    for group_number, name in enumerate(sorted(set(references) | set(predictions)), 1):
        group_id = f"G{group_number:04d}"
        refs, preds = references.get(name, []), predictions.get(name, [])
        candidates: list[Candidate] = []
        for ref_start in range(0, len(refs), batch_size):
            for pred_start in range(0, len(preds), batch_size):
                group: dict[str, object] = {
                    "group_id": group_id, "group_name": name,
                    "reference_problems": [asdict(problem) for problem
                                           in refs[ref_start:ref_start + batch_size]],
                    "predicted_problems": [asdict(problem) for problem
                                           in preds[pred_start:pred_start + batch_size]],
                }
                call_id = (
                    f"{group_id}_R{ref_start // batch_size + 1:04d}"
                    f"_P{pred_start // batch_size + 1:04d}"
                )
                write_json(output / "calls" / f"{call_id}.input.json", group)
                response = judge(group, call_id)
                candidates.extend(validate_response(response, group))
        matches = match_candidates(refs, preds, candidates)
        full = sum(candidate.label == "full" for candidate in matches)
        partial = sum(candidate.label == "partial" for candidate in matches)
        matched_refs = {candidate.reference_id for candidate in matches}
        matched_preds = {candidate.predicted_id for candidate in matches}
        group_results.append({
            "group_id": group_id, "group_name": name,
            "reference_problems": [asdict(problem) for problem in refs],
            "predicted_problems": [asdict(problem) for problem in preds],
            "candidates": [asdict(candidate) for candidate in candidates],
            "matches": [asdict(candidate) for candidate in matches],
            "unmatched_reference_ids": [p.id for p in refs if p.id not in matched_refs],
            "unmatched_prediction_ids": [p.id for p in preds if p.id not in matched_preds],
            "metrics": metrics(len(refs), len(preds), full, partial),
        })
        total_n += len(refs)
        total_m += len(preds)
        total_full += full
        total_partial += partial
    return {
        "status": "pending_review",
        "note": "模型初评，尚未人工确认；partial、未匹配及多候选竞争项需重点复核。",
        "overall": metrics(total_n, total_m, total_full, total_partial),
        "groups": group_results,
    }


def export_excel(result: dict[str, object], path: Path) -> None:
    def append(worksheet: Worksheet, values: list[object]) -> None:
        for value in values:
            if isinstance(value, str) and (
                len(value) > 32767 or ILLEGAL_CHARACTERS_RE.search(value)
            ):
                raise ValueError("结果含 Excel 无法完整保存的文本，请检查长度或控制字符。")
        worksheet.append(values)
        for cell in worksheet[worksheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = "s"

    workbook = Workbook()
    summary = workbook.active
    if summary is None:
        raise RuntimeError("无法创建结果工作表。")
    summary.title = "指标汇总"
    append(summary, [
        "分组", "参考数", "算法数", "一致数", "部分一致数",
        "严格P", "严格R", "严格F1", "加权P", "加权R", "加权F1", "状态",
    ])
    details = workbook.create_sheet("匹配明细")
    append(details, [
        "分组", "Excel ID", "Excel 序号", "共性特征", "Excel 原文", "Excel 位置",
        "MD ID", "MD 原文", "MD 位置", "标签", "分数", "依据", "差异", "复核建议",
    ])
    unmatched = workbook.create_sheet("误报与漏检")
    append(unmatched, ["分组", "类型", "ID", "序号", "共性特征", "原文", "位置", "复核状态"])
    review = workbook.create_sheet("候选复核")
    append(review, [
        "分组", "Excel ID", "MD ID", "标签", "是否最终匹配",
        "参考证据", "算法证据", "理由", "差异",
    ])
    metric_fields = (
        "reference_count", "prediction_count", "full_count", "partial_count",
        "strict_precision", "strict_recall", "strict_f1",
        "weighted_precision", "weighted_recall", "weighted_f1",
    )
    for raw_group in require_list(result["groups"], "结果分组"):
        group = require_object(raw_group, "结果分组")
        name = group["group_name"]
        counts = require_object(group["metrics"], "分组指标")
        append(summary, [name] + [counts[key] for key in metric_fields] + ["待人工复核"])
        refs = {p["id"]: p for p in (
            require_object(raw, "参考") for raw in require_list(group["reference_problems"], "参考")
        )}
        preds = {p["id"]: p for p in (
            require_object(raw, "算法") for raw in require_list(group["predicted_problems"], "算法")
        )}
        matches = [require_object(raw, "匹配")
                   for raw in require_list(group["matches"], "匹配")]
        candidates = [require_object(raw, "候选")
                      for raw in require_list(group["candidates"], "候选")]
        chosen = {(match["reference_id"], match["predicted_id"]) for match in matches}
        for match in matches:
            rid, pid = match["reference_id"], match["predicted_id"]
            ref, pred = refs[rid], preds[pid]
            competing = sum(c["reference_id"] == rid for c in candidates) > 1 or sum(
                c["predicted_id"] == pid for c in candidates
            ) > 1
            append(details, [
                name, rid, ref["sequence"], ref["common_features"],
                ref["source_text"], ref["source_locator"],
                pid, pred["source_text"], pred["source_locator"],
                match["label"], 1 if match["label"] == "full" else 0.5, match["reason"],
                "；".join(string_list(match["differences"], "差异")),
                "重点复核" if match["label"] == "partial" or competing else "抽样复核",
            ])
        for key, problems, kind in (
            ("unmatched_reference_ids", refs, "漏检"),
            ("unmatched_prediction_ids", preds, "误报"),
        ):
            for problem_id in string_list(group[key], key):
                problem = problems[problem_id]
                append(unmatched, [
                    name, kind, problem_id, problem["sequence"], problem["common_features"],
                    problem["source_text"], problem["source_locator"], "待人工复核",
                ])
        for candidate in candidates:
            append(review, [
                name, candidate["reference_id"], candidate["predicted_id"], candidate["label"],
                "是" if (candidate["reference_id"], candidate["predicted_id"]) in chosen else "否",
                candidate["reference_evidence"], candidate["predicted_evidence"],
                candidate["reason"], "；".join(string_list(candidate["differences"], "差异")),
            ])
    overall = require_object(result["overall"], "整体指标")
    append(summary, ["整体"] + [overall[key] for key in metric_fields] + ["待人工复核"])
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        if worksheet.title == "指标汇总":
            for row in worksheet.iter_rows(min_row=2, min_col=6, max_col=11):
                for cell in row:
                    cell.number_format = "0.00%"
    workbook.save(path)
    workbook.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--excel", type=Path, required=True)
    parser.add_argument("--markdown", type=Path, required=True)
    parser.add_argument("--sheet", help="工作表名称，默认第一个工作表")
    parser.add_argument("--output", type=Path, required=True, help="必须为尚不存在的新目录")
    parser.add_argument("--group-map", type=Path, help="MD 分组名到 Excel 分组名的 JSON 映射")
    parser.add_argument("--prompt", type=Path, default=Path(__file__).with_name("共性问题评估Prompt.txt"))
    parser.add_argument("--batch-size", type=int, default=20, help="每次调用每侧的问题数上限")
    parser.add_argument("--responses-dir", type=Path, help="离线回放/人工修改后的响应目录，不调用模型")
    args = parser.parse_args(argv)
    output_created = False
    try:
        if args.batch_size < 1:
            raise ValueError("batch-size 必须大于零。")
        references = read_excel(args.excel, args.sheet)
        predictions = remap_groups(
            read_markdown(args.markdown), references, args.group_map,
        )
        prompt = args.prompt.read_text(encoding="utf-8-sig")
        if not prompt.strip():
            raise ValueError("Prompt 文件不能为空。")
        args.output.mkdir(parents=True, exist_ok=False)
        output_created = True
        (args.output / "calls").mkdir()
        (args.output / "responses").mkdir()
        write_json(args.output / "inputs.json", {
            "references": {name: [asdict(p) for p in problems]
                           for name, problems in references.items()},
            "predictions": {name: [asdict(p) for p in problems]
                            for name, problems in predictions.items()},
        })
        write_json(args.output / "config.json", {
            "batch_size": args.batch_size,
            "mode": "replay" if args.responses_dir is not None else "function",
            "responses_dir": str(args.responses_dir) if args.responses_dir else None,
        })
        (args.output / "prompt.txt").write_text(prompt, encoding="utf-8")
        judge = make_judge(prompt, args.output, args.responses_dir)
        result = evaluate(references, predictions, judge, args.batch_size, args.output)
        export_excel(result, args.output / "评估结果.xlsx")
        write_json(args.output / "results.json", result)
        print(json.dumps(result["overall"], ensure_ascii=False, indent=2))
        print(f"模型初评已保存：{args.output.resolve()}；尚待人工复核。")
        return 0
    except (OSError, ValueError, RuntimeError, BadZipFile, InvalidFileException) as error:
        message = f"{type(error).__name__}: {error}"
        print(f"评估失败：{message}", file=sys.stderr)
        if output_created:
            write_json(args.output / "failure.json", {"status": "failed", "error": message})
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
