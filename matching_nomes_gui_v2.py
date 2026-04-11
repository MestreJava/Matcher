from __future__ import annotations

import os
import re
import subprocess
import sys
import threading
import traceback
import unicodedata
from collections import defaultdict, deque
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from rapidfuzz import fuzz
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


ProgressCallback = Callable[[str, float | None], None]


# =========================
# MATCHING CORE
# =========================


@dataclass
class AnalysisResult:
    config: dict[str, Any]
    results_df: pd.DataFrame
    candidates_df: pd.DataFrame
    catalog_df: pd.DataFrame
    quota_df: pd.DataFrame
    summary_df: pd.DataFrame
    review_df: pd.DataFrame
    preview_df: pd.DataFrame


class FlowEdge:
    def __init__(self, to_node: int, rev_index: int, capacity: int, cost: int) -> None:
        self.to_node = to_node
        self.rev_index = rev_index
        self.capacity = capacity
        self.cost = cost


def emit_progress(callback: ProgressCallback | None, message: str, percent: float | None = None) -> None:
    if callback:
        callback(message, percent)


def excel_col_to_index(col: str) -> int:
    col = str(col).strip().upper()
    if not col:
        raise ValueError("Column letter cannot be empty.")
    value = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid Excel column: {col}")
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value - 1


def normalize_name(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^A-Z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def first_token(text: str) -> str:
    parts = text.split()
    return parts[0] if parts else ""


def last_token(text: str) -> str:
    parts = text.split()
    return parts[-1] if parts else ""


def token_set(text: str) -> set[str]:
    return set(text.split()) if text else set()


def safe_float(value: Any) -> float | None:
    if pd.isna(value) or value == "":
        return None
    try:
        return float(value)
    except Exception:
        return None


def add_flag(flags: list[str], flag: str, enabled: bool) -> None:
    if enabled and flag not in flags:
        flags.append(flag)


def flags_to_text(flags: list[str]) -> str:
    return "; ".join(sorted(set(flag for flag in flags if flag)))


def score_candidate(full_name: str, external_name: str, max_external_chars: int) -> dict[str, Any]:
    full_tokens = token_set(full_name)
    ext_tokens = token_set(external_name)

    same_first = first_token(full_name) == first_token(external_name) and first_token(full_name) != ""
    same_last = last_token(full_name) == last_token(external_name) and last_token(full_name) != ""
    ext_subset_in_full = bool(ext_tokens) and ext_tokens.issubset(full_tokens)
    full_subset_in_ext = bool(full_tokens) and full_tokens.issubset(ext_tokens)
    starts_like = full_name.startswith(external_name) or external_name.startswith(full_name)

    score_token_set = float(fuzz.token_set_ratio(full_name, external_name))
    score_partial = float(fuzz.partial_ratio(full_name, external_name))
    score_sort = float(fuzz.token_sort_ratio(full_name, external_name))
    score_prefix = float(fuzz.ratio(full_name[:max_external_chars], external_name[:max_external_chars]))

    score = (
        0.35 * score_token_set
        + 0.25 * score_partial
        + 0.20 * score_sort
        + 0.20 * score_prefix
    )
    if same_first:
        score += 6
    if same_last:
        score += 4
    if ext_subset_in_full or full_subset_in_ext:
        score += 8
    if starts_like:
        score += 4

    score = min(score, 100.0)
    structure_ok = same_first and (same_last or ext_subset_in_full or score_token_set >= 88)

    return {
        "score": round(score, 2),
        "same_first": same_first,
        "same_last": same_last,
        "ext_subset_in_full": ext_subset_in_full,
        "full_subset_in_ext": full_subset_in_ext,
        "starts_like": starts_like,
        "score_token_set": round(score_token_set, 2),
        "score_partial": round(score_partial, 2),
        "score_sort": round(score_sort, 2),
        "score_prefix": round(score_prefix, 2),
        "structure_ok": structure_ok,
    }


def build_summary(df: pd.DataFrame, status_column: str = "final_status") -> pd.DataFrame:
    total = max(len(df), 1)
    summary = (
        df[status_column]
        .fillna("SEM_STATUS")
        .value_counts(dropna=False)
        .rename_axis("status")
        .reset_index(name="quantidade")
    )
    summary["percentual"] = (summary["quantidade"] / total * 100).round(2)
    return summary


def _autosize_columns(ws, max_width: int = 56) -> None:
    for col in ws.columns:
        if not col:
            continue
        letter = col[0].column_letter
        best = 0
        for cell in col[:3000]:
            try:
                value = "" if cell.value is None else str(cell.value)
            except Exception:
                value = ""
            best = max(best, len(value))
        ws.column_dimensions[letter].width = min(max(best + 2, 10), max_width)


def _find_header_index(ws, header_name: str) -> int | None:
    target = header_name.strip()
    for idx, cell in enumerate(ws[1], start=1):
        if str(cell.value).strip() == target:
            return idx
    return None


def format_output_workbook(output_file: Path) -> None:
    wb = load_workbook(output_file)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    fill_accepted = PatternFill("solid", fgColor="D9EAD3")
    fill_review = PatternFill("solid", fgColor="FFF2CC")
    fill_no_match = PatternFill("solid", fgColor="F4CCCC")
    fill_summary = PatternFill("solid", fgColor="D9EAF7")
    fill_conflict = PatternFill("solid", fgColor="FCE5CD")

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        _autosize_columns(ws)

        if ws.title == "resumo":
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.fill = fill_summary
            continue

        status_col = _find_header_index(ws, "final_status")
        conflict_col = _find_header_index(ws, "final_conflict_flags")

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            fill = None
            if status_col is not None:
                status = str(row[status_col - 1].value or "").strip().upper()
                if status == "ACEITO":
                    fill = fill_accepted
                elif status == "REVISAR":
                    fill = fill_review
                elif status == "SEM_MATCH":
                    fill = fill_no_match

            if fill:
                for cell in row:
                    cell.fill = fill

            if conflict_col is not None:
                conflict_value = str(row[conflict_col - 1].value or "").strip()
                if conflict_value:
                    row[conflict_col - 1].fill = fill_conflict

    wb.save(output_file)


def open_file_with_default_app(path: Path) -> None:
    path = path.resolve()
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def collect_workbook_preview(config: dict[str, Any]) -> str:
    input_file = Path(config["input_file"])
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    xls = pd.ExcelFile(input_file)
    lines = [f"Workbook: {input_file}", f"Sheets: {', '.join(xls.sheet_names)}", ""]

    for label, sheet_key, header_key, col_key in (
        ("Table 1", "sheet_t1", "header_row_t1", "name_col_t1"),
        ("Table 2", "sheet_t2", "header_row_t2", "name_col_t2"),
    ):
        sheet_name = config[sheet_key]
        header_row = int(config[header_key])
        name_col = config[col_key]
        lines.append(f"{label}:")
        if sheet_name not in xls.sheet_names:
            lines.append(f"  - Missing sheet: {sheet_name}")
            lines.append("")
            continue

        preview_df = pd.read_excel(
            input_file,
            sheet_name=sheet_name,
            header=header_row - 1,
            dtype=str,
            nrows=5,
        )
        headers = [str(col) for col in preview_df.columns]
        lines.append(f"  - Header row: {header_row}")
        lines.append(f"  - Columns: {', '.join(headers[:15])}")
        selected_index = excel_col_to_index(name_col)
        if selected_index >= len(headers):
            lines.append(f"  - Name column {name_col} is out of range")
        else:
            lines.append(f"  - Name column {name_col} -> {headers[selected_index]}")
            sample_values = preview_df.iloc[:, selected_index].fillna("").astype(str).head(5).tolist()
            lines.append(f"  - Sample values: {sample_values}")
        lines.append("")
    return "\n".join(lines).strip()


def validate_config(config: dict[str, Any], validate_workbook: bool = True) -> dict[str, Any]:
    normalized = dict(config)
    normalized["input_file"] = str(Path(normalized["input_file"]).expanduser())
    normalized["output_file"] = str(Path(normalized["output_file"]).expanduser())

    required_text = [
        "input_file",
        "output_file",
        "sheet_t1",
        "sheet_t2",
        "name_col_t1",
        "name_col_t2",
    ]
    for key in required_text:
        if not str(normalized.get(key, "")).strip():
            raise ValueError(f"Field '{key}' cannot be empty.")

    normalized["header_row_t1"] = int(normalized["header_row_t1"])
    normalized["header_row_t2"] = int(normalized["header_row_t2"])
    normalized["max_external_chars"] = int(normalized["max_external_chars"])
    normalized["top_candidates_to_keep"] = int(normalized["top_candidates_to_keep"])
    normalized["max_matches_per_t2_name"] = int(normalized["max_matches_per_t2_name"])
    normalized["accept_score"] = float(normalized["accept_score"])
    normalized["review_score"] = float(normalized["review_score"])
    normalized["min_gap_for_accept"] = float(normalized["min_gap_for_accept"])
    normalized["allow_reuse_t2_matches"] = bool(normalized["allow_reuse_t2_matches"])
    normalized["auto_open_output"] = bool(normalized["auto_open_output"])

    if normalized["header_row_t1"] <= 0 or normalized["header_row_t2"] <= 0:
        raise ValueError("Header rows must be greater than zero.")
    if normalized["max_external_chars"] <= 0:
        raise ValueError("Prefix length must be greater than zero.")
    if normalized["top_candidates_to_keep"] <= 0:
        raise ValueError("Top candidates to keep must be greater than zero.")
    if normalized["max_matches_per_t2_name"] <= 0:
        raise ValueError("Quota override must be greater than zero.")
    if normalized["accept_score"] < normalized["review_score"]:
        raise ValueError("Accept score must be greater than or equal to review score.")
    if normalized["accept_score"] > 100 or normalized["review_score"] > 100:
        raise ValueError("Scores must be between 0 and 100.")

    excel_col_to_index(normalized["name_col_t1"])
    excel_col_to_index(normalized["name_col_t2"])

    input_file = Path(normalized["input_file"])
    if validate_workbook:
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")
        xls = pd.ExcelFile(input_file)
        missing = [sheet for sheet in (normalized["sheet_t1"], normalized["sheet_t2"]) if sheet not in xls.sheet_names]
        if missing:
            raise ValueError(f"Missing sheet(s): {', '.join(missing)}")

    return normalized


def prepare_input_frames(config: dict[str, Any], progress_callback: ProgressCallback | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    input_file = Path(config["input_file"])
    emit_progress(progress_callback, "Reading worksheets...", 5)
    df1 = pd.read_excel(
        input_file,
        sheet_name=config["sheet_t1"],
        header=config["header_row_t1"] - 1,
        dtype=str,
    )
    df2 = pd.read_excel(
        input_file,
        sheet_name=config["sheet_t2"],
        header=config["header_row_t2"] - 1,
        dtype=str,
    )

    idx_t1 = excel_col_to_index(config["name_col_t1"])
    idx_t2 = excel_col_to_index(config["name_col_t2"])
    if idx_t1 >= len(df1.columns):
        raise IndexError(f"Column {config['name_col_t1']} does not exist in {config['sheet_t1']}.")
    if idx_t2 >= len(df2.columns):
        raise IndexError(f"Column {config['name_col_t2']} does not exist in {config['sheet_t2']}.")

    emit_progress(progress_callback, "Normalizing names and metadata...", 12)
    df1 = df1.copy()
    df2 = df2.copy()

    df1["source_row_id"] = df1.index + 1
    df2["target_row_id"] = df2.index + 1
    df1["excel_row_t1"] = config["header_row_t1"] + df1.index + 1
    df2["excel_row_t2"] = config["header_row_t2"] + df2.index + 1

    df1["nome_t1_original"] = df1.iloc[:, idx_t1].fillna("").astype(str)
    df2["nome_t2_original"] = df2.iloc[:, idx_t2].fillna("").astype(str)
    df1["nome_t1_norm"] = df1["nome_t1_original"].apply(normalize_name)
    df2["nome_t2_norm"] = df2["nome_t2_original"].apply(normalize_name)
    df1["first_token"] = df1["nome_t1_norm"].apply(first_token)
    df1["last_token"] = df1["nome_t1_norm"].apply(last_token)
    df2["first_token"] = df2["nome_t2_norm"].apply(first_token)
    df2["last_token"] = df2["nome_t2_norm"].apply(last_token)
    df1["key_prefix"] = df1["nome_t1_norm"].str[: config["max_external_chars"]]
    df2["key_prefix"] = df2["nome_t2_norm"].str[: config["max_external_chars"]]

    return df1, df2


def build_target_catalog(df2: pd.DataFrame, config: dict[str, Any]) -> tuple[pd.DataFrame, dict[str, list[dict[str, Any]]]]:
    df2_valid = df2[df2["nome_t2_norm"] != ""].copy()
    if df2_valid.empty:
        return pd.DataFrame(), {"by_prefix": defaultdict(list), "by_first": defaultdict(list), "by_last": defaultdict(list), "all": []}

    grouped = (
        df2_valid.sort_values("excel_row_t2")
        .groupby("nome_t2_norm", as_index=False)
        .agg(
            nome_t2_original=("nome_t2_original", "first"),
            excel_row_t2=("excel_row_t2", "min"),
            first_token=("first_token", "first"),
            last_token=("last_token", "first"),
            key_prefix=("key_prefix", "first"),
            quota_original=("nome_t2_norm", "size"),
        )
    )
    if config["allow_reuse_t2_matches"]:
        grouped["quota_limit"] = grouped["quota_original"].apply(lambda count: max(int(count), config["max_matches_per_t2_name"]))
    else:
        grouped["quota_limit"] = grouped["quota_original"]

    grouped["quota_limit"] = grouped["quota_limit"].astype(int)
    grouped["quota_original"] = grouped["quota_original"].astype(int)

    indexes: dict[str, Any] = {
        "by_prefix": defaultdict(list),
        "by_first": defaultdict(list),
        "by_last": defaultdict(list),
        "all": [],
    }

    records = grouped.to_dict("records")
    indexes["all"] = records
    for record in records:
        if record["key_prefix"]:
            indexes["by_prefix"][record["key_prefix"]].append(record)
        if record["first_token"]:
            indexes["by_first"][record["first_token"]].append(record)
        if record["last_token"]:
            indexes["by_last"][record["last_token"]].append(record)

    return grouped, indexes


def choose_candidate_pool(
    name_norm: str,
    key_prefix: str,
    ft: str,
    lt: str,
    target_indexes: dict[str, Any],
) -> list[dict[str, Any]]:
    seen: dict[str, dict[str, Any]] = {}

    def absorb(records: list[dict[str, Any]]) -> None:
        for record in records:
            seen.setdefault(record["nome_t2_norm"], record)

    if key_prefix:
        absorb(target_indexes["by_prefix"].get(key_prefix, []))
    if ft:
        absorb(target_indexes["by_first"].get(ft, []))
    if lt:
        absorb(target_indexes["by_last"].get(lt, []))

    if not seen:
        absorb(target_indexes["all"])
    return list(seen.values())


def candidate_utility(candidate: dict[str, Any]) -> int:
    utility = int(round(candidate["score"] * 100))
    if candidate["exact_norm"]:
        utility += 5000
    if candidate["exact_prefix"]:
        utility += 3000
    if candidate["same_last"]:
        utility += 300
    if candidate["ext_subset_in_full"] or candidate["full_subset_in_ext"]:
        utility += 250
    utility += max(0, 100 - int(candidate["rank"])) * 5
    return utility


def add_graph_edge(graph: list[list[FlowEdge]], source: int, target: int, capacity: int, cost: int) -> FlowEdge:
    forward = FlowEdge(target, len(graph[target]), capacity, cost)
    backward = FlowEdge(source, len(graph[source]), 0, -cost)
    graph[source].append(forward)
    graph[target].append(backward)
    return forward


def shortest_path_spfa(graph: list[list[FlowEdge]], source: int, sink: int) -> tuple[list[float], list[int], list[int]]:
    distance = [float("inf")] * len(graph)
    in_queue = [False] * len(graph)
    prev_node = [-1] * len(graph)
    prev_edge = [-1] * len(graph)

    distance[source] = 0
    queue: deque[int] = deque([source])
    in_queue[source] = True

    while queue:
        node = queue.popleft()
        in_queue[node] = False
        for edge_index, edge in enumerate(graph[node]):
            if edge.capacity <= 0:
                continue
            new_distance = distance[node] + edge.cost
            if new_distance < distance[edge.to_node]:
                distance[edge.to_node] = new_distance
                prev_node[edge.to_node] = node
                prev_edge[edge.to_node] = edge_index
                if not in_queue[edge.to_node]:
                    queue.append(edge.to_node)
                    in_queue[edge.to_node] = True
    return distance, prev_node, prev_edge


def solve_global_assignment(candidates_df: pd.DataFrame, quota_limits: dict[str, int]) -> pd.DataFrame:
    eligible = candidates_df[candidates_df["eligible_for_global"]].copy()
    if eligible.empty:
        return pd.DataFrame(columns=["source_row_id", "nome_t2_norm", "assignment_utility"])

    row_ids = sorted(eligible["source_row_id"].unique().tolist())
    targets = sorted(eligible["nome_t2_norm"].unique().tolist())
    row_node = {row_id: index + 1 for index, row_id in enumerate(row_ids)}
    target_node = {name: len(row_ids) + index + 1 for index, name in enumerate(targets)}
    sink = len(row_ids) + len(targets) + 1
    graph: list[list[FlowEdge]] = [[] for _ in range(sink + 1)]

    for row_id in row_ids:
        add_graph_edge(graph, 0, row_node[row_id], 1, 0)
    for target_name in targets:
        add_graph_edge(graph, target_node[target_name], sink, int(quota_limits.get(target_name, 0)), 0)

    assignment_edges: dict[tuple[int, str], FlowEdge] = {}
    for candidate in eligible.to_dict("records"):
        edge = add_graph_edge(
            graph,
            row_node[int(candidate["source_row_id"])],
            target_node[str(candidate["nome_t2_norm"])],
            1,
            -int(candidate["utility"]),
        )
        assignment_edges[(int(candidate["source_row_id"]), str(candidate["nome_t2_norm"]))] = edge

    while True:
        distance, prev_node, prev_edge = shortest_path_spfa(graph, 0, sink)
        if distance[sink] == float("inf") or distance[sink] >= 0:
            break

        flow = 1
        node = sink
        while node != 0:
            edge = graph[prev_node[node]][prev_edge[node]]
            flow = min(flow, edge.capacity)
            node = prev_node[node]

        node = sink
        while node != 0:
            edge = graph[prev_node[node]][prev_edge[node]]
            edge.capacity -= flow
            reverse = graph[edge.to_node][edge.rev_index]
            reverse.capacity += flow
            node = prev_node[node]

    assigned_rows: list[dict[str, Any]] = []
    for (row_id, target_name), edge in assignment_edges.items():
        if edge.capacity == 0:
            assigned_rows.append(
                {
                    "source_row_id": row_id,
                    "nome_t2_norm": target_name,
                    "assignment_utility": -edge.cost,
                }
            )
    return pd.DataFrame(assigned_rows)


def initialize_result_columns(results_df: pd.DataFrame, top_candidates_to_keep: int) -> None:
    for rank in range(1, top_candidates_to_keep + 1):
        results_df[f"cand_{rank}_nome"] = ""
        results_df[f"cand_{rank}_score"] = pd.NA

    defaults = {
        "analysis_status": "",
        "analysis_method": "",
        "analysis_match_t2_original": "",
        "analysis_match_t2_norm": "",
        "analysis_line_match_t2": pd.NA,
        "analysis_score": pd.NA,
        "analysis_score_gap": pd.NA,
        "analysis_conflict_flags": "",
        "analysis_review_reason": "",
        "manual_status": "",
        "manual_match_t2_original": "",
        "manual_match_t2_norm": "",
        "manual_line_match_t2": pd.NA,
        "manual_score": pd.NA,
        "manual_note": "",
        "manual_sequence": pd.NA,
        "final_status": "",
        "final_method": "",
        "final_match_t2_original": "",
        "final_match_t2_norm": "",
        "final_line_match_t2": pd.NA,
        "final_score": pd.NA,
        "final_score_gap": pd.NA,
        "final_conflict_flags": "",
        "final_review_required": False,
        "final_quota_limit": pd.NA,
        "final_quota_order": pd.NA,
        "final_within_quota": pd.NA,
    }
    for column_name, default_value in defaults.items():
        results_df[column_name] = default_value


def analyze_matching(config: dict[str, Any], progress_callback: ProgressCallback | None = None) -> AnalysisResult:
    config = validate_config(config, validate_workbook=True)
    emit_progress(progress_callback, "Preparing analysis session...", 0)
    df1, df2 = prepare_input_frames(config, progress_callback)
    catalog_df, target_indexes = build_target_catalog(df2, config)

    results_df = df1.copy()
    initialize_result_columns(results_df, config["top_candidates_to_keep"])

    if catalog_df.empty:
        results_df["analysis_status"] = "SEM_MATCH"
        results_df["analysis_method"] = "SEM_TABELA2"
        results_df["analysis_review_reason"] = "Table 2 does not contain normalized names."
        recompute_final_state(results_df, pd.DataFrame(columns=catalog_df.columns))
        summary_df = build_summary(results_df)
        return AnalysisResult(
            config=config,
            results_df=results_df,
            candidates_df=pd.DataFrame(),
            catalog_df=pd.DataFrame(),
            quota_df=pd.DataFrame(),
            summary_df=summary_df,
            review_df=results_df[results_df["final_status"] == "REVISAR"].copy(),
            preview_df=results_df.head(30).copy(),
        )

    candidate_rows: list[dict[str, Any]] = []
    emit_progress(progress_callback, "Scoring candidate pools...", 18)
    rows = list(results_df.index)
    internal_keep = max(config["top_candidates_to_keep"], 8)

    for position, row_index in enumerate(rows, start=1):
        if position % 50 == 0 or position == len(rows):
            percent = 18 + (position / max(len(rows), 1)) * 40
            emit_progress(progress_callback, f"Scoring candidates {position}/{len(rows)}...", percent)

        row = results_df.loc[row_index]
        source_row_id = int(row["source_row_id"])
        name_norm = row["nome_t1_norm"]
        if not name_norm:
            results_df.at[row_index, "analysis_status"] = "SEM_MATCH"
            results_df.at[row_index, "analysis_method"] = "SEM_DADO"
            results_df.at[row_index, "analysis_review_reason"] = "Blank input name."
            continue

        pool = choose_candidate_pool(
            name_norm,
            str(row["key_prefix"]),
            str(row["first_token"]),
            str(row["last_token"]),
            target_indexes,
        )
        scored: list[dict[str, Any]] = []
        for record in pool:
            metrics = score_candidate(name_norm, str(record["nome_t2_norm"]), config["max_external_chars"])
            exact_norm = name_norm == record["nome_t2_norm"]
            exact_prefix = bool(row["key_prefix"]) and str(row["key_prefix"]) == str(record["key_prefix"])
            scored.append(
                {
                    "source_row_id": source_row_id,
                    "excel_row_t1": row["excel_row_t1"],
                    "nome_t1_original": row["nome_t1_original"],
                    "nome_t1_norm": name_norm,
                    "nome_t2_original": record["nome_t2_original"],
                    "nome_t2_norm": record["nome_t2_norm"],
                    "excel_row_t2": record["excel_row_t2"],
                    "quota_original": record["quota_original"],
                    "quota_limit": record["quota_limit"],
                    "exact_norm": exact_norm,
                    "exact_prefix": exact_prefix,
                    **metrics,
                }
            )

        scored.sort(
            key=lambda item: (
                item["score"],
                item["exact_norm"],
                item["exact_prefix"],
                item["same_last"],
            ),
            reverse=True,
        )

        kept = scored[:internal_keep]
        for rank, candidate in enumerate(kept, start=1):
            next_score = kept[rank]["score"] if rank < len(kept) else 0.0
            gap_to_next = round(candidate["score"] - next_score, 2)
            candidate["rank"] = rank
            candidate["gap_to_next"] = gap_to_next
            candidate["review_eligible"] = bool(candidate["score"] >= config["review_score"] and candidate["same_first"])
            candidate["eligible_for_global"] = bool(
                candidate["exact_norm"]
                or candidate["exact_prefix"]
                or (candidate["score"] >= config["accept_score"] and candidate["structure_ok"])
            )
            candidate["confident_if_top"] = bool(
                rank == 1
                and (
                    candidate["exact_norm"]
                    or candidate["exact_prefix"]
                    or (
                        candidate["score"] >= config["accept_score"]
                        and candidate["structure_ok"]
                        and gap_to_next >= config["min_gap_for_accept"]
                    )
                )
            )
            candidate["utility"] = candidate_utility(candidate)
            candidate_rows.append(candidate)

            if rank <= config["top_candidates_to_keep"]:
                results_df.at[row_index, f"cand_{rank}_nome"] = candidate["nome_t2_original"]
                results_df.at[row_index, f"cand_{rank}_score"] = candidate["score"]

    candidates_df = pd.DataFrame(candidate_rows)
    emit_progress(progress_callback, "Running global quota-aware assignment...", 62)

    quota_limits = (
        catalog_df.set_index("nome_t2_norm")["quota_limit"].astype(int).to_dict()
        if not catalog_df.empty
        else {}
    )
    assignments_df = solve_global_assignment(candidates_df, quota_limits)

    best_candidates = candidates_df[candidates_df["rank"] == 1].copy() if not candidates_df.empty else pd.DataFrame()
    assigned_candidates = (
        assignments_df.merge(
            candidates_df,
            on=["source_row_id", "nome_t2_norm"],
            how="left",
        )
        if not assignments_df.empty
        else pd.DataFrame()
    )

    best_map = {int(row["source_row_id"]): row for row in best_candidates.to_dict("records")}
    assigned_map = {int(row["source_row_id"]): row for row in assigned_candidates.to_dict("records")}
    eligible_counts = (
        candidates_df[candidates_df["eligible_for_global"]]
        .groupby("source_row_id")
        .size()
        .to_dict()
        if not candidates_df.empty
        else {}
    )

    emit_progress(progress_callback, "Classifying analysis outcomes...", 78)
    for row_index in results_df.index:
        source_row_id = int(results_df.at[row_index, "source_row_id"])
        best = best_map.get(source_row_id)
        assigned = assigned_map.get(source_row_id)
        flags: list[str] = []

        if not results_df.at[row_index, "analysis_status"]:
            if not best:
                results_df.at[row_index, "analysis_status"] = "SEM_MATCH"
                results_df.at[row_index, "analysis_method"] = "SEM_CANDIDATO"
                results_df.at[row_index, "analysis_review_reason"] = "No candidate generated."
            else:
                add_flag(flags, "LOW_GAP", best["gap_to_next"] < config["min_gap_for_accept"] and not (best["exact_norm"] or best["exact_prefix"]))
                add_flag(flags, "STRUCTURE_WARNING", best["score"] >= config["accept_score"] and not best["structure_ok"])
                add_flag(flags, "QUOTA_CONFLICT", bool(eligible_counts.get(source_row_id)) and assigned is None)
                add_flag(flags, "GLOBAL_REALLOCATED", assigned is not None and int(assigned["rank"]) > 1)

                chosen = assigned if assigned is not None else best
                results_df.at[row_index, "analysis_match_t2_original"] = chosen["nome_t2_original"]
                results_df.at[row_index, "analysis_match_t2_norm"] = chosen["nome_t2_norm"]
                results_df.at[row_index, "analysis_line_match_t2"] = chosen["excel_row_t2"]
                results_df.at[row_index, "analysis_score"] = chosen["score"]
                results_df.at[row_index, "analysis_score_gap"] = chosen["gap_to_next"]

                if assigned is not None and int(assigned["rank"]) == 1 and bool(assigned["confident_if_top"]):
                    results_df.at[row_index, "analysis_status"] = "ACEITO"
                    results_df.at[row_index, "analysis_method"] = (
                        "EXATO_GLOBAL" if assigned["exact_norm"] or assigned["exact_prefix"] else "FUZZY_GLOBAL"
                    )
                    results_df.at[row_index, "analysis_review_reason"] = ""
                elif assigned is not None:
                    results_df.at[row_index, "analysis_status"] = "REVISAR"
                    results_df.at[row_index, "analysis_method"] = "GLOBAL_REVIEW"
                    reasons = []
                    if int(assigned["rank"]) > 1:
                        reasons.append("Global assignment used a fallback candidate.")
                    if best["gap_to_next"] < config["min_gap_for_accept"]:
                        reasons.append("Primary candidate gap is below the auto-accept threshold.")
                    results_df.at[row_index, "analysis_review_reason"] = " ".join(reasons) or "Global fallback should be reviewed."
                elif bool(eligible_counts.get(source_row_id)):
                    results_df.at[row_index, "analysis_status"] = "REVISAR"
                    results_df.at[row_index, "analysis_method"] = "QUOTA_CONFLICT"
                    results_df.at[row_index, "analysis_review_reason"] = "Strong candidate lost quota in the global assignment."
                elif bool(best["review_eligible"]):
                    results_df.at[row_index, "analysis_status"] = "REVISAR"
                    results_df.at[row_index, "analysis_method"] = "FUZZY_REVIEW"
                    results_df.at[row_index, "analysis_review_reason"] = "Candidate needs manual confirmation."
                else:
                    results_df.at[row_index, "analysis_status"] = "SEM_MATCH"
                    results_df.at[row_index, "analysis_method"] = "SEM_MATCH"
                    results_df.at[row_index, "analysis_review_reason"] = "No candidate met the review threshold."

        results_df.at[row_index, "analysis_conflict_flags"] = flags_to_text(flags)

    emit_progress(progress_callback, "Applying final state defaults...", 88)
    recompute_final_state(results_df, catalog_df)
    summary_df = build_summary(results_df)
    review_df = results_df[results_df["final_status"] == "REVISAR"].copy()
    preview_columns = [
        "excel_row_t1",
        "nome_t1_original",
        "analysis_status",
        "final_status",
        "final_match_t2_original",
        "final_score",
        "final_conflict_flags",
    ]
    preview_df = results_df[[col for col in preview_columns if col in results_df.columns]].head(40).copy()
    emit_progress(progress_callback, "Analysis complete.", 100)
    return AnalysisResult(
        config=config,
        results_df=results_df,
        candidates_df=candidates_df,
        catalog_df=catalog_df,
        quota_df=build_quota_summary(results_df, catalog_df),
        summary_df=summary_df,
        review_df=review_df,
        preview_df=preview_df,
    )


def build_quota_summary(results_df: pd.DataFrame, catalog_df: pd.DataFrame) -> pd.DataFrame:
    if catalog_df.empty:
        return pd.DataFrame(columns=["nome_t2_norm", "nome_t2_original", "quota_original", "quota_limit", "accepted_count"])

    accepted_counts = (
        results_df.loc[
            (results_df["final_status"] == "ACEITO") & (results_df["final_match_t2_norm"] != ""),
            "final_match_t2_norm",
        ]
        .value_counts()
        .to_dict()
    )
    quota_df = catalog_df.copy()
    quota_df["accepted_count"] = quota_df["nome_t2_norm"].map(accepted_counts).fillna(0).astype(int)
    quota_df["remaining_quota"] = (quota_df["quota_limit"] - quota_df["accepted_count"]).astype(int)
    quota_df["is_full"] = quota_df["remaining_quota"] <= 0
    return quota_df.sort_values(["is_full", "nome_t2_original"], ascending=[False, True]).reset_index(drop=True)


def recompute_final_state(results_df: pd.DataFrame, catalog_df: pd.DataFrame) -> None:
    quota_map = (
        catalog_df.set_index("nome_t2_norm")["quota_limit"].astype(int).to_dict()
        if not catalog_df.empty
        else {}
    )
    original_map = (
        catalog_df.set_index("nome_t2_norm")["nome_t2_original"].to_dict()
        if not catalog_df.empty
        else {}
    )
    line_map = (
        catalog_df.set_index("nome_t2_norm")["excel_row_t2"].to_dict()
        if not catalog_df.empty
        else {}
    )

    provisional: list[dict[str, Any]] = []
    for row_index in results_df.index:
        analysis_status = str(results_df.at[row_index, "analysis_status"] or "")
        analysis_method = str(results_df.at[row_index, "analysis_method"] or "")
        analysis_match_norm = str(results_df.at[row_index, "analysis_match_t2_norm"] or "")
        analysis_match_original = str(results_df.at[row_index, "analysis_match_t2_original"] or "")
        analysis_score = results_df.at[row_index, "analysis_score"]
        analysis_gap = results_df.at[row_index, "analysis_score_gap"]
        analysis_line = results_df.at[row_index, "analysis_line_match_t2"]
        analysis_flags = str(results_df.at[row_index, "analysis_conflict_flags"] or "")

        manual_status = str(results_df.at[row_index, "manual_status"] or "").upper()
        manual_norm = str(results_df.at[row_index, "manual_match_t2_norm"] or "")
        manual_original = str(results_df.at[row_index, "manual_match_t2_original"] or "")
        manual_line = results_df.at[row_index, "manual_line_match_t2"]
        manual_score = results_df.at[row_index, "manual_score"]
        manual_note = str(results_df.at[row_index, "manual_note"] or "")

        final_status = analysis_status
        final_method = analysis_method
        final_norm = analysis_match_norm
        final_original = analysis_match_original
        final_line = analysis_line
        final_score = analysis_score
        final_gap = analysis_gap
        final_flags = [flag for flag in analysis_flags.split("; ") if flag]

        if manual_status == "ACCEPT" and manual_norm:
            final_status = "ACEITO"
            final_method = "MANUAL_ACCEPT"
            final_norm = manual_norm
            final_original = manual_original or original_map.get(manual_norm, "")
            final_line = manual_line if pd.notna(manual_line) else line_map.get(manual_norm, pd.NA)
            final_score = manual_score if pd.notna(manual_score) else final_score
            add_flag(final_flags, "MANUAL_OVERRIDE", True)
        elif manual_status == "NO_MATCH":
            final_status = "SEM_MATCH"
            final_method = "MANUAL_NO_MATCH"
            final_norm = ""
            final_original = ""
            final_line = pd.NA
            final_score = pd.NA
            final_gap = pd.NA
            add_flag(final_flags, "MANUAL_OVERRIDE", True)
        elif manual_status == "REVIEW":
            final_status = "REVISAR"
            final_method = "MANUAL_REVIEW"
            add_flag(final_flags, "MANUAL_OVERRIDE", True)

        if manual_note:
            add_flag(final_flags, "MANUAL_NOTE", True)

        provisional.append(
            {
                "row_index": row_index,
                "final_status": final_status,
                "final_method": final_method,
                "final_match_t2_norm": final_norm,
                "final_match_t2_original": final_original,
                "final_line_match_t2": final_line,
                "final_score": final_score,
                "final_score_gap": final_gap,
                "final_flags": final_flags,
                "manual_status": manual_status,
                "manual_sequence": results_df.at[row_index, "manual_sequence"],
            }
        )

    provisional_df = pd.DataFrame(provisional)
    accepted_df = provisional_df[
        (provisional_df["final_status"] == "ACEITO") & (provisional_df["final_match_t2_norm"] != "")
    ].copy()

    keep_rows: set[int] = set()
    if not accepted_df.empty:
        accepted_df["manual_priority"] = accepted_df["manual_status"].eq("ACCEPT").astype(int)
        accepted_df["score_priority"] = accepted_df["final_score"].apply(lambda value: safe_float(value) or 0.0)
        accepted_df["sequence_priority"] = accepted_df["manual_sequence"].apply(lambda value: safe_float(value) or 0.0)
        accepted_df = accepted_df.sort_values(
            ["final_match_t2_norm", "manual_priority", "score_priority", "sequence_priority", "row_index"],
            ascending=[True, False, False, True, True],
        )
        accepted_df["quota_order"] = accepted_df.groupby("final_match_t2_norm").cumcount() + 1
        accepted_df["quota_limit"] = accepted_df["final_match_t2_norm"].map(quota_map).fillna(0).astype(int)
        keep_rows = set(accepted_df.loc[accepted_df["quota_order"] <= accepted_df["quota_limit"], "row_index"].tolist())

    for row in provisional:
        row_index = row["row_index"]
        final_status = row["final_status"]
        final_method = row["final_method"]
        final_norm = row["final_match_t2_norm"]
        final_flags = row["final_flags"]
        quota_limit = quota_map.get(final_norm) if final_norm else None
        quota_order = pd.NA
        within_quota: Any = pd.NA

        if final_status == "ACEITO" and final_norm:
            matching_rows = accepted_df[accepted_df["row_index"] == row_index]
            if not matching_rows.empty:
                quota_order = int(matching_rows.iloc[0]["quota_order"])
                quota_limit = int(matching_rows.iloc[0]["quota_limit"])
                within_quota = quota_order <= quota_limit
            else:
                within_quota = False

            if row_index not in keep_rows:
                final_status = "REVISAR"
                final_method = "FINAL_QUOTA_CONFLICT"
                add_flag(final_flags, "FINAL_QUOTA_CONFLICT", True)

        results_df.at[row_index, "final_status"] = final_status
        results_df.at[row_index, "final_method"] = final_method
        results_df.at[row_index, "final_match_t2_norm"] = row["final_match_t2_norm"] if final_status == "ACEITO" else (
            row["final_match_t2_norm"] if final_status == "REVISAR" else ""
        )
        results_df.at[row_index, "final_match_t2_original"] = row["final_match_t2_original"] if final_status != "SEM_MATCH" else ""
        results_df.at[row_index, "final_line_match_t2"] = row["final_line_match_t2"] if final_status != "SEM_MATCH" else pd.NA
        results_df.at[row_index, "final_score"] = row["final_score"] if final_status != "SEM_MATCH" else pd.NA
        results_df.at[row_index, "final_score_gap"] = row["final_score_gap"] if final_status != "SEM_MATCH" else pd.NA
        results_df.at[row_index, "final_conflict_flags"] = flags_to_text(final_flags)
        results_df.at[row_index, "final_review_required"] = final_status == "REVISAR"
        results_df.at[row_index, "final_quota_limit"] = quota_limit if quota_limit is not None else pd.NA
        results_df.at[row_index, "final_quota_order"] = quota_order
        results_df.at[row_index, "final_within_quota"] = within_quota


def export_analysis_result(
    result: AnalysisResult,
    output_file: str | Path | None = None,
    progress_callback: ProgressCallback | None = None,
) -> Path:
    config = dict(result.config)
    if output_file is not None:
        config["output_file"] = str(output_file)
    output_path = Path(config["output_file"])

    emit_progress(progress_callback, "Refreshing final state before export...", 10)
    recompute_final_state(result.results_df, result.catalog_df)
    result.summary_df = build_summary(result.results_df)
    result.review_df = result.results_df[result.results_df["final_status"] == "REVISAR"].copy()
    result.preview_df = result.results_df.head(40).copy()
    result.quota_df = build_quota_summary(result.results_df, result.catalog_df)

    conflicts_df = result.results_df[result.results_df["final_conflict_flags"].fillna("") != ""].copy()
    accepted_df = result.results_df[result.results_df["final_status"] == "ACEITO"].copy()
    review_df = result.results_df[result.results_df["final_status"] == "REVISAR"].copy()
    sem_match_df = result.results_df[result.results_df["final_status"] == "SEM_MATCH"].copy()
    used_targets = set(accepted_df["final_match_t2_norm"].dropna().astype(str))
    full_catalog = build_export_catalog(result)
    unused_targets_df = full_catalog[~full_catalog["nome_t2_norm"].isin(used_targets)].copy() if not full_catalog.empty else pd.DataFrame()

    emit_progress(progress_callback, "Writing Excel workbook...", 55)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result.summary_df.to_excel(writer, sheet_name="resumo", index=False)
        result.results_df.to_excel(writer, sheet_name="resultados_todos", index=False)
        accepted_df.to_excel(writer, sheet_name="aceitos", index=False)
        review_df.to_excel(writer, sheet_name="revisao_pendente", index=False)
        sem_match_df.to_excel(writer, sheet_name="sem_match", index=False)
        conflicts_df.to_excel(writer, sheet_name="conflitos", index=False)
        result.quota_df.to_excel(writer, sheet_name="quotas_t2", index=False)
        result.candidates_df.to_excel(writer, sheet_name="candidatos", index=False)
        unused_targets_df.to_excel(writer, sheet_name="t2_nao_utilizados", index=False)

    emit_progress(progress_callback, "Formatting workbook...", 88)
    format_output_workbook(output_path)
    emit_progress(progress_callback, f"Export complete: {output_path}", 100)
    return output_path


def build_export_catalog(result: AnalysisResult) -> pd.DataFrame:
    if result.catalog_df.empty:
        columns = ["nome_t2_norm", "nome_t2_original", "quota_original", "quota_limit", "excel_row_t2"]
        return pd.DataFrame(columns=columns)
    return result.catalog_df.copy()


def run_matching(config: dict[str, Any], progress_callback: ProgressCallback | None = None) -> Path:
    result = analyze_matching(config, progress_callback=progress_callback)
    return export_analysis_result(result, progress_callback=progress_callback)


# =========================
# GUI
# =========================


class MatcherApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Matcher de Nomes - Analysis / Review / Export")
        self.root.geometry("1320x920")
        self.last_output_path: Path | None = None
        self.analysis_result: AnalysisResult | None = None
        self.catalog_df: pd.DataFrame = pd.DataFrame()
        self.manual_sequence = 0

        self.vars: dict[str, tk.Variable] = {
            "input_file": tk.StringVar(),
            "output_file": tk.StringVar(value="resultado_matching.xlsx"),
            "sheet_t1": tk.StringVar(value="Tabela1"),
            "sheet_t2": tk.StringVar(value="Tabela2"),
            "header_row_t1": tk.StringVar(value="1"),
            "header_row_t2": tk.StringVar(value="1"),
            "name_col_t1": tk.StringVar(value="C"),
            "name_col_t2": tk.StringVar(value="E"),
            "max_external_chars": tk.StringVar(value="30"),
            "accept_score": tk.StringVar(value="92"),
            "review_score": tk.StringVar(value="85"),
            "min_gap_for_accept": tk.StringVar(value="4"),
            "top_candidates_to_keep": tk.StringVar(value="5"),
            "allow_reuse_t2_matches": tk.BooleanVar(value=False),
            "max_matches_per_t2_name": tk.StringVar(value="3"),
            "auto_open_output": tk.BooleanVar(value=True),
        }
        self.status_var = tk.StringVar(value="Ready.")
        self.progress_var = tk.DoubleVar(value=0.0)
        self.manual_note_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=10)
        container.pack(fill="both", expand=True)

        header = ttk.Frame(container)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="Matcher V2 Rewrite", font=("Segoe UI", 15, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Validation -> Analysis -> Manual Review -> Export",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(2, 0))

        status_bar = ttk.Frame(container)
        status_bar.pack(fill="x", pady=(0, 10))
        ttk.Label(status_bar, textvariable=self.status_var, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Progressbar(status_bar, variable=self.progress_var, maximum=100, length=360).pack(side="right")

        self.notebook = ttk.Notebook(container)
        self.notebook.pack(fill="both", expand=True)

        self.tab_config = ttk.Frame(self.notebook, padding=10)
        self.tab_analyze = ttk.Frame(self.notebook, padding=10)
        self.tab_review = ttk.Frame(self.notebook, padding=10)
        self.tab_export = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_config, text="Configuration")
        self.notebook.add(self.tab_analyze, text="Analyze")
        self.notebook.add(self.tab_review, text="Review")
        self.notebook.add(self.tab_export, text="Export")

        self._build_config_tab()
        self._build_analyze_tab()
        self._build_review_tab()
        self._build_export_tab()

        self.log("Application initialized.")

    def _build_config_tab(self) -> None:
        files_frame = ttk.LabelFrame(self.tab_config, text="Files", padding=10)
        files_frame.pack(fill="x", pady=(0, 8))
        self._add_file_row(files_frame, "Input workbook", "input_file", self.pick_input_file, 0)
        self._add_file_row(files_frame, "Output workbook", "output_file", self.pick_output_file, 1)

        config_frame = ttk.LabelFrame(self.tab_config, text="Matching configuration", padding=10)
        config_frame.pack(fill="x", pady=(0, 8))

        fields = [
            ("Sheet T1", "sheet_t1"),
            ("Sheet T2", "sheet_t2"),
            ("Header row T1", "header_row_t1"),
            ("Header row T2", "header_row_t2"),
            ("Name column T1", "name_col_t1"),
            ("Name column T2", "name_col_t2"),
            ("Prefix length", "max_external_chars"),
            ("Accept score", "accept_score"),
            ("Review score", "review_score"),
            ("Min gap", "min_gap_for_accept"),
            ("Preview candidates", "top_candidates_to_keep"),
            ("Quota override", "max_matches_per_t2_name"),
        ]
        for index, (label, key) in enumerate(fields):
            row = index // 2
            col = (index % 2) * 2
            ttk.Label(config_frame, text=label).grid(row=row, column=col, sticky="w", padx=6, pady=5)
            ttk.Entry(config_frame, textvariable=self.vars[key], width=24).grid(row=row, column=col + 1, sticky="ew", padx=6, pady=5)
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(3, weight=1)

        options_frame = ttk.LabelFrame(self.tab_config, text="Options", padding=10)
        options_frame.pack(fill="x", pady=(0, 8))
        ttk.Checkbutton(
            options_frame,
            text="Allow T2 names to be reused up to quota override",
            variable=self.vars["allow_reuse_t2_matches"],
        ).pack(anchor="w")
        ttk.Checkbutton(
            options_frame,
            text="Open exported workbook automatically",
            variable=self.vars["auto_open_output"],
        ).pack(anchor="w", pady=(6, 0))

        button_row = ttk.Frame(self.tab_config)
        button_row.pack(fill="x", pady=(0, 8))
        ttk.Button(button_row, text="Autofill Output Name", command=self.autofill_output_name).pack(side="left")
        ttk.Button(button_row, text="Validate + Preview Workbook", command=self.validate_and_preview).pack(side="left", padx=8)
        ttk.Button(button_row, text="Start Analysis", command=self.start_analysis).pack(side="left", padx=8)

        preview_frame = ttk.LabelFrame(self.tab_config, text="Validation preview", padding=10)
        preview_frame.pack(fill="both", expand=True)
        self.validation_text = tk.Text(preview_frame, wrap="word", height=22)
        self.validation_text.pack(fill="both", expand=True)

    def _build_analyze_tab(self) -> None:
        actions = ttk.Frame(self.tab_analyze)
        actions.pack(fill="x", pady=(0, 8))
        self.analyze_button = ttk.Button(actions, text="Run Analysis", command=self.start_analysis)
        self.analyze_button.pack(side="left")
        ttk.Button(actions, text="Refresh Preview", command=self.refresh_analysis_views).pack(side="left", padx=8)

        content = ttk.Panedwindow(self.tab_analyze, orient="vertical")
        content.pack(fill="both", expand=True)

        summary_frame = ttk.LabelFrame(content, text="Summary", padding=8)
        preview_frame = ttk.LabelFrame(content, text="Result preview", padding=8)
        log_frame = ttk.LabelFrame(content, text="Analysis log", padding=8)
        content.add(summary_frame, weight=1)
        content.add(preview_frame, weight=3)
        content.add(log_frame, weight=2)

        self.summary_tree = ttk.Treeview(summary_frame, columns=("status", "qty", "pct"), show="headings", height=6)
        for column, title, width in (("status", "Status", 180), ("qty", "Count", 100), ("pct", "Percent", 100)):
            self.summary_tree.heading(column, text=title)
            self.summary_tree.column(column, width=width, anchor="center")
        self.summary_tree.pack(fill="both", expand=True)

        preview_columns = ("excel_row", "name", "analysis", "final", "match", "score", "flags")
        self.preview_tree = ttk.Treeview(preview_frame, columns=preview_columns, show="headings", height=14)
        headers = {
            "excel_row": "Row",
            "name": "T1 Name",
            "analysis": "Analysis",
            "final": "Final",
            "match": "Suggested T2",
            "score": "Score",
            "flags": "Flags",
        }
        widths = {"excel_row": 70, "name": 240, "analysis": 100, "final": 100, "match": 240, "score": 80, "flags": 260}
        for column in preview_columns:
            self.preview_tree.heading(column, text=headers[column])
            self.preview_tree.column(column, width=widths[column], anchor="w")
        self.preview_tree.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, wrap="word", height=10)
        self.log_text.pack(fill="both", expand=True)

    def _build_review_tab(self) -> None:
        layout = ttk.Panedwindow(self.tab_review, orient="horizontal")
        layout.pack(fill="both", expand=True)

        queue_frame = ttk.LabelFrame(layout, text="Rows requiring review", padding=8)
        detail_frame = ttk.LabelFrame(layout, text="Selected row details", padding=8)
        layout.add(queue_frame, weight=2)
        layout.add(detail_frame, weight=3)

        queue_columns = ("row", "name", "status", "flags", "suggested")
        self.review_tree = ttk.Treeview(queue_frame, columns=queue_columns, show="headings", height=24)
        for column, title, width in (
            ("row", "Row", 70),
            ("name", "T1 Name", 220),
            ("status", "Final Status", 110),
            ("flags", "Flags", 230),
            ("suggested", "Suggested T2", 220),
        ):
            self.review_tree.heading(column, text=title)
            self.review_tree.column(column, width=width, anchor="w")
        self.review_tree.pack(fill="both", expand=True)
        self.review_tree.bind("<<TreeviewSelect>>", self.on_review_row_selected)

        ttk.Label(detail_frame, text="Row preview").pack(anchor="w")
        self.review_detail_text = tk.Text(detail_frame, height=10, wrap="word")
        self.review_detail_text.pack(fill="x", pady=(0, 8))

        ttk.Label(detail_frame, text="Candidate preview").pack(anchor="w")
        candidate_columns = ("rank", "candidate", "score", "auto", "review", "quota")
        self.candidate_tree = ttk.Treeview(detail_frame, columns=candidate_columns, show="headings", height=14)
        for column, title, width in (
            ("rank", "Rank", 60),
            ("candidate", "T2 Candidate", 260),
            ("score", "Score", 80),
            ("auto", "Auto", 70),
            ("review", "Review", 70),
            ("quota", "Quota", 90),
        ):
            self.candidate_tree.heading(column, text=title)
            self.candidate_tree.column(column, width=width, anchor="w")
        self.candidate_tree.pack(fill="both", expand=True, pady=(0, 8))

        note_frame = ttk.Frame(detail_frame)
        note_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(note_frame, text="Manual note").pack(side="left")
        ttk.Entry(note_frame, textvariable=self.manual_note_var).pack(side="left", fill="x", expand=True, padx=(8, 0))

        buttons = ttk.Frame(detail_frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Accept Selected Candidate", command=self.accept_selected_candidate).pack(side="left")
        ttk.Button(buttons, text="Mark No Match", command=self.mark_selected_no_match).pack(side="left", padx=8)
        ttk.Button(buttons, text="Keep In Review", command=self.keep_selected_in_review).pack(side="left")
        ttk.Button(buttons, text="Reset Manual Decision", command=self.reset_manual_decision).pack(side="left", padx=8)

    def _build_export_tab(self) -> None:
        actions = ttk.Frame(self.tab_export)
        actions.pack(fill="x", pady=(0, 8))
        self.export_button = ttk.Button(actions, text="Export Reviewed Results", command=self.start_export)
        self.export_button.pack(side="left")
        ttk.Button(actions, text="Open Last Export", command=self.open_last_output).pack(side="left", padx=8)

        info_frame = ttk.LabelFrame(self.tab_export, text="Export status", padding=8)
        info_frame.pack(fill="both", expand=True)
        self.export_text = tk.Text(info_frame, wrap="word", height=28)
        self.export_text.pack(fill="both", expand=True)

    def _add_file_row(self, parent: ttk.LabelFrame, label: str, var_key: str, command: Callable[[], None], row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(parent, textvariable=self.vars[var_key], width=96).grid(row=row, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="Browse...", command=command).grid(row=row, column=2, padx=6, pady=5)
        parent.columnconfigure(1, weight=1)

    def log(self, message: str) -> None:
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def set_status(self, message: str, percent: float | None = None) -> None:
        self.status_var.set(message)
        if percent is not None:
            self.progress_var.set(percent)

    def pick_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select input workbook",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            self.vars["input_file"].set(path)
            self.autofill_output_name()
            self.log(f"Input workbook selected: {path}")

    def pick_output_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save output workbook as",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.vars["output_file"].get() or "resultado_matching.xlsx").name,
        )
        if path:
            self.vars["output_file"].set(path)
            self.log(f"Output workbook set to: {path}")

    def autofill_output_name(self) -> None:
        input_path = self.vars["input_file"].get().strip()
        if not input_path:
            self.vars["output_file"].set("resultado_matching.xlsx")
            return
        path = Path(input_path)
        self.vars["output_file"].set(str(path.with_name(f"{path.stem}_resultado_matching.xlsx")))

    def collect_config_from_vars(self) -> dict[str, Any]:
        return {
            key: variable.get().strip() if isinstance(variable, tk.StringVar) else variable.get()
            for key, variable in self.vars.items()
        }

    def validate_and_preview(self) -> None:
        try:
            config = validate_config(self.collect_config_from_vars(), validate_workbook=True)
            preview = collect_workbook_preview(config)
            self.validation_text.delete("1.0", "end")
            self.validation_text.insert("1.0", preview)
            self.set_status("Workbook validation passed.", 0)
            self.log("Workbook validation preview refreshed.")
            self.notebook.select(self.tab_config)
        except Exception as exc:
            messagebox.showerror("Validation error", str(exc))

    def start_analysis(self) -> None:
        try:
            config = validate_config(self.collect_config_from_vars(), validate_workbook=True)
        except Exception as exc:
            messagebox.showerror("Validation error", str(exc))
            return

        self.set_busy(True)
        self.set_status("Starting analysis...", 0)
        self.log("Starting analysis.")

        def worker() -> None:
            try:
                result = analyze_matching(config, progress_callback=self.safe_progress)
                self.root.after(0, lambda: self.on_analysis_success(result))
            except Exception as exc:
                tb = traceback.format_exc()
                self.root.after(0, lambda: self.on_background_error(str(exc), tb))

        threading.Thread(target=worker, daemon=True).start()

    def safe_progress(self, message: str, percent: float | None = None) -> None:
        self.root.after(0, lambda: self._update_progress(message, percent))

    def _update_progress(self, message: str, percent: float | None = None) -> None:
        self.set_status(message, percent)
        self.log(message)

    def set_busy(self, is_busy: bool) -> None:
        state = "disabled" if is_busy else "normal"
        self.analyze_button.config(state=state)
        self.export_button.config(state=state)

    def on_analysis_success(self, result: AnalysisResult) -> None:
        self.analysis_result = result
        self.catalog_df = result.catalog_df.copy()
        self.last_output_path = None
        self.set_busy(False)
        self.set_status("Analysis completed.", 100)
        self.refresh_analysis_views()
        self.refresh_review_views()
        self.refresh_export_view()
        self.notebook.select(self.tab_analyze)
        self.log("Analysis completed successfully.")

    def on_background_error(self, error_message: str, tb: str) -> None:
        self.set_busy(False)
        self.set_status("Operation failed.", 0)
        self.log("ERROR:")
        self.log(error_message)
        self.log(tb)
        messagebox.showerror("Error", error_message)

    def refresh_analysis_views(self) -> None:
        if self.analysis_result is None:
            return

        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        summary_df = build_summary(self.analysis_result.results_df)
        self.analysis_result.summary_df = summary_df
        for row in summary_df.itertuples(index=False):
            self.summary_tree.insert("", "end", values=(row.status, row.quantidade, f"{row.percentual:.2f}%"))

        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        preview_df = self.analysis_result.results_df.head(50)
        for row in preview_df.itertuples(index=False):
            self.preview_tree.insert(
                "",
                "end",
                values=(
                    getattr(row, "excel_row_t1", ""),
                    getattr(row, "nome_t1_original", ""),
                    getattr(row, "analysis_status", ""),
                    getattr(row, "final_status", ""),
                    getattr(row, "final_match_t2_original", ""),
                    getattr(row, "final_score", ""),
                    getattr(row, "final_conflict_flags", ""),
                ),
            )

    def refresh_review_views(self) -> None:
        if self.analysis_result is None:
            return

        recompute_final_state(self.analysis_result.results_df, self.catalog_df)
        self.analysis_result.review_df = self.analysis_result.results_df[self.analysis_result.results_df["final_status"] == "REVISAR"].copy()

        for item in self.review_tree.get_children():
            self.review_tree.delete(item)
        for row in self.analysis_result.review_df.itertuples(index=False):
            self.review_tree.insert(
                "",
                "end",
                iid=str(getattr(row, "source_row_id")),
                values=(
                    getattr(row, "excel_row_t1", ""),
                    getattr(row, "nome_t1_original", ""),
                    getattr(row, "final_status", ""),
                    getattr(row, "final_conflict_flags", ""),
                    getattr(row, "final_match_t2_original", ""),
                ),
            )

        self.review_detail_text.delete("1.0", "end")
        for item in self.candidate_tree.get_children():
            self.candidate_tree.delete(item)

    def on_review_row_selected(self, _event=None) -> None:
        if self.analysis_result is None:
            return
        selection = self.review_tree.selection()
        if not selection:
            return

        source_row_id = int(selection[0])
        row_df = self.analysis_result.results_df[self.analysis_result.results_df["source_row_id"] == source_row_id]
        if row_df.empty:
            return
        row = row_df.iloc[0]
        self.manual_note_var.set(str(row.get("manual_note", "") or ""))

        details = [
            f"Excel row: {row['excel_row_t1']}",
            f"T1 original: {row['nome_t1_original']}",
            f"T1 normalized: {row['nome_t1_norm']}",
            f"Analysis status: {row['analysis_status']}",
            f"Analysis method: {row['analysis_method']}",
            f"Suggested T2: {row['analysis_match_t2_original']}",
            f"Suggested score: {row['analysis_score']}",
            f"Review reason: {row['analysis_review_reason']}",
            f"Conflict flags: {row['final_conflict_flags']}",
        ]
        self.review_detail_text.delete("1.0", "end")
        self.review_detail_text.insert("1.0", "\n".join(details))

        for item in self.candidate_tree.get_children():
            self.candidate_tree.delete(item)

        candidates = self.analysis_result.candidates_df[
            self.analysis_result.candidates_df["source_row_id"] == source_row_id
        ].sort_values(["rank", "score"], ascending=[True, False])
        for candidate in candidates.itertuples(index=False):
            quota_text = f"{getattr(candidate, 'quota_limit', '')}"
            self.candidate_tree.insert(
                "",
                "end",
                iid=f"{source_row_id}:{getattr(candidate, 'rank')}",
                values=(
                    getattr(candidate, "rank"),
                    getattr(candidate, "nome_t2_original"),
                    getattr(candidate, "score"),
                    "Y" if getattr(candidate, "eligible_for_global") else "",
                    "Y" if getattr(candidate, "review_eligible") else "",
                    quota_text,
                ),
            )

    def _selected_source_row_id(self) -> int | None:
        selection = self.review_tree.selection()
        if not selection:
            messagebox.showwarning("Review", "Select a row from the review queue first.")
            return None
        return int(selection[0])

    def accept_selected_candidate(self) -> None:
        if self.analysis_result is None:
            return
        source_row_id = self._selected_source_row_id()
        if source_row_id is None:
            return
        candidate_selection = self.candidate_tree.selection()
        if not candidate_selection:
            messagebox.showwarning("Review", "Select a candidate to accept.")
            return

        rank = int(candidate_selection[0].split(":")[1])
        candidate_df = self.analysis_result.candidates_df[
            (self.analysis_result.candidates_df["source_row_id"] == source_row_id)
            & (self.analysis_result.candidates_df["rank"] == rank)
        ]
        if candidate_df.empty:
            return
        candidate = candidate_df.iloc[0]
        row_mask = self.analysis_result.results_df["source_row_id"] == source_row_id
        self.manual_sequence += 1
        self.analysis_result.results_df.loc[row_mask, "manual_status"] = "ACCEPT"
        self.analysis_result.results_df.loc[row_mask, "manual_match_t2_original"] = candidate["nome_t2_original"]
        self.analysis_result.results_df.loc[row_mask, "manual_match_t2_norm"] = candidate["nome_t2_norm"]
        self.analysis_result.results_df.loc[row_mask, "manual_line_match_t2"] = candidate["excel_row_t2"]
        self.analysis_result.results_df.loc[row_mask, "manual_score"] = candidate["score"]
        self.analysis_result.results_df.loc[row_mask, "manual_note"] = self.manual_note_var.get().strip()
        self.analysis_result.results_df.loc[row_mask, "manual_sequence"] = self.manual_sequence
        recompute_final_state(self.analysis_result.results_df, self.catalog_df)
        self.refresh_analysis_views()
        self.refresh_review_views()
        self.refresh_export_view()
        self.log(f"Manual accept applied to row {source_row_id} using candidate rank {rank}.")

    def mark_selected_no_match(self) -> None:
        if self.analysis_result is None:
            return
        source_row_id = self._selected_source_row_id()
        if source_row_id is None:
            return
        row_mask = self.analysis_result.results_df["source_row_id"] == source_row_id
        self.manual_sequence += 1
        self.analysis_result.results_df.loc[row_mask, "manual_status"] = "NO_MATCH"
        self.analysis_result.results_df.loc[row_mask, "manual_match_t2_original"] = ""
        self.analysis_result.results_df.loc[row_mask, "manual_match_t2_norm"] = ""
        self.analysis_result.results_df.loc[row_mask, "manual_line_match_t2"] = pd.NA
        self.analysis_result.results_df.loc[row_mask, "manual_score"] = pd.NA
        self.analysis_result.results_df.loc[row_mask, "manual_note"] = self.manual_note_var.get().strip()
        self.analysis_result.results_df.loc[row_mask, "manual_sequence"] = self.manual_sequence
        recompute_final_state(self.analysis_result.results_df, self.catalog_df)
        self.refresh_analysis_views()
        self.refresh_review_views()
        self.refresh_export_view()
        self.log(f"Row {source_row_id} manually marked as no match.")

    def keep_selected_in_review(self) -> None:
        if self.analysis_result is None:
            return
        source_row_id = self._selected_source_row_id()
        if source_row_id is None:
            return
        row_mask = self.analysis_result.results_df["source_row_id"] == source_row_id
        self.manual_sequence += 1
        self.analysis_result.results_df.loc[row_mask, "manual_status"] = "REVIEW"
        self.analysis_result.results_df.loc[row_mask, "manual_note"] = self.manual_note_var.get().strip()
        self.analysis_result.results_df.loc[row_mask, "manual_sequence"] = self.manual_sequence
        recompute_final_state(self.analysis_result.results_df, self.catalog_df)
        self.refresh_analysis_views()
        self.refresh_review_views()
        self.refresh_export_view()
        self.log(f"Row {source_row_id} kept in manual review.")

    def reset_manual_decision(self) -> None:
        if self.analysis_result is None:
            return
        source_row_id = self._selected_source_row_id()
        if source_row_id is None:
            return
        row_mask = self.analysis_result.results_df["source_row_id"] == source_row_id
        for column in [
            "manual_status",
            "manual_match_t2_original",
            "manual_match_t2_norm",
            "manual_note",
        ]:
            self.analysis_result.results_df.loc[row_mask, column] = ""
        self.analysis_result.results_df.loc[row_mask, "manual_line_match_t2"] = pd.NA
        self.analysis_result.results_df.loc[row_mask, "manual_score"] = pd.NA
        self.analysis_result.results_df.loc[row_mask, "manual_sequence"] = pd.NA
        recompute_final_state(self.analysis_result.results_df, self.catalog_df)
        self.refresh_analysis_views()
        self.refresh_review_views()
        self.refresh_export_view()
        self.log(f"Manual decision reset for row {source_row_id}.")

    def refresh_export_view(self) -> None:
        self.export_text.delete("1.0", "end")
        if self.analysis_result is None:
            self.export_text.insert("1.0", "Run an analysis to populate export details.")
            return

        recompute_final_state(self.analysis_result.results_df, self.catalog_df)
        summary_df = build_summary(self.analysis_result.results_df)
        quota_df = build_quota_summary(self.analysis_result.results_df, self.catalog_df)
        unresolved = int((self.analysis_result.results_df["final_status"] == "REVISAR").sum())
        conflict_count = int((self.analysis_result.results_df["final_conflict_flags"].fillna("") != "").sum())
        lines = ["Current export state", "==================", ""]
        for row in summary_df.itertuples(index=False):
            lines.append(f"{row.status}: {row.quantidade} ({row.percentual:.2f}%)")
        lines.extend(
            [
                "",
                f"Rows still in review: {unresolved}",
                f"Rows with conflict flags: {conflict_count}",
                f"Output workbook: {self.vars['output_file'].get()}",
                "",
                "Top filled quotas:",
            ]
        )
        if quota_df.empty:
            lines.append("- No T2 catalog available.")
        else:
            top_quota = quota_df.sort_values(["accepted_count", "nome_t2_original"], ascending=[False, True]).head(10)
            for row in top_quota.itertuples(index=False):
                lines.append(
                    f"- {row.nome_t2_original}: {row.accepted_count}/{row.quota_limit} accepted"
                )
        self.export_text.insert("1.0", "\n".join(lines))

    def start_export(self) -> None:
        if self.analysis_result is None:
            messagebox.showwarning("Export", "Run analysis before exporting.")
            return
        try:
            config = validate_config(self.collect_config_from_vars(), validate_workbook=False)
            self.analysis_result.config["output_file"] = config["output_file"]
        except Exception as exc:
            messagebox.showerror("Export validation", str(exc))
            return

        self.set_busy(True)
        self.set_status("Exporting workbook...", 0)
        self.log("Starting export.")

        def worker() -> None:
            try:
                output_path = export_analysis_result(
                    self.analysis_result,
                    output_file=self.analysis_result.config["output_file"],
                    progress_callback=self.safe_progress,
                )
                self.root.after(0, lambda: self.on_export_success(output_path))
            except Exception as exc:
                tb = traceback.format_exc()
                self.root.after(0, lambda: self.on_background_error(str(exc), tb))

        threading.Thread(target=worker, daemon=True).start()

    def on_export_success(self, output_path: Path) -> None:
        self.set_busy(False)
        self.last_output_path = output_path
        self.refresh_export_view()
        self.set_status(f"Export completed: {output_path}", 100)
        self.log(f"Export completed: {output_path}")
        if self.vars["auto_open_output"].get():
            try:
                open_file_with_default_app(output_path)
                self.log("Exported workbook opened automatically.")
            except Exception as exc:
                self.log(f"Could not auto-open workbook: {exc}")
        messagebox.showinfo("Export complete", f"Workbook generated:\n\n{output_path}")

    def open_last_output(self) -> None:
        if not self.last_output_path:
            messagebox.showwarning("Open export", "No exported workbook is available yet.")
            return
        try:
            open_file_with_default_app(self.last_output_path)
        except Exception as exc:
            messagebox.showerror("Open export", str(exc))


def main() -> None:
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    app = MatcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
