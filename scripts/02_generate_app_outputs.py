"""
02_generate_app_outputs.py

Generate all precomputed outputs used by the GitHub Pages app.

This script reads the cleaned files created by scripts/01_prepare_dataset.py and
writes analysis-ready JSON files to data/app/. The web app must only read these
JSON files and render tables/charts; it should not calculate descriptive
statistics, crosstabs, chi-square or Cramer's V in the browser.

Run from the repository root:
    python scripts/01_prepare_dataset.py --input data/raw/questionario_original.xlsx
    python scripts/02_generate_app_outputs.py
"""

from __future__ import annotations

import argparse
import ast
import json
import math
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
from scipy.stats import chi2_contingency, fisher_exact

ROOT = Path(__file__).resolve().parents[1]
PROCESSED_DIR = ROOT / "data" / "processed"
APP_DATA_DIR = ROOT / "data" / "app"

MISSING_LABEL = "Sem resposta"
ALPHA = 0.05


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate precomputed JSON outputs for the static app.")
    parser.add_argument("--processed-dir", type=Path, default=PROCESSED_DIR)
    parser.add_argument("--app-dir", type=Path, default=APP_DATA_DIR)
    parser.add_argument("--alpha", type=float, default=ALPHA)
    return parser.parse_args()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def to_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    return json.loads(df.replace({np.nan: None}).to_json(orient="records", force_ascii=False))


def is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    if isinstance(value, str) and value.strip().lower() in {"nan", "none", "null"}:
        return True
    return False


def clean_category(value: Any, include_missing: bool = False) -> Optional[str]:
    if is_missing(value):
        return MISSING_LABEL if include_missing else None
    return str(value).strip()


def parse_listish(value: Any) -> List[str]:
    if isinstance(value, list):
        return [str(v) for v in value]
    if is_missing(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    try:
        parsed = ast.literal_eval(text)
        if isinstance(parsed, list):
            return [str(v) for v in parsed]
    except Exception:
        pass
    return [text]


def strip_question_prefix(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return text.split(".", 1)[1].strip() if "." in text[:8] else text.strip()


def question_order_key(qid: str) -> Tuple[str, int, int, str]:
    import re

    match = re.match(r"^([A-E])(\d+)(?:\.(\d+))?([a-z]?)$", str(qid))
    if not match:
        return ("Z", 999, 999, str(qid))
    block, main, sub, suffix = match.groups()
    return (block, int(main), int(sub or 0), suffix or "")


def round_or_none(value: Any, digits: int = 2) -> Optional[float]:
    if value is None:
        return None
    try:
        if np.isnan(value):
            return None
    except Exception:
        pass
    return round(float(value), digits)


def compute_single_question(
    clean: pd.DataFrame,
    question: Dict[str, Any],
    include_missing: bool = False,
) -> Optional[Dict[str, Any]]:
    variables = parse_listish(question.get("variables"))
    if not variables:
        return None
    variable = variables[0]
    if variable not in clean.columns:
        return None

    total_n = len(clean)
    series = clean[variable].map(lambda v: clean_category(v, include_missing=include_missing))
    valid = series.dropna()
    valid_n = int(valid.shape[0])
    missing_n = int(total_n - valid_n)

    counts = valid.value_counts(dropna=False)
    rows = []
    for category, n in counts.items():
        n = int(n)
        rows.append(
            {
                "category": str(category),
                "n": n,
                "percent_valid": round_or_none((n / valid_n) * 100 if valid_n else None),
                "percent_total": round_or_none((n / total_n) * 100 if total_n else None),
            }
        )

    rows = sorted(rows, key=lambda r: (-r["n"], r["category"]))

    return {
        "question_id": question.get("question_id"),
        "question_text": question.get("question_text"),
        "question_label": strip_question_prefix(question.get("question_text", "")),
        "app_section": question.get("app_section"),
        "objective": question.get("objective"),
        "variable_type": "single_choice_or_text",
        "variable": variable,
        "default_chart": question.get("default_chart", "bar"),
        "summary": {
            "total_n": total_n,
            "valid_n": valid_n,
            "missing_n": missing_n,
            "n_categories": len(rows),
        },
        "rows": rows,
    }


def compute_multiple_question(
    multiple: pd.DataFrame,
    question: Dict[str, Any],
) -> Optional[Dict[str, Any]]:
    qid = question.get("question_id")
    subset = multiple[multiple["question_id"].astype(str) == str(qid)].copy()
    if subset.empty:
        return None

    subset["selected_num"] = pd.to_numeric(subset["selected"], errors="coerce")
    total_respondents = int(subset["respondent_id"].nunique())

    rows = []
    for option, group in subset.groupby("option", dropna=False):
        valid = group["selected_num"].dropna()
        denominator = int(valid.shape[0])
        selected_n = int((valid == 1).sum())
        rows.append(
            {
                "option": str(option),
                "selected_n": selected_n,
                "valid_n": denominator,
                "percent_valid": round_or_none((selected_n / denominator) * 100 if denominator else None),
                "percent_total_respondents": round_or_none((selected_n / total_respondents) * 100 if total_respondents else None),
            }
        )

    rows = sorted(rows, key=lambda r: (-r["selected_n"], r["option"]))

    return {
        "question_id": qid,
        "question_text": question.get("question_text"),
        "question_label": strip_question_prefix(question.get("question_text", "")),
        "app_section": question.get("app_section"),
        "objective": question.get("objective"),
        "variable_type": "multiple_choice",
        "default_chart": question.get("default_chart", "horizontal_bar"),
        "summary": {
            "total_respondents": total_respondents,
            "n_options": len(rows),
            "total_selected": int(sum(r["selected_n"] for r in rows)),
        },
        "rows": rows,
    }


def cramers_v(chi2_value: float, n: int, rows: int, cols: int) -> Optional[float]:
    min_dim = min(rows - 1, cols - 1)
    if n <= 0 or min_dim <= 0:
        return None
    return math.sqrt(chi2_value / (n * min_dim))


def effect_label(v: Optional[float]) -> str:
    if v is None:
        return "Não aplicável"
    if v < 0.10:
        return "Muito fraco"
    if v < 0.30:
        return "Fraco"
    if v < 0.50:
        return "Moderado"
    return "Forte"


def table_to_rows(df: pd.DataFrame, row_name: str = "row_category") -> List[Dict[str, Any]]:
    rows = []
    for idx, values in df.iterrows():
        record = {row_name: str(idx)}
        for col, value in values.items():
            record[str(col)] = None if pd.isna(value) else (round(float(value), 2) if isinstance(value, float) else int(value))
        rows.append(record)
    return rows


def matrix_to_long(counts: pd.DataFrame, row_pct: pd.DataFrame) -> List[Dict[str, Any]]:
    rows = []
    for row_cat in counts.index:
        for col_cat in counts.columns:
            rows.append(
                {
                    "row_category": str(row_cat),
                    "column_category": str(col_cat),
                    "n": int(counts.loc[row_cat, col_cat]),
                    "row_percent": round_or_none(row_pct.loc[row_cat, col_cat]),
                }
            )
    return rows


def compute_crosstab(
    clean: pd.DataFrame,
    crosstab_def: Dict[str, Any],
    question_lookup: Dict[str, Dict[str, Any]],
    alpha: float,
) -> Optional[Dict[str, Any]]:
    x_var = crosstab_def.get("x")
    y_var = crosstab_def.get("y")
    if x_var not in clean.columns or y_var not in clean.columns:
        return None

    tmp = clean[[x_var, y_var]].copy()
    tmp[x_var] = tmp[x_var].map(lambda v: clean_category(v, include_missing=False))
    tmp[y_var] = tmp[y_var].map(lambda v: clean_category(v, include_missing=False))
    tmp = tmp.dropna()

    valid_n = int(len(tmp))
    if valid_n == 0:
        return None

    counts = pd.crosstab(tmp[x_var], tmp[y_var])
    if counts.empty:
        return None

    chi2_value = None
    p_value = None
    dof = None
    expected = None
    expected_lt5 = None
    expected_lt5_percent = None
    fisher_p = None
    fisher_odds_ratio = None
    warning = None

    if counts.shape[0] >= 2 and counts.shape[1] >= 2:
        chi2_value, p_value, dof, expected = chi2_contingency(counts.values, correction=False)
        expected_df = pd.DataFrame(expected, index=counts.index, columns=counts.columns)
        expected_lt5 = int((expected_df < 5).sum().sum())
        expected_cells = int(expected_df.size)
        expected_lt5_percent = (expected_lt5 / expected_cells) * 100 if expected_cells else None
        if expected_lt5 > 0:
            warning = (
                f"Há {expected_lt5} células com frequência esperada inferior a 5 "
                f"({round_or_none(expected_lt5_percent)}% das células). Interpretar o Qui-quadrado com cautela."
            )
        if counts.shape == (2, 2):
            fisher_odds_ratio, fisher_p = fisher_exact(counts.values)
    else:
        expected_df = pd.DataFrame(index=counts.index, columns=counts.columns)
        warning = "Tabela com menos de duas categorias em uma das variáveis; teste Qui-quadrado não aplicável."

    row_pct = counts.div(counts.sum(axis=1), axis=0) * 100
    col_pct = counts.div(counts.sum(axis=0), axis=1) * 100
    total_n = int(counts.values.sum())
    cv = cramers_v(float(chi2_value), total_n, counts.shape[0], counts.shape[1]) if chi2_value is not None else None

    return {
        "id": crosstab_def.get("id"),
        "title": crosstab_def.get("title"),
        "x_variable": x_var,
        "y_variable": y_var,
        "test": crosstab_def.get("test", "chi_square"),
        "summary": {
            "valid_n": valid_n,
            "rows": int(counts.shape[0]),
            "columns": int(counts.shape[1]),
        },
        "statistics": {
            "chi_square": round_or_none(chi2_value, 4),
            "degrees_of_freedom": None if dof is None else int(dof),
            "p_value": round_or_none(p_value, 6),
            "alpha": alpha,
            "significant": bool(p_value < alpha) if p_value is not None else None,
            "cramers_v": round_or_none(cv, 4),
            "effect_size_label": effect_label(cv),
            "fisher_exact_p_value": round_or_none(fisher_p, 6),
            "fisher_exact_odds_ratio": round_or_none(fisher_odds_ratio, 4),
            "expected_lt5_cells": expected_lt5,
            "expected_lt5_percent": round_or_none(expected_lt5_percent),
        },
        "warning": warning,
        "counts_table": table_to_rows(counts),
        "row_percent_table": table_to_rows(row_pct),
        "column_percent_table": table_to_rows(col_pct),
        "expected_table": table_to_rows(expected_df) if expected is not None else [],
        "plot_rows": matrix_to_long(counts, row_pct),
    }


def compute_dashboard_cards(clean: pd.DataFrame, config: Dict[str, Any]) -> List[Dict[str, Any]]:
    cards = []
    total_n = int(len(clean))

    for card in config.get("dashboard_cards", []):
        out = dict(card)
        if card.get("type") == "count_rows":
            out.update({"value": total_n, "display_value": f"{total_n}", "note": "N total"})
        elif card.get("type") == "share_equals":
            var = card.get("variable")
            target = card.get("value")
            if var in clean.columns:
                valid = clean[var].dropna().map(lambda v: str(v).strip())
                valid_n = int(valid.shape[0])
                count = int((valid == str(target)).sum())
                pct = (count / valid_n) * 100 if valid_n else None
                out.update(
                    {
                        "n": count,
                        "valid_n": valid_n,
                        "percent_valid": round_or_none(pct),
                        "display_value": f"{round_or_none(pct)}%" if pct is not None else "—",
                        "note": f"{count}/{valid_n} válidas" if valid_n else "Sem dados válidos",
                    }
                )
            else:
                out.update({"display_value": "—", "note": "Variável não encontrada"})
        cards.append(out)
    return cards


def main() -> None:
    args = parse_args()
    processed_dir = args.processed_dir
    app_dir = args.app_dir
    app_dir.mkdir(parents=True, exist_ok=True)

    # Remove stale files that would suggest browser-side calculations or duplicate catalogs.
    for stale_name in [
        "survey_clean.json",
        "multiple_choice_long.json",
        "question_catalog.json",
        "question_catalog_app.json",
        "question_catalog_full.json",
    ]:
        stale_path = app_dir / stale_name
        if stale_path.exists():
            stale_path.unlink()

    clean = pd.read_csv(processed_dir / "questionario_clean_wide.csv")
    multiple = pd.read_csv(processed_dir / "multiple_choice_long.csv")
    questions = pd.read_csv(processed_dir / "question_catalog_app.csv")
    config = read_json(app_dir / "app_config.json")

    question_records = to_records(questions)
    question_lookup = {str(q["question_id"]): q for q in question_records}

    descriptive_results = []
    multiple_results = []
    for q in sorted(question_records, key=lambda item: question_order_key(item.get("question_id", ""))):
        if q.get("variable_type") == "multiple_choice":
            result = compute_multiple_question(multiple, q)
            if result is not None:
                multiple_results.append(result)
        else:
            result = compute_single_question(clean, q)
            if result is not None:
                descriptive_results.append(result)

    crosstab_results = []
    for ct in config.get("crosstabs", []):
        result = compute_crosstab(clean, ct, question_lookup, alpha=args.alpha)
        if result is not None:
            crosstab_results.append(result)

    dashboard_cards = compute_dashboard_cards(clean, config)

    manifest = {
        "generated_by": "scripts/02_generate_app_outputs.py",
        "source_files": {
            "clean_wide": "data/processed/questionario_clean_wide.csv",
            "multiple_choice_long": "data/processed/multiple_choice_long.csv",
            "questions": "data/processed/question_catalog_app.csv",
            "config": "data/app/app_config.json",
        },
        "outputs": {
            "questions": "data/app/questions.json",
            "dashboard_cards": "data/app/dashboard_cards.json",
            "descriptive_results": "data/app/descriptive_results.json",
            "multiple_choice_results": "data/app/multiple_choice_results.json",
            "crosstab_results": "data/app/crosstab_results.json",
        },
        "notes": [
            "A app web apenas lê estes JSON e renderiza tabelas/gráficos.",
            "As frequências, percentagens, Qui-quadrado e Cramer's V são calculados neste script Python.",
        ],
    }

    write_json(app_dir / "questions.json", question_records)
    write_json(app_dir / "dashboard_cards.json", dashboard_cards)
    write_json(app_dir / "descriptive_results.json", descriptive_results)
    write_json(app_dir / "multiple_choice_results.json", multiple_results)
    write_json(app_dir / "crosstab_results.json", crosstab_results)
    write_json(app_dir / "app_manifest.json", manifest)

    # CSV copies for auditing outside the app.
    flat_single = []
    for item in descriptive_results:
        for row in item["rows"]:
            flat_single.append({"question_id": item["question_id"], "question_text": item["question_text"], **row})
    pd.DataFrame(flat_single).to_csv(processed_dir / "descriptive_results.csv", index=False, encoding="utf-8-sig")

    flat_multi = []
    for item in multiple_results:
        for row in item["rows"]:
            flat_multi.append({"question_id": item["question_id"], "question_text": item["question_text"], **row})
    pd.DataFrame(flat_multi).to_csv(processed_dir / "multiple_choice_results.csv", index=False, encoding="utf-8-sig")

    flat_ct = []
    for item in crosstab_results:
        flat_ct.append({"id": item["id"], "title": item["title"], **item["summary"], **item["statistics"], "warning": item["warning"]})
    pd.DataFrame(flat_ct).to_csv(processed_dir / "crosstab_summary.csv", index=False, encoding="utf-8-sig")

    print("App outputs generated in:", app_dir)
    print("Descriptive questions:", len(descriptive_results))
    print("Multiple-choice questions:", len(multiple_results))
    print("Crosstabs:", len(crosstab_results))


if __name__ == "__main__":
    main()
