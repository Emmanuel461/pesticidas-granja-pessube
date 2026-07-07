#!/usr/bin/env python3
"""
01_prepare_dataset.py

Clean and reorganize the Kobo/Excel questionnaire dataset for the pesticide
perception study in Granja de Pessubé.

Main outputs:
- data/processed/questionario_clean.xlsx
- data/processed/questionario_clean_wide.csv
- data/processed/single_choice_wide.csv
- data/processed/multiple_choice_long.csv
- data/processed/question_catalog.csv
- data/processed/question_catalog_app.csv
- data/processed/value_audit.csv
- data/app/questions.json
- data/app/app_config.json

This script prepares the clean data and metadata only.
It does not generate final analysis outputs.
Run scripts/02_generate_app_outputs.py afterwards.

Run from the repository root:
    python scripts/01_prepare_dataset.py \
        --input data/raw/questionario_original.xlsx
"""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------
# Project paths
# ---------------------------------------------------------------------
ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = ROOT / "data" / "raw"
PROCESSED_DIR = ROOT / "data" / "processed"
APP_DATA_DIR = ROOT / "data" / "app"


# ---------------------------------------------------------------------
# Friendly variable names based on questionnaire codes
# ---------------------------------------------------------------------
FRIENDLY_NAMES: Dict[str, str] = {
    "A0": "nome",
    "A1": "bairro",
    "A2": "associacao",
    "A3": "escolaridade",
    "A4": "idade",
    "A5": "anos_horticultura",
    "A6": "formacao_pesticidas",
    "A6a": "quem_ministrou_formacao",
    "B7": "culturas",
    "B8": "organizacao_canteiro",
    "B9": "rotacao_culturas",
    "B10": "conhece_pesticidas",
    "B11": "usa_pesticidas",
    "B12": "tipos_pesticidas",
    "B13": "local_compra_pesticidas",
    "B14": "momento_uso_pesticidas",
    "B15": "duracao_aplicacao",
    "B16": "frequencia_aplicacao",
    "B17": "criterios_escolha_pesticidas",
    "B18": "usa_pesticidas_outras_culturas",
    "B19": "usa_produtos_naturais",
    "B19.1": "tipos_produtos_naturais",
    "B19.2": "quem_ensinou_produtos_naturais",
    "B19.3": "eficacia_produtos_naturais",
    "B20": "agricultura_sem_pesticidas_quimicos",
    "B21": "finalidade_dinheiro",
    "C21": "conhece_regras_seguranca",
    "C21a": "fonte_regras_seguranca",
    "C22": "conhece_pesticidas_proibidos_restritos",
    "C22.1": "sabe_quais_pesticidas_proibidos",
    "C22.2": "razao_pesticidas_proibidos_restritos",
    "C23": "aplicador_pesticidas",
    "C24": "le_rotulos",
    "C24.1": "porque_nao_le_rotulos",
    "C25": "armazenamento_pesticidas",
    "C26": "local_preparo_solucao",
    "C27": "reutiliza_recipientes",
    "C27.1": "finalidade_reutilizacao_recipientes",
    "C28": "destino_sobras_pesticidas",
    "C29": "destino_pesticidas_vencidos",
    "C30": "destino_embalagens_vazias",
    "C31": "usa_epi",
    "C31.1": "frequencia_uso_epi",
    "D32": "percepcao_impacto_saude",
    "D32.1": "sintomas_apos_aplicacao",
    "D32.2": "sintomas_sentidos",
    "D33": "percepcao_impacto_ambiente",
    "D33.1": "impactos_ambientais_percebidos",
    "D33.2": "razoes_continuar_usando_pesticidas",
    "E34": "quer_mais_formacao",
    "E34.1": "temas_formacao",
    "E35": "formas_sensibilizacao",
    "E36": "conselhos_audio",
}

BLOCK_NAMES: Dict[str, str] = {
    "A": "Caracterizacao pessoal",
    "B": "Praticas horticolas e uso de pesticidas",
    "C": "Conhecimento sobre seguranca",
    "D": "Percepcoes sobre saude e ambiente",
    "E": "Educacao ambiental",
}

# Columns that should not be published in the anonymized clean data.
# Keep the raw file locally in data/raw/ and do not push it to GitHub.
IDENTIFIER_PATTERNS = [
    r"^A0\.",
    r"^_id$",
    r"^_uuid$",
    r"^_submission_time$",
    r"^_submitted_by$",
    r"^meta/rootUuid$",
    r"_URL$",
    r"\.m4a$",
]

IDENTIFIER_QUESTION_IDS = {"A0", "E36"}

METADATA_COLUMNS = {
    "start",
    "end",
    "_validation_status",
    "_notes",
    "_status",
    "__version__",
    "_tags",
    "_index",
}


# ---------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------
def normalize_spaces(value: Any) -> Any:
    """Strip and collapse repeated spaces in string values."""
    if pd.isna(value):
        return np.nan
    if isinstance(value, str):
        value = value.replace("\xa0", " ")
        value = re.sub(r"\s+", " ", value).strip()
        return value
    return value


def clean_answer(value: Any, variable_name: Optional[str] = None) -> Any:
    """Clean individual cell values while preserving numeric information."""
    value = normalize_spaces(value)

    if pd.isna(value):
        return np.nan

    # Specific Kobo/export artefact found in the age question.
    if isinstance(value, str):
        value = re.sub(r"^Option\s+[a-z]\.?\s*", "", value, flags=re.IGNORECASE).strip()

        # Standardize common categorical labels.
        replacements = {
            "Sim": "Sim",
            "Não": "Não",
            "Nao": "Não",
            "Não sabe": "Não sabe",
            "Nao sabe": "Não sabe",
            "Outra": "Outro",
            "Outros": "Outro",
            "várias": "Várias",
            "varias": "Várias",
        }
        value = replacements.get(value, value)

    # In this export, B8 has two records coded as 1. Treat them as a valid category.
    if variable_name == "organizacao_canteiro" and value == 1:
        return "Uma cultura"

    return value


def strip_duplicate_suffix(text: str) -> str:
    """Remove pandas duplicate suffixes such as '.1' at the end of labels."""
    return re.sub(r"\.\d+$", "", text).strip()


def slugify(text: Any, max_len: int = 70) -> str:
    """Convert text into a safe snake_case string."""
    text = "" if pd.isna(text) else str(text)
    text = strip_duplicate_suffix(text)
    text = normalize_spaces(text)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = text.replace("ç", "c")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:max_len].strip("_") or "sem_nome"


def extract_question_id(label: str) -> Optional[str]:
    """Extract questionnaire ID, for example A1, A6a, B19.1 or D33.2."""
    label = normalize_spaces(label)
    if not isinstance(label, str):
        return None
    label = label.split("/", 1)[0].strip()
    match = re.match(r"^([A-E]\d+(?:\.\d+)?[a-z]?)\.", label)
    return match.group(1) if match else None


def get_block(question_id: Optional[str]) -> Optional[str]:
    """Return the block name from the question code."""
    if not question_id:
        return None
    return BLOCK_NAMES.get(question_id[0])


def friendly_base_name(question_id: Optional[str], original_label: str) -> str:
    """Return a readable base variable name."""
    if question_id and question_id in FRIENDLY_NAMES:
        return f"{question_id.lower().replace('.', '_')}_{FRIENDLY_NAMES[question_id]}"
    if question_id:
        return f"{question_id.lower().replace('.', '_')}_{slugify(original_label, 45)}"
    return slugify(original_label, 60)


def is_identifier_column(col: str) -> bool:
    """Detect personal identifiers or technical IDs that should not be published."""
    for pattern in IDENTIFIER_PATTERNS:
        if re.search(pattern, col):
            return True
    return False


def is_empty_or_block_column(series: pd.Series, col: str) -> bool:
    """Detect empty Kobo block/header columns."""
    if series.notna().sum() == 0:
        if col.startswith("BLOCO") or col.startswith("VAR"):
            return True
        # Also drop totally empty duplicated structural columns.
        return True
    return False


def parse_option_column(col: str) -> Optional[Tuple[str, str, str, str]]:
    """
    Parse a multiple-choice option column.

    Returns:
        (question_id, parent_label, option_label, base_variable_name)
    """
    if "/" not in col:
        return None

    parent_label, option_label = col.split("/", 1)
    parent_label = normalize_spaces(parent_label)
    option_label = strip_duplicate_suffix(normalize_spaces(option_label))
    question_id = extract_question_id(parent_label)

    if question_id is None:
        return None

    base_name = friendly_base_name(question_id, parent_label)
    return question_id, parent_label, option_label, base_name


def is_binary_option(series: pd.Series) -> bool:
    """Check whether a column behaves like a 0/1 multiple-choice option."""
    values = pd.Series(series.dropna().unique())
    if values.empty:
        return False
    numeric_values = pd.to_numeric(values, errors="coerce")
    if numeric_values.isna().any():
        return False
    return set(numeric_values.astype(float).unique()).issubset({0.0, 1.0})


# ---------------------------------------------------------------------
# Main processing functions
# ---------------------------------------------------------------------
def read_input_excel(input_path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Read the first sheet or a named sheet from the input workbook."""
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if sheet_name:
        return pd.read_excel(input_path, sheet_name=sheet_name)

    excel = pd.ExcelFile(input_path)
    return pd.read_excel(input_path, sheet_name=excel.sheet_names[0])


def build_catalog_and_clean_data(df_raw: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Build clean wide data, single-choice data, multiple-choice long data and metadata.
    """
    df = df_raw.copy()
    df.columns = [normalize_spaces(c) for c in df.columns]

    # Create a safe respondent ID. Do not expose names or UUIDs in public outputs.
    df.insert(0, "respondent_id", [f"R{i:03d}" for i in range(1, len(df) + 1)])

    catalog_rows: List[Dict[str, Any]] = []
    dropped_rows: List[Dict[str, Any]] = []
    duplicate_merge_rows: List[Dict[str, Any]] = []

    clean_wide = pd.DataFrame({"respondent_id": df["respondent_id"]})
    single_choice_wide = pd.DataFrame({"respondent_id": df["respondent_id"]})

    # Multiple-choice option columns can be duplicated across Kobo versions.
    # We group by question + option and combine them row-wise.
    option_groups: Dict[Tuple[str, str], List[str]] = {}
    option_group_meta: Dict[Tuple[str, str], Dict[str, Any]] = {}

    for col in df.columns:
        if col == "respondent_id":
            continue

        question_id = extract_question_id(col)
        block = get_block(question_id)
        option_info = parse_option_column(col)
        non_null = int(df[col].notna().sum())

        if is_empty_or_block_column(df[col], col):
            dropped_rows.append({
                "original_column": col,
                "reason": "empty_or_block_header",
                "non_null": non_null,
            })
            catalog_rows.append({
                "original_column": col,
                "clean_variable": None,
                "question_id": question_id,
                "block": block,
                "question_text": col,
                "option": None,
                "variable_type": "empty_or_block_header",
                "included_clean_wide": False,
                "non_null": non_null,
            })
            continue

        if col in METADATA_COLUMNS or is_identifier_column(col) or question_id in IDENTIFIER_QUESTION_IDS:
            dropped_rows.append({
                "original_column": col,
                "reason": "identifier_or_metadata",
                "non_null": non_null,
            })
            catalog_rows.append({
                "original_column": col,
                "clean_variable": None,
                "question_id": question_id,
                "block": block,
                "question_text": col,
                "option": None,
                "variable_type": "identifier_or_metadata",
                "included_clean_wide": False,
                "non_null": non_null,
            })
            continue

        if option_info and is_binary_option(df[col]):
            qid, parent_label, option_label, base_name = option_info
            key = (qid, slugify(option_label, 60))
            option_groups.setdefault(key, []).append(col)
            option_group_meta[key] = {
                "question_id": qid,
                "block": get_block(qid),
                "question_text": parent_label,
                "option": option_label,
                "base_name": base_name,
            }
            option_slug = slugify(option_label, 60)
            clean_name = f"{base_name}__{option_slug}"
            catalog_rows.append({
                "original_column": col,
                "clean_variable": clean_name,
                "question_id": qid,
                "block": get_block(qid),
                "question_text": parent_label,
                "option": option_label,
                "variable_type": "multiple_choice_option",
                "included_clean_wide": True,
                "non_null": non_null,
            })
            continue

        # Single-choice or open-text parent variable.
        base_name = friendly_base_name(question_id, col)
        clean_name = base_name
        if clean_name in clean_wide.columns:
            clean_name = f"{clean_name}_{len(clean_wide.columns)}"

        cleaned_series = df[col].map(lambda x: clean_answer(x, FRIENDLY_NAMES.get(question_id or "", None)))
        clean_wide[clean_name] = cleaned_series
        single_choice_wide[clean_name] = cleaned_series

        variable_type = "single_or_text"
        if "/" not in col and any(opt_col.startswith(f"{col}/") for opt_cols in option_groups.values() for opt_col in opt_cols):
            variable_type = "multiple_choice_parent_text"

        catalog_rows.append({
            "original_column": col,
            "clean_variable": clean_name,
            "question_id": question_id,
            "block": block,
            "question_text": col,
            "option": None,
            "variable_type": variable_type,
            "included_clean_wide": True,
            "non_null": non_null,
        })

    # Combine multiple-choice option columns and add them to clean_wide.
    multiple_long_rows: List[pd.DataFrame] = []
    for (qid, option_slug), cols in option_groups.items():
        meta = option_group_meta[(qid, option_slug)]
        clean_name = f"{meta['base_name']}__{option_slug}"

        option_matrix = df[cols].apply(pd.to_numeric, errors="coerce")
        combined = option_matrix.max(axis=1, skipna=True)
        combined = combined.where(option_matrix.notna().any(axis=1), np.nan)

        clean_wide[clean_name] = combined.astype("Int64")

        if len(cols) > 1:
            duplicate_merge_rows.append({
                "question_id": qid,
                "option": meta["option"],
                "clean_variable": clean_name,
                "merged_columns": " | ".join(cols),
                "n_columns_merged": len(cols),
            })

        tmp = pd.DataFrame({
            "respondent_id": df["respondent_id"],
            "question_id": qid,
            "block": meta["block"],
            "question_text": meta["question_text"],
            "option": meta["option"],
            "selected": combined.astype("Int64"),
        })
        multiple_long_rows.append(tmp)

    multiple_choice_long = (
        pd.concat(multiple_long_rows, ignore_index=True)
        if multiple_long_rows
        else pd.DataFrame(columns=["respondent_id", "question_id", "block", "question_text", "option", "selected"])
    )

    catalog = pd.DataFrame(catalog_rows)
    dropped = pd.DataFrame(dropped_rows)
    duplicate_merges = pd.DataFrame(duplicate_merge_rows)

    # Add derived analysis variables used later in descriptive/crosstab scripts.
    clean_wide = add_derived_variables(clean_wide)
    single_choice_wide = add_derived_variables(single_choice_wide)

    value_audit = build_value_audit(single_choice_wide)

    return clean_wide, single_choice_wide, multiple_choice_long, catalog, dropped, duplicate_merges, value_audit


def add_derived_variables(df: pd.DataFrame) -> pd.DataFrame:
    """Create simple derived variables needed for later chi-square analyses."""
    df = df.copy()

    if "a5_anos_horticultura" in df.columns:
        years = pd.to_numeric(df["a5_anos_horticultura"], errors="coerce")
        df["derived_experiencia_grupo"] = pd.cut(
            years,
            bins=[-np.inf, 5, 10, 20, np.inf],
            labels=["0-5 anos", "6-10 anos", "11-20 anos", "Mais de 20 anos"],
        )

    if "a4_idade" in df.columns:
        df["derived_idade_grupo"] = df["a4_idade"].map(normalize_spaces)

    if "d32_percepcao_impacto_saude" in df.columns:
        df["derived_percepcao_saude_bin"] = df["d32_percepcao_impacto_saude"].map(
            lambda x: "Sim" if x == "Sim" else ("Não/Não sabe" if pd.notna(x) else np.nan)
        )

    if "d33_percepcao_impacto_ambiente" in df.columns:
        df["derived_percepcao_ambiente_bin"] = df["d33_percepcao_impacto_ambiente"].map(
            lambda x: "Sim" if x == "Sim" else ("Não/Não sabe" if pd.notna(x) else np.nan)
        )

    return df


def build_value_audit(df: pd.DataFrame, max_categories: int = 30) -> pd.DataFrame:
    """Create a compact audit of categorical values for manual review."""
    rows: List[Dict[str, Any]] = []
    for col in df.columns:
        if col == "respondent_id":
            continue
        s = df[col]
        nunique = s.nunique(dropna=True)
        if nunique <= max_categories:
            counts = s.value_counts(dropna=False)
            for value, n in counts.items():
                rows.append({
                    "variable": col,
                    "value": "<NA>" if pd.isna(value) else value,
                    "n": int(n),
                    "percent_total": round(100 * n / len(df), 2) if len(df) else np.nan,
                    "n_unique_variable": int(nunique),
                })
    return pd.DataFrame(rows)


def natural_question_sort_key(question_id: Any) -> Tuple[str, int, float, str]:
    """Sort question IDs such as A1, A6a, B19.1, D33.2 in questionnaire order."""
    qid = "" if pd.isna(question_id) else str(question_id).strip()
    match = re.match(r"^([A-E])(\d+)(?:\.(\d+))?([a-z]?)$", qid, flags=re.IGNORECASE)
    if not match:
        return ("Z", 999, 999.0, qid)
    block, main, decimal, suffix = match.groups()
    sub = float(decimal) if decimal is not None else 0.0
    return (block.upper(), int(main), sub, suffix or "")


def infer_app_section(question_id: Any, block: Any) -> str:
    """Assign each question to a user-facing section in the GitHub Pages app."""
    qid = "" if pd.isna(question_id) else str(question_id).upper().strip()
    block_text = "" if pd.isna(block) else str(block).lower()

    if qid.startswith("A") or "caracterizacao" in block_text:
        return "Caracterização da amostra"
    if qid.startswith("B") or "praticas" in block_text:
        return "Práticas hortícolas"
    if qid.startswith("C") or "seguranca" in block_text:
        return "Conhecimento e segurança"
    if qid.startswith("D") or "saude" in block_text or "ambiente" in block_text:
        return "Saúde e ambiente"
    if qid.startswith("E") or "educacao" in block_text:
        return "Educação ambiental"
    return "Outras perguntas"


def infer_objective(question_id: Any, block: Any) -> str:
    """Link each question to the closest research objective."""
    qid = "" if pd.isna(question_id) else str(question_id).upper().strip()

    if qid.startswith("A"):
        return "Caracterizar a amostra"
    if qid.startswith("B"):
        return "Identificar práticas hortícolas associadas ao uso de pesticidas"
    if qid.startswith("C"):
        return "Averiguar conhecimentos sobre medidas de segurança"
    if qid.startswith("D"):
        return "Compreender impactos percebidos na saúde e no ambiente"
    if qid.startswith("E"):
        return "Propor estratégias de educação ambiental"
    return "Apoiar a análise geral do questionário"


def infer_default_chart(variable_type: str) -> str:
    """Suggest a default chart type for the future web app."""
    return "horizontal_bar" if variable_type == "multiple_choice" else "bar"


def build_app_question_catalog(catalog: pd.DataFrame) -> pd.DataFrame:
    """
    Build a simplified question catalog for the GitHub Pages app.

    The full catalog keeps all original columns for audit. This app catalog keeps
    only usable questions, deduplicates repeated Kobo/Excel option columns and
    groups multiple-choice options under the same question_id.
    """
    usable = catalog.copy()
    usable = usable[
        (usable["included_clean_wide"] == True)
        & usable["clean_variable"].notna()
        & usable["question_id"].notna()
        & (usable["clean_variable"].astype(str).str.strip() != "")
        & (usable["question_id"].astype(str).str.strip() != "")
    ].copy()

    if usable.empty:
        return pd.DataFrame(
            columns=[
                "question_id",
                "question_order",
                "block",
                "app_section",
                "objective",
                "question_text",
                "variable_type",
                "default_chart",
                "n_variables",
                "variables",
                "options",
            ]
        )

    usable["is_multiple"] = usable["variable_type"].eq("multiple_choice_option")
    rows: List[Dict[str, Any]] = []

    for qid, group in usable.groupby("question_id", dropna=True):
        group = group.copy()
        qid_str = str(qid).strip()
        first = group.iloc[0]
        block = first.get("block", "")

        if group["is_multiple"].any():
            var_type = "multiple_choice"
            option_rows = group[group["is_multiple"]].copy()
            option_rows = option_rows.drop_duplicates(
                subset=["question_id", "clean_variable", "option"],
                keep="last",
            )
            variables = option_rows["clean_variable"].astype(str).tolist()
            options = option_rows["option"].fillna("").astype(str).tolist()
            question_text = option_rows.iloc[0].get("question_text", first.get("question_text", ""))
        else:
            var_type = "single_choice_or_text"
            single_rows = group[~group["is_multiple"]].copy()
            single_rows = single_rows.drop_duplicates(subset=["clean_variable"], keep="last")
            variables = single_rows["clean_variable"].astype(str).tolist()
            options = []
            question_text = single_rows.iloc[0].get("question_text", first.get("question_text", ""))

        rows.append(
            {
                "question_id": qid_str,
                "question_order": ".".join(map(str, natural_question_sort_key(qid_str))),
                "block": block,
                "app_section": infer_app_section(qid_str, block),
                "objective": infer_objective(qid_str, block),
                "question_text": question_text,
                "variable_type": var_type,
                "default_chart": infer_default_chart(var_type),
                "n_variables": len(variables),
                "variables": variables,
                "options": options,
            }
        )

    app_catalog = pd.DataFrame(rows)
    app_catalog = app_catalog.sort_values(
        by="question_id",
        key=lambda s: s.map(natural_question_sort_key),
    ).reset_index(drop=True)

    return app_catalog



def build_app_config(total_respondents: int) -> Dict[str, Any]:
    """
    Configuration file for the GitHub Pages app.

    This file tells the app which precomputed JSON files, sections,
    questions and crosstabs should be available in the interface.
    Frequencies, percentages, tables, chi-square and Cramer's V are
    calculated by scripts/02_generate_app_outputs.py.
    """
    return {
        "title": "Análise do questionário sobre pesticidas",
        "subtitle": "Horticultoras de Granja de Pessubé",
        "language": "pt",
        "data_files": {
            "questions": "../data/app/questions.json",
            "dashboard_cards": "../data/app/dashboard_cards.json",
            "descriptive_results": "../data/app/descriptive_results.json",
            "multiple_choice_results": "../data/app/multiple_choice_results.json",
            "crosstab_results": "../data/app/crosstab_results.json"
        },
        "total_respondents": int(total_respondents),
        "sections": [
            {
                "id": "caracterizacao",
                "title": "Caracterização da amostra",
                "objective": "Caracterizar as horticultoras entrevistadas.",
                "questions": ["A1", "A2", "A3", "A4", "A5", "A6", "A6a"],
            },
            {
                "id": "praticas",
                "title": "Práticas hortícolas e uso de pesticidas",
                "objective": "Identificar práticas hortícolas associadas ao uso de pesticidas, os principais pesticidas utilizados e os critérios de seleção.",
                "questions": [
                    "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14",
                    "B15", "B16", "B17", "B18", "B19", "B19.1", "B19.2",
                    "B19.3", "B20", "B21",
                ],
            },
            {
                "id": "seguranca",
                "title": "Conhecimento e segurança no uso de pesticidas",
                "objective": "Averiguar conhecimentos sobre medidas de segurança no manuseio, aplicação, armazenamento e descarte de pesticidas.",
                "questions": [
                    "C21", "C21a", "C22", "C22.1", "C22.2", "C23", "C24",
                    "C24.1", "C25", "C26", "C27", "C28", "C29", "C30",
                    "C31", "C31.1",
                ],
            },
            {
                "id": "saude",
                "title": "Perceção dos impactos na saúde humana",
                "objective": "Compreender os impactos percebidos pelas horticultoras na saúde humana.",
                "questions": ["D32", "D32.1", "D32.2"],
            },
            {
                "id": "ambiente",
                "title": "Perceção dos impactos ambientais",
                "objective": "Compreender os impactos percebidos pelas horticultoras no ambiente.",
                "questions": ["D33", "D33.1", "D33.2"],
            },
            {
                "id": "educacao",
                "title": "Estratégias de educação ambiental",
                "objective": "Apoiar a proposta de estratégias de sensibilização e proteção integrada.",
                "questions": ["E34", "E34.1", "E35"],
            },
            {
                "id": "cruzamentos",
                "title": "Cruzamentos e associação estatística",
                "objective": "Analisar a associação entre escolaridade, idade, experiência, formação e as perceções das horticultoras.",
                "questions": [],
            },
        ],
        "dashboard_cards": [
            {
                "id": "n_respostas",
                "title": "Respostas",
                "type": "count_rows",
            },
            {
                "id": "formacao_sim",
                "title": "Recebeu formação",
                "type": "share_equals",
                "variable": "a6_formacao_pesticidas",
                "value": "Sim",
            },
            {
                "id": "usa_pesticidas_sim",
                "title": "Usa ou já usou pesticidas",
                "type": "share_equals",
                "variable": "b11_usa_pesticidas",
                "value": "Sim",
            },
            {
                "id": "usa_epi_sim",
                "title": "Usa EPI",
                "type": "share_equals",
                "variable": "c31_usa_epi",
                "value": "Sim",
            },
            {
                "id": "percebe_saude_sim",
                "title": "Percebe impactos na saúde",
                "type": "share_equals",
                "variable": "d32_percepcao_impacto_saude",
                "value": "Sim",
            },
            {
                "id": "percebe_ambiente_sim",
                "title": "Percebe impactos ambientais",
                "type": "share_equals",
                "variable": "d33_percepcao_impacto_ambiente",
                "value": "Sim",
            },
        ],
        "crosstabs": [
            {
                "id": "formacao_conhecimento_seguranca",
                "title": "Formação × conhecimento de segurança",
                "x": "a6_formacao_pesticidas",
                "y": "c21_conhece_regras_seguranca",
                "test": "chi_square",
            },
            {
                "id": "formacao_uso_epi",
                "title": "Formação × uso de EPI",
                "x": "a6_formacao_pesticidas",
                "y": "c31_usa_epi",
                "test": "chi_square",
            },
            {
                "id": "formacao_percepcao_saude",
                "title": "Formação × perceção dos impactos na saúde",
                "x": "a6_formacao_pesticidas",
                "y": "d32_percepcao_impacto_saude",
                "test": "chi_square",
            },
            {
                "id": "formacao_percepcao_ambiente",
                "title": "Formação × perceção dos impactos ambientais",
                "x": "a6_formacao_pesticidas",
                "y": "d33_percepcao_impacto_ambiente",
                "test": "chi_square",
            },
            {
                "id": "escolaridade_conhecimento_seguranca",
                "title": "Escolaridade × conhecimento de segurança",
                "x": "a3_escolaridade",
                "y": "c21_conhece_regras_seguranca",
                "test": "chi_square",
            },
            {
                "id": "escolaridade_uso_epi",
                "title": "Escolaridade × uso de EPI",
                "x": "a3_escolaridade",
                "y": "c31_usa_epi",
                "test": "chi_square",
            },
            {
                "id": "escolaridade_percepcao_saude",
                "title": "Escolaridade × perceção dos impactos na saúde",
                "x": "a3_escolaridade",
                "y": "d32_percepcao_impacto_saude",
                "test": "chi_square",
            },
            {
                "id": "escolaridade_percepcao_ambiente",
                "title": "Escolaridade × perceção dos impactos ambientais",
                "x": "a3_escolaridade",
                "y": "d33_percepcao_impacto_ambiente",
                "test": "chi_square",
            },
            {
                "id": "idade_percepcao_saude",
                "title": "Idade × perceção dos impactos na saúde",
                "x": "derived_idade_grupo",
                "y": "d32_percepcao_impacto_saude",
                "test": "chi_square",
            },
            {
                "id": "idade_percepcao_ambiente",
                "title": "Idade × perceção dos impactos ambientais",
                "x": "derived_idade_grupo",
                "y": "d33_percepcao_impacto_ambiente",
                "test": "chi_square",
            },
            {
                "id": "experiencia_percepcao_saude",
                "title": "Experiência em horticultura × perceção dos impactos na saúde",
                "x": "derived_experiencia_grupo",
                "y": "d32_percepcao_impacto_saude",
                "test": "chi_square",
            },
            {
                "id": "experiencia_percepcao_ambiente",
                "title": "Experiência em horticultura × perceção dos impactos ambientais",
                "x": "derived_experiencia_grupo",
                "y": "d33_percepcao_impacto_ambiente",
                "test": "chi_square",
            },
        ],
        "notes": [
            "As frequências, percentagens, tabelas cruzadas, Qui-quadrado e Cramer's V são calculados previamente pelos scripts Python.",
            "O termo recomendado para a interpretação estatística é associação, não influência causal.",
            "Quando houver frequências esperadas inferiores a 5, os scripts geram aviso metodológico para a app apresentar.",
        ],
    }


def dataframe_records_for_json(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Convert a DataFrame to clean JSON records with nulls instead of NaN/<NA>."""
    return json.loads(df.to_json(orient="records", force_ascii=False))


def write_json(path: Path, data: Any) -> None:
    """Write UTF-8 JSON with stable formatting."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def write_outputs(
    clean_wide: pd.DataFrame,
    single_choice_wide: pd.DataFrame,
    multiple_choice_long: pd.DataFrame,
    catalog: pd.DataFrame,
    dropped: pd.DataFrame,
    duplicate_merges: pd.DataFrame,
    value_audit: pd.DataFrame,
    app_question_catalog: pd.DataFrame,
    output_excel: Path,
) -> None:
    """Write Excel, CSV and JSON outputs."""
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    output_excel.parent.mkdir(parents=True, exist_ok=True)

    # Remove stale app inputs/results from previous experimental versions.
    # The app must read only precomputed outputs generated by script 02.
    for stale_name in [
        "survey_clean.json",
        "multiple_choice_long.json",
        "question_catalog.json",
        "question_catalog_app.json",
        "question_catalog_full.json",
    ]:
        stale_path = APP_DATA_DIR / stale_name
        if stale_path.exists():
            stale_path.unlink()

    readme = pd.DataFrame(
        {
            "item": [
                "description",
                "rows_clean_wide",
                "columns_clean_wide",
                "rows_multiple_choice_long",
                "note_privacy",
                "next_step",
            ],
            "value": [
                "Dataset cleaned and reorganized from the original questionnaire export.",
                len(clean_wide),
                len(clean_wide.columns),
                len(multiple_choice_long),
                "Personal identifiers and technical metadata were excluded by default.",
                "Run scripts/02_generate_app_outputs.py to calculate app-ready descriptive tables, crosstabs and dashboard metrics.",
            ],
        }
    )

    # Excel workbook with multiple organized sheets.
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        readme.to_excel(writer, index=False, sheet_name="00_README")
        clean_wide.to_excel(writer, index=False, sheet_name="01_clean_wide")
        single_choice_wide.to_excel(writer, index=False, sheet_name="02_single_choice")
        multiple_choice_long.to_excel(writer, index=False, sheet_name="03_multiple_long")
        catalog.to_excel(writer, index=False, sheet_name="04_catalog")
        value_audit.to_excel(writer, index=False, sheet_name="05_value_audit")
        dropped.to_excel(writer, index=False, sheet_name="06_dropped")
        duplicate_merges.to_excel(writer, index=False, sheet_name="07_duplicate_merges")
        app_question_catalog.to_excel(writer, index=False, sheet_name="08_app_catalog")

    # CSV outputs for scripts, R, SPSS or app export.
    clean_wide.to_csv(PROCESSED_DIR / "questionario_clean_wide.csv", index=False, encoding="utf-8-sig")
    single_choice_wide.to_csv(PROCESSED_DIR / "single_choice_wide.csv", index=False, encoding="utf-8-sig")
    multiple_choice_long.to_csv(PROCESSED_DIR / "multiple_choice_long.csv", index=False, encoding="utf-8-sig")
    catalog.to_csv(PROCESSED_DIR / "question_catalog.csv", index=False, encoding="utf-8-sig")
    app_question_catalog.to_csv(PROCESSED_DIR / "question_catalog_app.csv", index=False, encoding="utf-8-sig")
    value_audit.to_csv(PROCESSED_DIR / "value_audit.csv", index=False, encoding="utf-8-sig")

    # Minimal metadata for the app. Final analysis outputs are generated by
    # scripts/02_generate_app_outputs.py. The app only reads precomputed JSON.
    app_payloads = {
        "questions.json": dataframe_records_for_json(app_question_catalog),
        "app_config.json": build_app_config(total_respondents=len(clean_wide)),
    }

    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    for filename, payload in app_payloads.items():
        write_json(APP_DATA_DIR / filename, payload)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Clean and reorganize the questionnaire Excel file.")
    parser.add_argument(
        "--input",
        type=Path,
        default=RAW_DIR / "questionario_original.xlsx",
        help="Path to the original Excel file.",
    )
    parser.add_argument(
        "--sheet",
        type=str,
        default=None,
        help="Optional sheet name. If omitted, the first sheet is used.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=PROCESSED_DIR / "questionario_clean.xlsx",
        help="Path to the organized output workbook.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    print(f"Reading input: {args.input}")
    df_raw = read_input_excel(args.input, sheet_name=args.sheet)
    print(f"Raw shape: {df_raw.shape[0]} rows x {df_raw.shape[1]} columns")

    outputs = build_catalog_and_clean_data(df_raw)
    clean_wide, single_choice_wide, multiple_choice_long, catalog, dropped, duplicate_merges, value_audit = outputs
    app_question_catalog = build_app_question_catalog(catalog)

    write_outputs(
        clean_wide=clean_wide,
        single_choice_wide=single_choice_wide,
        multiple_choice_long=multiple_choice_long,
        catalog=catalog,
        dropped=dropped,
        duplicate_merges=duplicate_merges,
        value_audit=value_audit,
        app_question_catalog=app_question_catalog,
        output_excel=args.output,
    )

    print("Done.")
    print(f"Clean wide shape: {clean_wide.shape[0]} rows x {clean_wide.shape[1]} columns")
    print(f"Multiple-choice long shape: {multiple_choice_long.shape[0]} rows x {multiple_choice_long.shape[1]} columns")
    print(f"App catalog shape: {app_question_catalog.shape[0]} questions x {app_question_catalog.shape[1]} fields")
    print(f"Output workbook: {args.output}")
    print(f"Processed data folder: {PROCESSED_DIR}")
    print(f"App data folder: {APP_DATA_DIR}")


if __name__ == "__main__":
    main()
