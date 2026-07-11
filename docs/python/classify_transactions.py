#!/usr/bin/env python
"""Classify extracted bank transactions and write a two-sheet Excel workbook."""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
import unicodedata
import zipfile
from dataclasses import dataclass, replace
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.sax.saxutils import escape


DEFAULT_RULES_FILES = ("classification_rules_completed.csv", "classification_rules.csv")
ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")

CLASSIFICATION_COLUMNS = [
    "status_classificacao",
    "grupo",
    "categoria",
    "subcategoria",
    "tipo_fluxo",
    "regra_aplicada",
    "motivo_classificacao",
    "texto_classificacao",
]

PENDING_COLUMNS = [
    "status_classificacao",
    "motivo_classificacao",
    "memo",
    "descricao",
    "quantidade",
    "valor_total",
    "primeira_data",
    "ultima_data",
    "tipo",
    "regra_aplicada",
    "exemplo_id_transacao",
]

SUMMARY_COLUMNS = [
    "status_classificacao",
    "grupo",
    "categoria",
    "subcategoria",
    "tipo_fluxo",
    "memo",
    "descricao_agrupamento",
    "regra_aplicada",
    "quantidade",
    "valor_total",
]


@dataclass(frozen=True)
class ClassificationRule:
    pattern: str
    normalized_pattern: str
    memo_pattern: str
    normalized_memo_pattern: str
    description_pattern: str
    normalized_description_pattern: str
    requires_description: str
    group: str
    category: str
    subcategory: str
    cash_flow_type: str
    note: str


def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    without_accents = "".join(
        char for char in normalized if not unicodedata.combining(char)
    )
    normalized_words = re.sub(r"[^0-9a-zA-Z]+", " ", without_accents.casefold())
    return " ".join(normalized_words.split())


def normalize_header(value: str) -> str:
    return normalize_text(value).replace(" ", "_")


def row_value(row: dict[str, str], *field_names: str) -> str:
    normalized_row = {
        normalize_header(key): value
        for key, value in row.items()
        if key is not None
    }

    for field_name in field_names:
        value = normalized_row.get(normalize_header(field_name))
        if value is not None:
            return value

    return ""


def read_text_with_fallback(path: Path) -> tuple[str, str]:
    data = path.read_bytes()

    for encoding in ENCODINGS:
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue

    return data.decode("latin-1", errors="replace"), "latin-1"


def read_csv_rows(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    text, _ = read_text_with_fallback(path)
    reader = csv.DictReader(io.StringIO(text), delimiter=";")
    if reader.fieldnames is None:
        return [], []

    rows = [dict(row) for row in reader]
    return rows, list(reader.fieldnames)


def has_rule_classification(rule: ClassificationRule) -> bool:
    return any(
        [
            rule.group,
            rule.category,
            rule.subcategory,
            rule.cash_flow_type,
        ]
    )


def classification_key(rule: ClassificationRule) -> tuple[str, str, str, str]:
    return (
        rule.group,
        rule.category,
        rule.subcategory,
        rule.cash_flow_type,
    )


def apply_group_as_category(rule: ClassificationRule) -> ClassificationRule:
    if rule.group and not rule.category:
        return replace(rule, category=rule.group)
    return rule


def infer_debit_classification(rule: ClassificationRule) -> ClassificationRule:
    if has_rule_classification(rule):
        return rule

    normalized_memo = rule.normalized_memo_pattern
    if not normalized_memo.startswith("deb"):
        return rule

    convenio_terms = {
        "conv",
        "tribut",
        "telecomunicacoes",
        "en eletrica",
        "saneamento",
        "orgaos gov",
        "seguros",
        "gps",
        "simples",
        "difal",
    }
    if any(term in normalized_memo for term in convenio_terms):
        return replace(rule, group="Despesas", category="Convênios")

    fornecedor_terms = {
        "tit",
        "titulo",
        "boleto",
        "cobranca",
        "pagamento",
    }
    if any(term in normalized_memo for term in fornecedor_terms):
        return replace(rule, group="Despesas", category="Fornecedores")

    return rule


def infer_cheque_classification(rule: ClassificationRule) -> ClassificationRule:
    if has_rule_classification(rule):
        return rule

    searchable_text = " ".join(
        [
            rule.normalized_pattern,
            rule.normalized_memo_pattern,
            rule.normalized_description_pattern,
        ]
    )
    if "cheque" in searchable_text or "talao" in searchable_text:
        return replace(rule, group="Cheque", category="Cheque")

    return rule


def inherit_unique_memo_classification(
    rules: list[ClassificationRule],
) -> list[ClassificationRule]:
    classifications_by_memo: dict[str, set[tuple[str, str, str, str]]] = {}

    for rule in rules:
        if not rule.normalized_memo_pattern or not has_rule_classification(rule):
            continue
        classifications_by_memo.setdefault(rule.normalized_memo_pattern, set()).add(
            classification_key(rule)
        )

    inherited_rules = []
    for rule in rules:
        if has_rule_classification(rule) or not rule.normalized_memo_pattern:
            inherited_rules.append(rule)
            continue

        memo_classifications = classifications_by_memo.get(rule.normalized_memo_pattern)
        if not memo_classifications or len(memo_classifications) != 1:
            inherited_rules.append(rule)
            continue

        group, category, subcategory, cash_flow_type = next(iter(memo_classifications))
        inherited_rules.append(
            replace(
                rule,
                group=group,
                category=category,
                subcategory=subcategory,
                cash_flow_type=cash_flow_type,
            )
        )

    return inherited_rules


def enrich_classification_rules(
    rules: list[ClassificationRule],
) -> list[ClassificationRule]:
    enriched = [apply_group_as_category(rule) for rule in rules]
    enriched = [infer_debit_classification(rule) for rule in enriched]
    enriched = [infer_cheque_classification(rule) for rule in enriched]
    enriched = inherit_unique_memo_classification(enriched)
    return [apply_group_as_category(rule) for rule in enriched]


def read_classification_rules(path: Path) -> list[ClassificationRule]:
    if not path.exists():
        return []

    rules: list[ClassificationRule] = []
    text, _ = read_text_with_fallback(path)
    reader = csv.DictReader(io.StringIO(text), delimiter=";")
    for row_number, row in enumerate(reader, start=2):
        pattern = row_value(row, "padrao").strip()
        memo_pattern = row_value(row, "memo").strip()
        description_pattern = row_value(row, "descricao").strip()
        requires_description = row_value(
            row,
            "precisa descricao",
        ).strip()

        normalized_pattern = normalize_text(pattern)
        normalized_memo_pattern = normalize_text(memo_pattern)
        normalized_description_pattern = normalize_text(description_pattern)

        if not any(
            [
                normalized_pattern,
                normalized_memo_pattern,
                normalized_description_pattern,
            ]
        ):
            continue

        rules.append(
            ClassificationRule(
                pattern=pattern,
                normalized_pattern=normalized_pattern,
                memo_pattern=memo_pattern,
                normalized_memo_pattern=normalized_memo_pattern,
                description_pattern=description_pattern,
                normalized_description_pattern=normalized_description_pattern,
                requires_description=requires_description,
                group=row_value(row, "grupo").strip(),
                category=row_value(row, "categoria").strip(),
                subcategory=row_value(row, "subcategoria").strip(),
                cash_flow_type=row_value(row, "tipo_fluxo").strip(),
                note=row_value(row, "observacao").strip() or f"Linha {row_number}",
            )
        )

    return enrich_classification_rules(rules)


def amount_to_decimal(value: str) -> Decimal | None:
    normalized = (value or "").strip().replace(".", "").replace(",", ".")
    if not normalized:
        return None

    try:
        return Decimal(normalized)
    except InvalidOperation:
        return None


def infer_cash_flow_type(row: dict[str, str]) -> str:
    amount = amount_to_decimal(row.get("valor", ""))
    if amount is None:
        return ""
    if amount < 0:
        return "Saida"
    if amount > 0:
        return "Entrada"
    return "Neutro"


def fallback_review_classification(row: dict[str, str]) -> tuple[str, str, str]:
    cash_flow_type = infer_cash_flow_type(row)
    if cash_flow_type == "Entrada":
        return "Outras entradas", "Avaliar", cash_flow_type
    if cash_flow_type == "Saida":
        return "Outros gastos", "Avaliar", cash_flow_type
    return "Outros", "Avaliar", cash_flow_type


def classification_text(row: dict[str, str]) -> str:
    parts = [
        row.get("descricao", ""),
        row.get("memo", ""),
    ]
    return " ".join(part for part in parts if part)


def pix_priority_classification(
    row: dict[str, str],
    normalized_memo: str,
    search_text: str,
) -> dict[str, str] | None:
    if "pix" not in normalized_memo:
        return None

    amount = amount_to_decimal(row.get("valor", ""))
    transaction_type = normalize_text(row.get("tipo", ""))
    incoming_terms = ("recebida", "recebido", "credito", "cred")
    outgoing_terms = ("realizada", "emitido", "debito", "deb")

    is_incoming = (
        any(term in normalized_memo for term in incoming_terms)
        or transaction_type == "credit"
        or (amount is not None and amount > 0)
    )
    is_outgoing = (
        any(term in normalized_memo for term in outgoing_terms)
        or transaction_type == "debit"
        or (amount is not None and amount < 0)
    )

    if not is_incoming and not is_outgoing:
        return None

    category = "Pix entrada" if is_incoming else "Pix saída"
    cash_flow_type = "Entrada" if is_incoming else "Saida"

    return {
        **row,
        "status_classificacao": "classificado",
        "grupo": "PIX",
        "categoria": category,
        "subcategoria": "",
        "tipo_fluxo": cash_flow_type,
        "regra_aplicada": f"prioridade_pix_por_memo={row.get('memo', '')}",
        "motivo_classificacao": "prioridade aplicada para operacao PIX",
        "texto_classificacao": search_text,
    }


def rule_requires_description(rule: ClassificationRule) -> bool:
    return normalize_text(rule.requires_description) in {"sim", "s", "yes", "true", "1"}


def rule_uses_description(rule: ClassificationRule) -> bool:
    if not rule.normalized_description_pattern:
        return False
    if not rule.normalized_memo_pattern:
        return True
    return rule_requires_description(rule)


def rule_label(rule: ClassificationRule) -> str:
    parts = []
    if rule.pattern:
        parts.append(f"padrao={rule.pattern}")
    if rule.memo_pattern:
        parts.append(f"memo={rule.memo_pattern}")
    if rule_uses_description(rule):
        parts.append(f"descricao={rule.description_pattern}")
    return " | ".join(parts)


def rule_match_reason(rule: ClassificationRule) -> str:
    if rule.pattern:
        return "regra aplicada por texto combinado"
    if rule.normalized_memo_pattern and rule_uses_description(rule):
        return "regra aplicada por memo e descricao"
    if rule.normalized_memo_pattern:
        return "regra aplicada por memo"
    if rule_uses_description(rule):
        return "regra aplicada por descricao"
    return "regra aplicada"


def rule_has_classification(rule: ClassificationRule) -> bool:
    return any(
        [
            rule.group,
            rule.category,
            rule.subcategory,
            rule.cash_flow_type,
        ]
    )


def rule_matches(
    rule: ClassificationRule,
    normalized_search_text: str,
    normalized_memo: str,
    normalized_description: str,
) -> bool:
    if rule.normalized_pattern:
        return rule.normalized_pattern in normalized_search_text

    if rule_requires_description(rule) and not rule.normalized_description_pattern:
        return False

    if rule.normalized_memo_pattern:
        if rule.normalized_memo_pattern not in normalized_memo:
            return False

    if rule_uses_description(rule):
        if rule.normalized_description_pattern not in normalized_description:
            return False

    return bool(rule.normalized_memo_pattern or rule_uses_description(rule))


def classify_row(
    row: dict[str, str],
    rules: list[ClassificationRule],
) -> dict[str, str]:
    search_text = classification_text(row)
    normalized_search_text = normalize_text(search_text)
    normalized_memo = normalize_text(row.get("memo", ""))
    normalized_description = normalize_text(row.get("descricao", ""))
    pix_classification = pix_priority_classification(row, normalized_memo, search_text)
    if pix_classification is not None:
        return pix_classification

    for rule in rules:
        if rule_matches(
            rule,
            normalized_search_text,
            normalized_memo,
            normalized_description,
        ):
            status = "classificado"
            reason = rule_match_reason(rule)
            if not rule_has_classification(rule):
                status = "regra_sem_classificacao"
                reason = "regra encontrada, mas sem grupo/categoria definidos"

            return {
                **row,
                "status_classificacao": status,
                "grupo": rule.group,
                "categoria": rule.category,
                "subcategoria": rule.subcategory,
                "tipo_fluxo": rule.cash_flow_type or infer_cash_flow_type(row),
                "regra_aplicada": rule_label(rule),
                "motivo_classificacao": reason,
                "texto_classificacao": search_text,
            }

    fallback_group, fallback_category, fallback_cash_flow_type = (
        fallback_review_classification(row)
    )
    return {
        **row,
        "status_classificacao": "avaliar",
        "grupo": fallback_group,
        "categoria": fallback_category,
        "subcategoria": "",
        "tipo_fluxo": fallback_cash_flow_type,
        "regra_aplicada": "fallback_sem_regra",
        "motivo_classificacao": "nenhuma regra encontrada; revisar e incorporar as regras",
        "texto_classificacao": search_text,
    }


def classified_columns(input_columns: list[str]) -> list[str]:
    columns = list(input_columns)
    for column in CLASSIFICATION_COLUMNS:
        if column not in columns:
            columns.append(column)
    return columns


def column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def cell_reference(row_number: int, column_number: int) -> str:
    return f"{column_name(column_number)}{row_number}"


def sanitize_xml_text(value: object) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


def numeric_value(column: str, value: str) -> str | None:
    if column not in {"valor", "valor_total"}:
        return None

    amount = amount_to_decimal(value)
    if amount is None:
        return None

    return format(amount, "f")


def cell_xml(row_number: int, column_number: int, column: str, value: str) -> str:
    reference = cell_reference(row_number, column_number)
    numeric = numeric_value(column, value)
    if numeric is not None:
        return f'<c r="{reference}"><v>{numeric}</v></c>'

    text = escape(sanitize_xml_text(value))
    return f'<c r="{reference}" t="inlineStr"><is><t>{text}</t></is></c>'


def row_xml(row_number: int, row: dict[str, str], columns: list[str]) -> str:
    cells = []
    for column_index, column in enumerate(columns, start=1):
        cells.append(cell_xml(row_number, column_index, column, row.get(column, "")))
    return f'<row r="{row_number}">{"".join(cells)}</row>'


def worksheet_xml(rows: list[dict[str, str]], columns: list[str]) -> str:
    max_row = max(len(rows) + 1, 1)
    max_column = max(len(columns), 1)
    dimension = f"A1:{cell_reference(max_row, max_column)}"

    header_row = {column: column for column in columns}
    xml_rows = [row_xml(1, header_row, columns)]
    for row_number, row in enumerate(rows, start=2):
        xml_rows.append(row_xml(row_number, row, columns))

    autofilter = f'<autoFilter ref="A1:{cell_reference(max_row, max_column)}"/>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>'
        f"<sheetData>{''.join(xml_rows)}</sheetData>"
        f"{autofilter}"
        "</worksheet>"
    )


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = []
    for index, sheet_name in enumerate(sheet_names, start=1):
        sheets.append(
            f'<sheet name="{escape(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>'
        )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{''.join(sheets)}</sheets>"
        "</workbook>"
    )


def workbook_relationships(sheet_count: int) -> str:
    relationships = []
    for index in range(1, sheet_count + 1):
        relationships.append(
            '<Relationship '
            f'Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{''.join(relationships)}"
        "</Relationships>"
    )


def root_relationships() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def content_types(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{''.join(overrides)}"
        "</Types>"
    )


def write_xlsx(
    output_path: Path,
    sheets: list[tuple[str, list[dict[str, str]], list[str]]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types(len(sheets)))
        workbook.writestr("_rels/.rels", root_relationships())
        workbook.writestr(
            "xl/_rels/workbook.xml.rels",
            workbook_relationships(len(sheets)),
        )
        workbook.writestr("xl/workbook.xml", workbook_xml([sheet[0] for sheet in sheets]))

        for index, (_, rows, columns) in enumerate(sheets, start=1):
            workbook.writestr(
                f"xl/worksheets/sheet{index}.xml",
                worksheet_xml(rows, columns),
            )


def format_decimal_br(value: Decimal) -> str:
    return f"{value.quantize(Decimal('0.01')):.2f}".replace(".", ",")


def grouped_description(row: dict[str, str]) -> str:
    rule_applied = row.get("regra_aplicada", "")
    if "descricao=" in rule_applied:
        return row.get("descricao", "")

    if row.get("status_classificacao") == "classificado":
        return ""

    if "pix" in normalize_text(row.get("memo", "")):
        return ""

    return row.get("descricao", "")


def summary_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    buckets: dict[tuple[str, ...], dict[str, object]] = {}

    for row in rows:
        key = (
            row.get("status_classificacao", ""),
            row.get("grupo", ""),
            row.get("categoria", ""),
            row.get("subcategoria", ""),
            row.get("tipo_fluxo", ""),
            row.get("memo", ""),
            grouped_description(row),
            row.get("regra_aplicada", ""),
        )
        amount = amount_to_decimal(row.get("valor", "")) or Decimal("0")

        if key not in buckets:
            buckets[key] = {
                "status_classificacao": key[0],
                "grupo": key[1],
                "categoria": key[2],
                "subcategoria": key[3],
                "tipo_fluxo": key[4],
                "memo": key[5],
                "descricao_agrupamento": key[6],
                "regra_aplicada": key[7],
                "quantidade": 0,
                "valor_total": Decimal("0"),
            }

        bucket = buckets[key]
        bucket["quantidade"] = int(bucket["quantidade"]) + 1
        bucket["valor_total"] = bucket["valor_total"] + amount

    output_rows = []
    for bucket in buckets.values():
        output_rows.append(
            {
                **bucket,
                "quantidade": str(bucket["quantidade"]),
                "valor_total": format_decimal_br(bucket["valor_total"]),
            }
        )

    output_rows.sort(
        key=lambda row: (
            row["status_classificacao"],
            row["grupo"],
            row["categoria"],
            row["memo"],
            row["descricao_agrupamento"],
        )
    )
    return output_rows


def pending_review_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    buckets: dict[tuple[str, str, str, str], dict[str, object]] = {}

    for row in rows:
        if row.get("status_classificacao") == "classificado":
            continue

        description = grouped_description(row)
        key = (
            row.get("status_classificacao", ""),
            row.get("memo", ""),
            description,
            row.get("regra_aplicada", ""),
        )
        amount = amount_to_decimal(row.get("valor", "")) or Decimal("0")

        if key not in buckets:
            buckets[key] = {
                "status_classificacao": row.get("status_classificacao", ""),
                "motivo_classificacao": row.get("motivo_classificacao", ""),
                "memo": row.get("memo", ""),
                "descricao": description,
                "quantidade": 0,
                "valor_total": Decimal("0"),
                "primeira_data": row.get("data", ""),
                "ultima_data": row.get("data", ""),
                "tipo": row.get("tipo", ""),
                "regra_aplicada": row.get("regra_aplicada", ""),
                "exemplo_id_transacao": row.get("id_transacao", ""),
            }

        bucket = buckets[key]
        bucket["quantidade"] = int(bucket["quantidade"]) + 1
        bucket["valor_total"] = bucket["valor_total"] + amount
        bucket["ultima_data"] = row.get("data", "")

    review_rows = []
    for bucket in buckets.values():
        review_rows.append(
            {
                **bucket,
                "quantidade": str(bucket["quantidade"]),
                "valor_total": format_decimal_br(bucket["valor_total"]),
            }
        )

    review_rows.sort(
        key=lambda row: (
            row["status_classificacao"],
            row["memo"],
            row["descricao"],
        )
    )
    return review_rows


def next_numbered_path(path: Path, number: int) -> Path:
    return path.with_name(f"{path.stem}_{number}{path.suffix}")


def write_pending_csv(path: Path, rows: list[dict[str, str]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)

    for attempt in range(100):
        output_path = path if attempt == 0 else next_numbered_path(path, attempt)
        try:
            with output_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
                writer = csv.DictWriter(
                    csv_file,
                    fieldnames=PENDING_COLUMNS,
                    delimiter=";",
                    lineterminator="\n",
                )
                writer.writeheader()
                writer.writerows(rows)
            return output_path
        except PermissionError:
            continue

    raise PermissionError(f"Nao foi possivel gravar o CSV de pendencias: {path}")


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_classificado.xlsx")


def default_pending_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}_operacoes_a_classificar.csv")


def classify_transactions_file(
    input_path: Path,
    output_path: Path,
    rules_path: Path,
    pending_output_path: Path,
) -> tuple[int, int, int, int, Path]:
    raw_rows, input_columns = read_csv_rows(input_path)
    rules = read_classification_rules(rules_path)
    classified_rows = [classify_row(row, rules) for row in raw_rows]
    classified_count = sum(
        1
        for row in classified_rows
        if row["status_classificacao"] == "classificado"
    )
    summary = summary_rows(classified_rows)
    review_rows = pending_review_rows(classified_rows)

    write_xlsx(
        output_path,
        [
            ("extrato_bruto", raw_rows, input_columns),
            ("extrato_classificado", classified_rows, classified_columns(input_columns)),
            ("resumo_classificacao", summary, SUMMARY_COLUMNS),
            ("operacoes_a_classificar", review_rows, PENDING_COLUMNS),
        ],
    )
    actual_pending_output_path = write_pending_csv(pending_output_path, review_rows)

    return (
        len(raw_rows),
        classified_count,
        len(review_rows),
        len(rules),
        actual_pending_output_path,
    )


def parse_args(argv: list[str]) -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent
    default_rules = next(
        (
            script_dir / file_name
            for file_name in DEFAULT_RULES_FILES
            if (script_dir / file_name).exists()
        ),
        script_dir / DEFAULT_RULES_FILES[-1],
    )
    parser = argparse.ArgumentParser(
        description="Classifica um CSV de extrato bancario e gera Excel com duas abas.",
    )
    parser.add_argument(
        "csv_file",
        type=Path,
        help="Caminho do CSV gerado pelo ofx_to_csv.py.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Caminho do Excel .xlsx de saida. Padrao: *_classificado.xlsx.",
    )
    parser.add_argument(
        "--rules",
        type=Path,
        default=default_rules,
        help="CSV de regras separado por ponto-e-virgula.",
    )
    parser.add_argument(
        "--pending-output",
        type=Path,
        help="CSV resumido com operacoes pendentes de classificacao.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    input_path = args.csv_file
    output_path = args.output or default_output_path(input_path)
    pending_output_path = args.pending_output or default_pending_path(output_path)

    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")

    if not input_path.exists():
        print(f"Erro: CSV de entrada nao encontrado: {input_path}", file=sys.stderr)
        return 1

    if not args.rules.exists():
        print(f"Aviso: arquivo de regras nao encontrado: {args.rules}")

    (
        transaction_count,
        classified_count,
        pending_count,
        rule_count,
        actual_pending_output_path,
    ) = classify_transactions_file(
        input_path=input_path,
        output_path=output_path,
        rules_path=args.rules,
        pending_output_path=pending_output_path,
    )

    print(f"Excel gerado: {output_path}")
    print(f"CSV de pendencias: {actual_pending_output_path}")
    print(f"Transacoes lidas: {transaction_count}")
    print(f"Transacoes classificadas: {classified_count}")
    print(f"Operacoes pendentes para revisar: {pending_count}")
    print(f"Regras carregadas: {rule_count}")

    if transaction_count == 0:
        print("Aviso: nenhuma transacao foi encontrada no CSV.", file=sys.stderr)
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
