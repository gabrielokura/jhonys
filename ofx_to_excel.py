#!/usr/bin/env python
"""Convert OFX bank statement transactions to a two-sheet Excel workbook."""

from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
import zipfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.sax.saxutils import escape

from ofx_to_csv import (
    COLUMNS,
    extract_account_data,
    extract_transactions,
    read_ofx,
    transaction_to_row,
)


CLASSIFIED_COLUMNS = COLUMNS + [
    "status_classificacao",
    "grupo",
    "categoria",
    "subcategoria",
    "tipo_fluxo",
    "regra_aplicada",
    "texto_classificacao",
]

RULE_COLUMNS = [
    "padrao",
    "grupo",
    "categoria",
    "subcategoria",
    "tipo_fluxo",
    "observacao",
]


@dataclass(frozen=True)
class ClassificationRule:
    pattern: str
    normalized_pattern: str
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
    return " ".join(without_accents.casefold().split())


def read_classification_rules(path: Path) -> list[ClassificationRule]:
    if not path.exists():
        return []

    rules: list[ClassificationRule] = []
    with path.open("r", encoding="utf-8-sig", newline="") as rules_file:
        reader = csv.DictReader(rules_file, delimiter=";")
        for row_number, row in enumerate(reader, start=2):
            pattern = (row.get("padrao") or "").strip()
            if not pattern:
                continue

            normalized_pattern = normalize_text(pattern)
            if not normalized_pattern:
                continue

            rules.append(
                ClassificationRule(
                    pattern=pattern,
                    normalized_pattern=normalized_pattern,
                    group=(row.get("grupo") or "").strip(),
                    category=(row.get("categoria") or "").strip(),
                    subcategory=(row.get("subcategoria") or "").strip(),
                    cash_flow_type=(row.get("tipo_fluxo") or "").strip(),
                    note=(row.get("observacao") or f"Linha {row_number}").strip(),
                )
            )

    return rules


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


def classify_row(
    row: dict[str, str],
    rules: list[ClassificationRule],
) -> dict[str, str]:
    search_text = " ".join(
        part for part in [row.get("descricao", ""), row.get("memo", "")] if part
    )
    normalized_search_text = normalize_text(search_text)

    for rule in rules:
        if rule.normalized_pattern in normalized_search_text:
            return {
                **row,
                "status_classificacao": "classificado",
                "grupo": rule.group,
                "categoria": rule.category,
                "subcategoria": rule.subcategory,
                "tipo_fluxo": rule.cash_flow_type or infer_cash_flow_type(row),
                "regra_aplicada": rule.pattern,
                "texto_classificacao": search_text,
            }

    return {
        **row,
        "status_classificacao": "a_classificar",
        "grupo": "",
        "categoria": "",
        "subcategoria": "",
        "tipo_fluxo": infer_cash_flow_type(row),
        "regra_aplicada": "",
        "texto_classificacao": search_text,
    }


def parse_ofx_rows(input_path: Path) -> tuple[list[dict[str, str]], str]:
    text, encoding = read_ofx(input_path)
    account = extract_account_data(text)
    transaction_blocks = extract_transactions(text)
    rows = [transaction_to_row(block, account) for block in transaction_blocks]
    return rows, encoding


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
    if column != "valor":
        return None

    amount = amount_to_decimal(value)
    if amount is None:
        return None

    return format(amount, "f")


def worksheet_xml(rows: list[dict[str, str]], columns: list[str]) -> str:
    max_row = max(len(rows) + 1, 1)
    max_column = max(len(columns), 1)
    dimension = f"A1:{cell_reference(max_row, max_column)}"

    xml_rows = [f'<row r="1">{header_cells(columns)}</row>']
    for row_index, row in enumerate(rows, start=2):
        cells = []
        for column_index, column in enumerate(columns, start=1):
            cells.append(cell_xml(row_index, column_index, column, row.get(column, "")))
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

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


def header_cells(columns: list[str]) -> str:
    cells = []
    for column_index, column in enumerate(columns, start=1):
        cells.append(
            cell_xml(
                row_number=1,
                column_number=column_index,
                column=column,
                value=column,
            )
        )
    return "".join(cells)


def cell_xml(row_number: int, column_number: int, column: str, value: str) -> str:
    reference = cell_reference(row_number, column_number)
    numeric = numeric_value(column, value)
    if numeric is not None:
        return f'<c r="{reference}"><v>{numeric}</v></c>'

    text = escape(sanitize_xml_text(value))
    return f'<c r="{reference}" t="inlineStr"><is><t>{text}</t></is></c>'


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = []
    for index, sheet_name in enumerate(sheet_names, start=1):
        escaped_name = escape(sheet_name)
        sheets.append(
            f'<sheet name="{escaped_name}" sheetId="{index}" r:id="rId{index}"/>'
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


def convert_ofx_to_excel(
    input_path: Path,
    output_path: Path,
    rules_path: Path,
) -> tuple[int, int, int, str]:
    raw_rows, encoding = parse_ofx_rows(input_path)
    rules = read_classification_rules(rules_path)
    classified_rows = [classify_row(row, rules) for row in raw_rows]
    classified_count = sum(
        1
        for row in classified_rows
        if row["status_classificacao"] == "classificado"
    )

    write_xlsx(
        output_path,
        [
            ("extrato_bruto", raw_rows, COLUMNS),
            ("extrato_classificado", classified_rows, CLASSIFIED_COLUMNS),
        ],
    )

    return len(raw_rows), classified_count, len(rules), encoding


def parse_args(argv: list[str]) -> argparse.Namespace:
    default_rules = Path(__file__).with_name("classification_rules.csv")

    parser = argparse.ArgumentParser(
        description="Converte um arquivo OFX em Excel com abas bruto e classificado.",
    )
    parser.add_argument(
        "ofx_file",
        type=Path,
        help="Caminho do arquivo .ofx de entrada.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Caminho do arquivo .xlsx de saida. Padrao: mesmo nome do OFX.",
    )
    parser.add_argument(
        "--rules",
        type=Path,
        default=default_rules,
        help="CSV de regras de classificacao separado por ponto-e-virgula.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    input_path = args.ofx_file
    output_path = args.output or input_path.with_suffix(".xlsx")

    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")

    if not input_path.exists():
        print(f"Erro: arquivo OFX nao encontrado: {input_path}", file=sys.stderr)
        return 1

    if not args.rules.exists():
        print(f"Aviso: arquivo de regras nao encontrado: {args.rules}")

    transaction_count, classified_count, rule_count, encoding = convert_ofx_to_excel(
        input_path=input_path,
        output_path=output_path,
        rules_path=args.rules,
    )

    print(f"Excel gerado: {output_path}")
    print(f"Transacoes exportadas: {transaction_count}")
    print(f"Transacoes classificadas: {classified_count}")
    print(f"Regras carregadas: {rule_count}")
    print(f"Encoding lido: {encoding}")

    if transaction_count == 0:
        print("Aviso: nenhuma transacao <STMTTRN> foi encontrada.", file=sys.stderr)
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
