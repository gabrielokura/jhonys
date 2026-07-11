#!/usr/bin/env python
"""Build an HTML report with monthly totals by group and category."""

from __future__ import annotations

import argparse
import html
import re
import sys
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

from ofx_to_csv import extract_tag, extract_transactions, read_ofx


WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

DEFAULT_SHEET_NAME = "extrato_classificado"
DETAIL_ROWS_PER_PAGE = 20


@dataclass
class MonthlyBucket:
    quantity: int = 0
    total: Decimal = Decimal("0")
    inflows: Decimal = Decimal("0")
    outflows: Decimal = Decimal("0")


@dataclass
class BalanceAnalysis:
    status: str
    start_date: str
    end_date: str
    balance_date: str
    opening_balance: Decimal | None
    final_balance: Decimal | None
    movement_total: Decimal
    ofx_movement_total: Decimal
    extraction_difference: Decimal
    reconciliation_difference: Decimal | None
    transaction_count: int
    ofx_transaction_count: int
    note: str


def xml_name(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def read_shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []

    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall(xml_name(WORKBOOK_NS, "si")):
        text_parts = [
            text_node.text or ""
            for text_node in item.iter(xml_name(WORKBOOK_NS, "t"))
        ]
        strings.append("".join(text_parts))
    return strings


def workbook_relationships(workbook: zipfile.ZipFile) -> dict[str, str]:
    root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    relationships = {}
    for relationship in root.findall(xml_name(PACKAGE_REL_NS, "Relationship")):
        relationships[relationship.attrib["Id"]] = relationship.attrib["Target"]
    return relationships


def sheet_path_by_name(workbook: zipfile.ZipFile, sheet_name: str) -> str:
    root = ET.fromstring(workbook.read("xl/workbook.xml"))
    relationships = workbook_relationships(workbook)

    for sheet in root.findall(f".//{xml_name(WORKBOOK_NS, 'sheet')}"):
        if sheet.attrib.get("name") != sheet_name:
            continue

        relation_id = sheet.attrib[xml_name(REL_NS, "id")]
        target = relationships[relation_id].lstrip("/")
        if target.startswith("xl/"):
            return target
        return f"xl/{target}"

    available = [
        sheet.attrib.get("name", "")
        for sheet in root.findall(f".//{xml_name(WORKBOOK_NS, 'sheet')}")
    ]
    raise ValueError(
        f"Aba '{sheet_name}' nao encontrada. Abas disponiveis: {', '.join(available)}"
    )


def column_index_from_reference(cell_reference: str) -> int:
    column_letters = re.match(r"[A-Z]+", cell_reference.upper())
    if not column_letters:
        return 0

    index = 0
    for letter in column_letters.group(0):
        index = index * 26 + (ord(letter) - ord("A") + 1)
    return index - 1


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        text_parts = [
            text_node.text or ""
            for text_node in cell.iter(xml_name(WORKBOOK_NS, "t"))
        ]
        return "".join(text_parts)

    value_node = cell.find(xml_name(WORKBOOK_NS, "v"))
    if value_node is None or value_node.text is None:
        return ""

    value = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (IndexError, ValueError):
            return value

    return value


def read_sheet_rows(workbook_path: Path, sheet_name: str) -> list[dict[str, str]]:
    with zipfile.ZipFile(workbook_path, "r") as workbook:
        sheet_path = sheet_path_by_name(workbook, sheet_name)
        shared_strings = read_shared_strings(workbook)
        root = ET.fromstring(workbook.read(sheet_path))

        table_rows: list[list[str]] = []
        for row in root.findall(f".//{xml_name(WORKBOOK_NS, 'row')}"):
            values: list[str] = []
            for cell in row.findall(xml_name(WORKBOOK_NS, "c")):
                reference = cell.attrib.get("r", "")
                column_index = column_index_from_reference(reference)
                while len(values) <= column_index:
                    values.append("")
                values[column_index] = cell_text(cell, shared_strings)
            table_rows.append(values)

    if not table_rows:
        return []

    headers = [header.strip() for header in table_rows[0]]
    rows = []
    for raw_row in table_rows[1:]:
        row = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            row[header] = raw_row[index] if index < len(raw_row) else ""
        rows.append(row)
    return rows


def parse_amount(value: str) -> Decimal:
    normalized = (value or "").strip()
    if not normalized:
        return Decimal("0")

    if "," in normalized:
        normalized = normalized.replace(".", "").replace(",", ".")

    try:
        return Decimal(normalized)
    except InvalidOperation:
        return Decimal("0")


def parse_ofx_amount(value: str) -> Decimal:
    normalized = (value or "").strip().replace(" ", "")
    if "," in normalized and "." not in normalized:
        normalized = normalized.replace(",", ".")
    try:
        return Decimal(normalized)
    except InvalidOperation:
        return Decimal("0")


def parse_optional_ofx_amount(value: str) -> Decimal | None:
    if not (value or "").strip():
        return None
    return parse_ofx_amount(value)


def format_ofx_display_date(value: str) -> str:
    match = re.match(r"\s*(\d{4})(\d{2})(\d{2})", value or "")
    if match:
        year, month, day = match.groups()
        return f"{day}/{month}/{year}"
    return value or "-"


def first_block(text: str, tag: str) -> str:
    match = re.search(
        rf"<\s*{re.escape(tag)}\b[^>]*>(.*?)(?:<\s*/\s*{re.escape(tag)}\s*>|$)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    return match.group(1) if match else ""


def analyze_balance(ofx_path: Path | None, rows: list[dict[str, str]]) -> BalanceAnalysis | None:
    if ofx_path is None:
        return None
    if not ofx_path.exists():
        return BalanceAnalysis(
            status="sem_ofx",
            start_date="-",
            end_date="-",
            balance_date="-",
            opening_balance=None,
            final_balance=None,
            movement_total=sum((parse_amount(row.get("valor", "")) for row in rows), Decimal("0")),
            ofx_movement_total=Decimal("0"),
            extraction_difference=Decimal("0"),
            reconciliation_difference=None,
            transaction_count=len(rows),
            ofx_transaction_count=0,
            note="Arquivo OFX nao encontrado para conciliacao.",
        )

    text, _ = read_ofx(ofx_path)
    transaction_blocks = extract_transactions(text)
    ledger_balance_block = first_block(text, "LEDGERBAL")
    final_balance = parse_optional_ofx_amount(
        extract_tag(ledger_balance_block or text, "BALAMT")
    )
    movement_total = sum((parse_amount(row.get("valor", "")) for row in rows), Decimal("0"))
    ofx_movement_total = sum(
        (parse_ofx_amount(extract_tag(block, "TRNAMT")) for block in transaction_blocks),
        Decimal("0"),
    )
    extraction_difference = ofx_movement_total - movement_total

    opening_balance: Decimal | None = None
    reconciliation_difference: Decimal | None = None
    if final_balance is not None:
        opening_balance = final_balance - movement_total
        reconciliation_difference = final_balance - (opening_balance + movement_total)

    status = "ok"
    notes = []
    if len(transaction_blocks) != len(rows):
        status = "atencao"
        notes.append(
            f"Quantidade de transacoes difere: OFX={len(transaction_blocks)}, planilha={len(rows)}."
        )
    if abs(extraction_difference) > Decimal("0.01"):
        status = "atencao"
        notes.append(
            f"Soma das movimentacoes difere do OFX em {format_brl(extraction_difference)}."
        )
    if final_balance is None:
        status = "sem_saldo_final"
        notes.append("OFX nao trouxe saldo final em LEDGERBAL/BALAMT.")
    elif reconciliation_difference is not None and abs(reconciliation_difference) > Decimal("0.01"):
        status = "atencao"
        notes.append(
            f"Conciliacao difere em {format_brl(reconciliation_difference)}."
        )

    if not notes and opening_balance is not None:
        notes.append(
            "Saldo inicial inferido pelo saldo final menos todas as movimentacoes; conciliacao bate."
        )

    return BalanceAnalysis(
        status=status,
        start_date=format_ofx_display_date(extract_tag(text, "DTSTART")),
        end_date=format_ofx_display_date(extract_tag(text, "DTEND")),
        balance_date=format_ofx_display_date(extract_tag(ledger_balance_block or text, "DTASOF")),
        opening_balance=opening_balance,
        final_balance=final_balance,
        movement_total=movement_total,
        ofx_movement_total=ofx_movement_total,
        extraction_difference=extraction_difference,
        reconciliation_difference=reconciliation_difference,
        transaction_count=len(rows),
        ofx_transaction_count=len(transaction_blocks),
        note=" ".join(notes),
    )


def month_key(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return "Sem data"

    iso_match = re.match(r"(\d{4})-(\d{2})", value)
    if iso_match:
        return f"{iso_match.group(1)}-{iso_match.group(2)}"

    br_match = re.match(r"\d{2}/(\d{2})/(\d{4})", value)
    if br_match:
        return f"{br_match.group(2)}-{br_match.group(1)}"

    compact_match = re.match(r"(\d{4})(\d{2})", value)
    if compact_match:
        return f"{compact_match.group(1)}-{compact_match.group(2)}"

    return value[:7] if len(value) >= 7 else value


def format_month(value: str) -> str:
    if re.match(r"\d{4}-\d{2}$", value):
        year, month = value.split("-")
        return f"{month}/{year}"
    return value


def format_brl(value: Decimal) -> str:
    sign = "-" if value < 0 else ""
    absolute = abs(value).quantize(Decimal("0.01"))
    integer_part, decimal_part = f"{absolute:.2f}".split(".")

    groups = []
    while integer_part:
        groups.append(integer_part[-3:])
        integer_part = integer_part[:-3]

    return f"{sign}R$ {'.'.join(reversed(groups))},{decimal_part}"


def group_category_from_row(row: dict[str, str]) -> tuple[str, str]:
    group = (row.get("grupo") or "Sem grupo").strip() or "Sem grupo"
    category = (row.get("categoria") or group).strip() or group
    return group, category


def grouped_monthly_totals(
    rows: list[dict[str, str]],
) -> dict[tuple[str, str], dict[str, MonthlyBucket]]:
    grouped: dict[tuple[str, str], dict[str, MonthlyBucket]] = defaultdict(
        lambda: defaultdict(MonthlyBucket)
    )

    for row in rows:
        group, category = group_category_from_row(row)
        month = month_key(row.get("data", ""))
        amount = parse_amount(row.get("valor", ""))
        bucket = grouped[(group, category)][month]

        bucket.quantity += 1
        bucket.total += amount
        if amount >= 0:
            bucket.inflows += amount
        else:
            bucket.outflows += amount

    return grouped


def grouped_detail_rows(
    rows: list[dict[str, str]],
) -> dict[tuple[str, str], list[dict[str, str]]]:
    grouped: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        grouped[group_category_from_row(row)].append(row)
    return grouped


def safe_anchor(value: str) -> str:
    anchor = re.sub(r"[^0-9a-zA-Z]+", "-", value).strip("-").lower()
    return anchor or "secao"


def render_overview(grouped: dict[tuple[str, str], dict[str, MonthlyBucket]]) -> str:
    rows = []
    for (group, category), months in sorted(grouped.items()):
        total = sum((bucket.total for bucket in months.values()), Decimal("0"))
        quantity = sum(bucket.quantity for bucket in months.values())
        label = f"{group} / {category}"
        anchor = safe_anchor(label)
        rows.append(
            "<tr>"
            f'<td><a href="#{anchor}">{html.escape(group)}</a></td>'
            f"<td>{html.escape(category)}</td>"
            f'<td class="number">{quantity}</td>'
            f'<td class="money {amount_class(total)}">{format_brl(total)}</td>'
            "</tr>"
        )

    return (
        '<section class="summary">'
        "<h2>Resumo Geral</h2>"
        "<table>"
        "<thead><tr><th>Grupo</th><th>Categoria</th><th>Qtde.</th><th>Total</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
        "</section>"
    )


def amount_class(value: Decimal) -> str:
    if value < 0:
        return "negative"
    if value > 0:
        return "positive"
    return "neutral"


def format_optional_brl(value: Decimal | None) -> str:
    if value is None:
        return "-"
    return format_brl(value)


def render_balance_analysis(balance: BalanceAnalysis | None) -> str:
    if balance is None:
        return ""

    status_label = {
        "ok": "Batendo",
        "atencao": "Atenção",
        "sem_saldo_final": "Sem saldo final",
        "sem_ofx": "Sem OFX",
    }.get(balance.status, balance.status)

    reconciliation = (
        format_optional_brl(balance.reconciliation_difference)
        if balance.reconciliation_difference is not None
        else "-"
    )

    return (
        f'<section class="balance-section {html.escape(balance.status)}">'
        "<h2>Conferência de Saldo</h2>"
        '<div class="balance-grid">'
        f"<div><span>Status</span><strong>{html.escape(status_label)}</strong></div>"
        f"<div><span>Período</span><strong>{html.escape(balance.start_date)} a {html.escape(balance.end_date)}</strong></div>"
        f"<div><span>Data do saldo</span><strong>{html.escape(balance.balance_date)}</strong></div>"
        f"<div><span>Saldo inicial inferido</span><strong>{format_optional_brl(balance.opening_balance)}</strong></div>"
        f"<div><span>Movimentações</span><strong class=\"{amount_class(balance.movement_total)}\">{format_brl(balance.movement_total)}</strong></div>"
        f"<div><span>Saldo final OFX</span><strong>{format_optional_brl(balance.final_balance)}</strong></div>"
        f"<div><span>Diferença extração x OFX</span><strong class=\"{amount_class(balance.extraction_difference)}\">{format_brl(balance.extraction_difference)}</strong></div>"
        f"<div><span>Diferença conciliação</span><strong>{reconciliation}</strong></div>"
        f"<div><span>Transações</span><strong>{balance.transaction_count} / {balance.ofx_transaction_count}</strong></div>"
        "</div>"
        f"<p>{html.escape(balance.note)}</p>"
        "</section>"
    )


def review_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    output = []
    for row in rows:
        status = row.get("status_classificacao", "")
        group = row.get("grupo", "")
        category = row.get("categoria", "")
        if (
            status != "classificado"
            or group.startswith("Outros")
            or group.startswith("Outras")
            or category == "Avaliar"
        ):
            output.append(row)
    return output


def render_review_section(rows: list[dict[str, str]]) -> str:
    rows_to_review = review_rows(rows)
    if not rows_to_review:
        return (
            '<section class="review-section ok">'
            "<h2>Movimentações Para Avaliação</h2>"
            "<p>Nenhuma movimentação nova ou sem regra foi encontrada.</p>"
            "</section>"
        )

    total = sum((parse_amount(row.get("valor", "")) for row in rows_to_review), Decimal("0"))
    inflows = sum(
        (
            amount
            for amount in (parse_amount(row.get("valor", "")) for row in rows_to_review)
            if amount > 0
        ),
        Decimal("0"),
    )
    outflows = sum(
        (
            amount
            for amount in (parse_amount(row.get("valor", "")) for row in rows_to_review)
            if amount < 0
        ),
        Decimal("0"),
    )
    columns = [
        "data",
        "tipo",
        "valor",
        "descricao",
        "memo",
        "id_transacao",
        "grupo",
        "categoria",
        "motivo_classificacao",
    ]
    return (
        '<section class="review-section warning">'
        "<h2>Movimentações Para Avaliação</h2>"
        f"<p>{len(rows_to_review)} linha(s) precisam ser avaliadas para futura incorporação nas regras.</p>"
        '<div class="review-metrics">'
        f"<div><span>Entradas</span><strong class=\"positive\">{format_brl(inflows)}</strong></div>"
        f"<div><span>Saídas</span><strong class=\"negative\">{format_brl(outflows)}</strong></div>"
        f"<div><span>Saldo</span><strong class=\"{amount_class(total)}\">{format_brl(total)}</strong></div>"
        "</div>"
        f"{render_detail_table(rows_to_review, columns)}"
        "</section>"
    )


def render_detail_table(rows: list[dict[str, str]], columns: list[str]) -> str:
    header_cells = "".join(
        f"<th>{html.escape(column)}</th>"
        for column in columns
    )
    body_rows = []
    for row in rows:
        cells = "".join(
            f"<td>{html.escape(str(row.get(column, '')))}</td>"
            for column in columns
        )
        body_rows.append(f"<tr>{cells}</tr>")

    pagination = ""
    if len(rows) > DETAIL_ROWS_PER_PAGE:
        pages = (len(rows) + DETAIL_ROWS_PER_PAGE - 1) // DETAIL_ROWS_PER_PAGE
        first_page_end = min(DETAIL_ROWS_PER_PAGE, len(rows))
        pagination = (
            '<div class="table-pagination">'
            '<button type="button" class="pagination-prev" aria-label="Página anterior">Anterior</button>'
            f'<span class="pagination-status" aria-live="polite">Mostrando 1-{first_page_end} de {len(rows)} · Página 1 de {pages}</span>'
            '<button type="button" class="pagination-next" aria-label="Próxima página">Próxima</button>'
            "</div>"
        )

    return (
        f'<div class="detail-table-shell" data-page-size="{DETAIL_ROWS_PER_PAGE}">'
        '<div class="detail-table-wrapper">'
        '<table class="detail-table">'
        f"<thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table>"
        "</div>"
        f"{pagination}"
        "</div>"
    )


def render_group_section(
    group: str,
    category: str,
    months: dict[str, MonthlyBucket],
    detail_rows: list[dict[str, str]],
    detail_columns: list[str],
) -> str:
    label = f"{group} / {category}"
    anchor = safe_anchor(label)
    detail_id = f"details-{anchor}"
    section_rows = []
    total_quantity = 0
    total_inflows = Decimal("0")
    total_outflows = Decimal("0")
    total_balance = Decimal("0")

    for month in sorted(months):
        bucket = months[month]
        total_quantity += bucket.quantity
        total_inflows += bucket.inflows
        total_outflows += bucket.outflows
        total_balance += bucket.total
        section_rows.append(
            "<tr>"
            f"<td>{html.escape(format_month(month))}</td>"
            f'<td class="number">{bucket.quantity}</td>'
            f'<td class="money positive">{format_brl(bucket.inflows)}</td>'
            f'<td class="money negative">{format_brl(bucket.outflows)}</td>'
            f'<td class="money {amount_class(bucket.total)}">{format_brl(bucket.total)}</td>'
            "</tr>"
        )

    section_rows.append(
        '<tr class="total-row">'
        "<td>Total</td>"
        f'<td class="number">{total_quantity}</td>'
        f'<td class="money positive">{format_brl(total_inflows)}</td>'
        f'<td class="money negative">{format_brl(total_outflows)}</td>'
        f'<td class="money {amount_class(total_balance)}">{format_brl(total_balance)}</td>'
        "</tr>"
    )

    return (
        f'<section class="group-section" id="{anchor}">'
        '<div class="section-heading">'
        "<div>"
        f"<h2>{html.escape(group)}</h2>"
        f"<h3>{html.escape(category)}</h3>"
        "</div>"
        f'<button class="toggle-details" type="button" data-target="{detail_id}" aria-controls="{detail_id}" aria-expanded="false">Ver mais</button>'
        "</div>"
        "<table>"
        "<thead>"
        "<tr><th>Mes</th><th>Qtde.</th><th>Entradas</th><th>Saidas</th><th>Total mensal</th></tr>"
        "</thead>"
        f"<tbody>{''.join(section_rows)}</tbody>"
        "</table>"
        f'<div class="details-panel" id="{detail_id}" hidden>'
        f"<h4>Linhas da planilha ({len(detail_rows)})</h4>"
        f"{render_detail_table(detail_rows, detail_columns)}"
        "</div>"
        "</section>"
    )


def render_html_report(
    workbook_path: Path,
    rows: list[dict[str, str]],
    grouped: dict[tuple[str, str], dict[str, MonthlyBucket]],
    balance: BalanceAnalysis | None = None,
) -> str:
    total_transactions = len(rows)
    total_amount = sum((parse_amount(row.get("valor", "")) for row in rows), Decimal("0"))
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")
    details_by_group = grouped_detail_rows(rows)
    detail_columns = list(rows[0].keys()) if rows else []

    sections = [
        render_group_section(
            group,
            category,
            months,
            details_by_group.get((group, category), []),
            detail_columns,
        )
        for (group, category), months in sorted(grouped.items())
    ]

    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; connect-src 'none'; img-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'; navigate-to 'none'">
  <title>Relatório Mensal por Grupo e Categoria</title>
  <style>
    :root {{
      --bg: #f6f7f9;
      --panel: #ffffff;
      --text: #1f2933;
      --muted: #5f6b7a;
      --line: #d9dee6;
      --header: #22324a;
      --positive: #0b6b3a;
      --negative: #a73535;
      --neutral: #425466;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: Arial, Helvetica, sans-serif;
      line-height: 1.45;
    }}
    header {{
      background: var(--header);
      color: white;
      padding: 28px 32px;
    }}
    header h1 {{
      margin: 0 0 10px;
      font-size: 28px;
      letter-spacing: 0;
    }}
    header p {{
      margin: 4px 0;
      color: #dbe4ef;
    }}
    main {{
      width: min(1180px, calc(100% - 32px));
      margin: 24px auto 48px;
    }}
    .metric-row {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }}
    .metric {{
      background: var(--panel);
      border: 1px solid var(--line);
      padding: 16px;
      border-radius: 6px;
    }}
    .metric span {{
      display: block;
      color: var(--muted);
      font-size: 13px;
      margin-bottom: 6px;
    }}
    .metric strong {{
      font-size: 22px;
    }}
    .balance-section, .review-section {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 6px;
      margin: 16px 0;
      padding: 18px;
      overflow: hidden;
    }}
    .balance-section h2, .review-section h2 {{
      margin: 0 0 14px;
      padding: 0;
      font-size: 20px;
    }}
    .balance-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
    }}
    .balance-grid div {{
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 12px;
      background: #fbfcfe;
    }}
    .balance-grid span {{
      display: block;
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 5px;
    }}
    .balance-grid strong {{
      font-size: 15px;
    }}
    .balance-section.ok {{
      border-color: #88bf9a;
    }}
    .balance-section.atencao, .balance-section.sem_saldo_final, .review-section.warning {{
      border-color: #e2b86b;
    }}
    .review-section.ok {{
      border-color: #88bf9a;
    }}
    .review-metrics {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
      margin: 12px 0 16px;
    }}
    .review-metrics div {{
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 12px;
      background: #fbfcfe;
    }}
    .review-metrics span {{
      display: block;
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 5px;
    }}
    section {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 6px;
      margin: 16px 0;
      overflow: hidden;
    }}
    .section-heading {{
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 14px;
      padding: 0 18px 14px;
    }}
    section h2, section h3 {{
      margin: 0;
    }}
    section h2 {{
      padding-top: 18px;
      font-size: 20px;
    }}
    section h3 {{
      padding-top: 4px;
      color: var(--muted);
      font-size: 15px;
      font-weight: 600;
    }}
    .toggle-details {{
      margin-top: 18px;
      border: 1px solid #9fb3c8;
      background: #ffffff;
      color: #1f4e79;
      border-radius: 6px;
      padding: 8px 12px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      white-space: nowrap;
    }}
    .toggle-details:hover {{
      background: #eef6ff;
    }}
    .details-panel {{
      border-top: 1px solid var(--line);
      background: #fbfcfe;
      padding: 16px 18px 18px;
    }}
    .details-panel h4 {{
      margin: 0 0 12px;
      font-size: 14px;
      color: var(--muted);
    }}
    .detail-table-shell {{
      display: grid;
      gap: 10px;
    }}
    .detail-table-wrapper {{
      overflow-x: auto;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: white;
    }}
    .detail-table {{
      min-width: 1450px;
      font-size: 12px;
    }}
    .detail-table th {{
      position: sticky;
      top: 0;
      z-index: 1;
    }}
    .detail-table td {{
      white-space: nowrap;
    }}
    .table-pagination {{
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 10px;
      color: var(--muted);
      font-size: 13px;
    }}
    .table-pagination button {{
      border: 1px solid #9fb3c8;
      background: #ffffff;
      color: #1f4e79;
      border-radius: 6px;
      min-height: 44px;
      padding: 9px 12px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
    }}
    .table-pagination button:disabled {{
      cursor: not-allowed;
      opacity: 0.55;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }}
    th, td {{
      padding: 10px 12px;
      border-top: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
    }}
    th {{
      background: #eef2f6;
      color: #26384f;
      font-size: 12px;
      text-transform: uppercase;
    }}
    button:focus-visible, a:focus-visible {{
      outline: 3px solid rgba(36, 95, 158, 0.35);
      outline-offset: 2px;
    }}
    a {{ color: #245f9e; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .number, .money {{ text-align: right; white-space: nowrap; }}
    .positive {{ color: var(--positive); }}
    .negative {{ color: var(--negative); }}
    .neutral {{ color: var(--neutral); }}
    .total-row td {{
      background: #f8fafc;
      font-weight: 700;
    }}
    @media (max-width: 760px) {{
      header {{ padding: 22px 18px; }}
      main {{ width: min(100% - 20px, 1180px); }}
      .metric-row {{ grid-template-columns: 1fr; }}
      .balance-grid {{ grid-template-columns: 1fr; }}
      .review-metrics {{ grid-template-columns: 1fr; }}
      .section-heading {{ flex-direction: column; }}
      .toggle-details {{ margin-top: 0; }}
      .table-pagination {{ justify-content: flex-start; flex-wrap: wrap; }}
      table {{ font-size: 13px; }}
      th, td {{ padding: 9px 8px; }}
    }}
  </style>
</head>
<body>
  <header>
    <h1>Relatorio Mensal por Grupo e Categoria</h1>
    <p>Fonte: {html.escape(str(workbook_path))}</p>
    <p>Gerado em: {html.escape(generated_at)}</p>
  </header>
  <main>
    <div class="metric-row">
      <div class="metric"><span>Transacoes consideradas</span><strong>{total_transactions}</strong></div>
      <div class="metric"><span>Grupos/Categorias</span><strong>{len(grouped)}</strong></div>
      <div class="metric"><span>Total geral</span><strong class="{amount_class(total_amount)}">{format_brl(total_amount)}</strong></div>
    </div>
    {render_balance_analysis(balance)}
    {render_review_section(rows)}
    {render_overview(grouped)}
    {''.join(sections)}
  </main>
  <script>
    document.querySelectorAll('.toggle-details').forEach((button) => {{
      button.addEventListener('click', () => {{
        const target = document.getElementById(button.dataset.target);
        if (!target) return;
        const isHidden = target.hasAttribute('hidden');
        if (isHidden) {{
          target.removeAttribute('hidden');
          button.textContent = 'Ver menos';
          button.setAttribute('aria-expanded', 'true');
        }} else {{
          target.setAttribute('hidden', '');
          button.textContent = 'Ver mais';
          button.setAttribute('aria-expanded', 'false');
        }}
      }});
    }});
    document.querySelectorAll('.detail-table-shell').forEach((shell) => {{
      const rows = Array.from(shell.querySelectorAll('tbody tr'));
      const pagination = shell.querySelector('.table-pagination');
      if (!pagination || rows.length <= Number(shell.dataset.pageSize || 20)) return;

      const pageSize = Number(shell.dataset.pageSize || 20);
      const totalPages = Math.ceil(rows.length / pageSize);
      const prev = pagination.querySelector('.pagination-prev');
      const next = pagination.querySelector('.pagination-next');
      const status = pagination.querySelector('.pagination-status');
      let currentPage = 1;

      const renderPage = () => {{
        const start = (currentPage - 1) * pageSize;
        const end = start + pageSize;
        rows.forEach((row, index) => {{
          row.hidden = index < start || index >= end;
        }});
        status.textContent = `Mostrando ${{start + 1}}-${{Math.min(end, rows.length)}} de ${{rows.length}} · Página ${{currentPage}} de ${{totalPages}}`;
        prev.disabled = currentPage === 1;
        next.disabled = currentPage === totalPages;
      }};

      prev.addEventListener('click', () => {{
        if (currentPage > 1) {{
          currentPage -= 1;
          renderPage();
        }}
      }});
      next.addEventListener('click', () => {{
        if (currentPage < totalPages) {{
          currentPage += 1;
          renderPage();
        }}
      }});
      renderPage();
    }});
  </script>
</body>
</html>
"""


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera HTML com somas mensais por grupo e categoria.",
    )
    parser.add_argument(
        "classified_workbook",
        type=Path,
        help="Caminho do Excel classificado gerado pelo classify_transactions.py.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Caminho do HTML de saida. Padrao: *_relatorio_mensal.html.",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET_NAME,
        help=f"Aba a ser lida. Padrao: {DEFAULT_SHEET_NAME}.",
    )
    parser.add_argument(
        "--ofx",
        type=Path,
        help="Arquivo OFX original para conferir saldo inicial/final e movimentacoes.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    workbook_path = args.classified_workbook
    output_path = args.output or workbook_path.with_name(
        f"{workbook_path.stem}_relatorio_mensal.html"
    )

    if not workbook_path.exists():
        print(f"Erro: Excel classificado nao encontrado: {workbook_path}", file=sys.stderr)
        return 1

    rows = read_sheet_rows(workbook_path, args.sheet)
    grouped = grouped_monthly_totals(rows)
    balance = analyze_balance(args.ofx, rows)
    html_report = render_html_report(workbook_path, rows, grouped, balance)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html_report, encoding="utf-8")

    print(f"HTML gerado: {output_path}")
    print(f"Transacoes consideradas: {len(rows)}")
    print(f"Tabelas grupo/categoria: {len(grouped)}")
    if balance is not None:
        print(f"Conferencia de saldo: {balance.status}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
