#!/usr/bin/env python
"""Build JSON data for the interactive monthly report v2."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from build_monthly_html_report import (
    BalanceAnalysis,
    MonthlyBucket,
    analyze_balance,
    amount_class,
    format_brl,
    format_month,
    grouped_detail_rows,
    grouped_monthly_totals,
    parse_amount,
    read_sheet_rows,
    review_rows,
)


DEFAULT_SHEET_NAME = "extrato_classificado"


def decimal_to_json(value: Decimal | None) -> str | None:
    if value is None:
        return None
    return format(value, "f")


def money_summary(value: Decimal) -> dict[str, str]:
    return {
        "value": format(value, "f"),
        "formatted": format_brl(value),
        "kind": amount_class(value),
    }


def balance_to_json(balance: BalanceAnalysis | None) -> dict[str, Any] | None:
    if balance is None:
        return None

    return {
        "status": balance.status,
        "start_date": balance.start_date,
        "end_date": balance.end_date,
        "balance_date": balance.balance_date,
        "opening_balance": decimal_to_json(balance.opening_balance),
        "opening_balance_formatted": (
            format_brl(balance.opening_balance)
            if balance.opening_balance is not None
            else "-"
        ),
        "final_balance": decimal_to_json(balance.final_balance),
        "final_balance_formatted": (
            format_brl(balance.final_balance)
            if balance.final_balance is not None
            else "-"
        ),
        "movement_total": money_summary(balance.movement_total),
        "ofx_movement_total": money_summary(balance.ofx_movement_total),
        "extraction_difference": money_summary(balance.extraction_difference),
        "reconciliation_difference": (
            money_summary(balance.reconciliation_difference)
            if balance.reconciliation_difference is not None
            else None
        ),
        "transaction_count": balance.transaction_count,
        "ofx_transaction_count": balance.ofx_transaction_count,
        "note": balance.note,
    }


def monthly_bucket_to_json(month: str, bucket: MonthlyBucket) -> dict[str, Any]:
    return {
        "month": month,
        "month_label": format_month(month),
        "quantity": bucket.quantity,
        "inflows": money_summary(bucket.inflows),
        "outflows": money_summary(bucket.outflows),
        "total": money_summary(bucket.total),
    }


def group_to_json(
    group: str,
    category: str,
    months: dict[str, MonthlyBucket],
    rows: list[dict[str, str]],
) -> dict[str, Any]:
    sorted_months = sorted(months.items())
    total_quantity = sum(bucket.quantity for _, bucket in sorted_months)
    total_inflows = sum((bucket.inflows for _, bucket in sorted_months), Decimal("0"))
    total_outflows = sum((bucket.outflows for _, bucket in sorted_months), Decimal("0"))
    total = sum((bucket.total for _, bucket in sorted_months), Decimal("0"))

    return {
        "group": group,
        "category": category,
        "quantity": total_quantity,
        "inflows": money_summary(total_inflows),
        "outflows": money_summary(total_outflows),
        "total": money_summary(total),
        "months": [
            monthly_bucket_to_json(month, bucket)
            for month, bucket in sorted_months
        ],
        "rows": rows,
    }


def build_report_data(
    workbook_path: Path,
    rows: list[dict[str, str]],
    grouped: dict[tuple[str, str], dict[str, MonthlyBucket]],
    balance: BalanceAnalysis | None = None,
) -> dict[str, Any]:
    total_amount = sum((parse_amount(row.get("valor", "")) for row in rows), Decimal("0"))
    rows_to_review = review_rows(rows)
    review_total = sum(
        (parse_amount(row.get("valor", "")) for row in rows_to_review),
        Decimal("0"),
    )
    review_inflows = sum(
        (
            amount
            for amount in (parse_amount(row.get("valor", "")) for row in rows_to_review)
            if amount > 0
        ),
        Decimal("0"),
    )
    review_outflows = sum(
        (
            amount
            for amount in (parse_amount(row.get("valor", "")) for row in rows_to_review)
            if amount < 0
        ),
        Decimal("0"),
    )
    details_by_group = grouped_detail_rows(rows)

    return {
        "version": 2,
        "source": str(workbook_path),
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "metrics": {
            "transaction_count": len(rows),
            "group_count": len(grouped),
            "total": money_summary(total_amount),
        },
        "balance": balance_to_json(balance),
        "review": {
            "quantity": len(rows_to_review),
            "inflows": money_summary(review_inflows),
            "outflows": money_summary(review_outflows),
            "total": money_summary(review_total),
            "rows": rows_to_review,
        },
        "groups": [
            group_to_json(
                group,
                category,
                months,
                details_by_group.get((group, category), []),
            )
            for (group, category), months in sorted(grouped.items())
        ],
    }


def build_report_data_from_workbook(
    workbook_path: Path,
    sheet_name: str = DEFAULT_SHEET_NAME,
    ofx_path: Path | None = None,
) -> dict[str, Any]:
    rows = read_sheet_rows(workbook_path, sheet_name)
    grouped = grouped_monthly_totals(rows)
    balance = analyze_balance(ofx_path, rows)
    return build_report_data(workbook_path, rows, grouped, balance)


def report_data_to_json(data: dict[str, Any]) -> str:
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera JSON para o relatorio mensal interativo v2.",
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
        help="Caminho do JSON de saida. Padrao: *_relatorio_mensal_v2.json.",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET_NAME,
        help=f"Aba a ser lida. Padrao: {DEFAULT_SHEET_NAME}.",
    )
    parser.add_argument(
        "--ofx",
        type=Path,
        help="Arquivo OFX original para conferencia de saldo.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    workbook_path = args.classified_workbook
    output_path = args.output or workbook_path.with_name(
        f"{workbook_path.stem}_relatorio_mensal_v2.json"
    )

    if not workbook_path.exists():
        print(f"Erro: Excel classificado nao encontrado: {workbook_path}", file=sys.stderr)
        return 1

    data = build_report_data_from_workbook(
        workbook_path,
        sheet_name=args.sheet,
        ofx_path=args.ofx,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(report_data_to_json(data), encoding="utf-8")

    print(f"JSON v2 gerado: {output_path}")
    print(f"Transacoes consideradas: {data['metrics']['transaction_count']}")
    print(f"Grupos/categorias: {data['metrics']['group_count']}")
    if data["balance"] is not None:
        print(f"Conferencia de saldo: {data['balance']['status']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
