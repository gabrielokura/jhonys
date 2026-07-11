#!/usr/bin/env python
"""Convert OFX bank statement transactions to a CSV for Brazilian Excel."""

from __future__ import annotations

import argparse
import csv
import html
import re
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path


ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")

COLUMNS = [
    "data",
    "tipo",
    "valor",
    "descricao",
    "nome",
    "memo",
    "id_transacao",
    "documento",
    "banco",
    "agencia",
    "conta",
    "moeda",
]


def read_ofx(path: Path) -> tuple[str, str]:
    data = path.read_bytes()

    for encoding in ENCODINGS:
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue

    return data.decode("latin-1", errors="replace"), "latin-1"


def clean_value(value: str | None) -> str:
    if value is None:
        return ""

    value = html.unescape(value)
    value = value.replace("\x00", "")
    return " ".join(value.strip().split())


def extract_tag(text: str, tag: str) -> str:
    escaped_tag = re.escape(tag)

    closed = re.search(
        rf"<\s*{escaped_tag}\b[^>]*>(.*?)<\s*/\s*{escaped_tag}\s*>",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if closed:
        return clean_value(closed.group(1))

    opened = re.search(
        rf"<\s*{escaped_tag}\b[^>]*>\s*([^<\r\n]*)",
        text,
        flags=re.IGNORECASE,
    )
    if opened:
        return clean_value(opened.group(1))

    return ""


def extract_transactions(text: str) -> list[str]:
    pattern = re.compile(
        r"<\s*STMTTRN\b[^>]*>(.*?)(?:<\s*/\s*STMTTRN\s*>|(?=<\s*STMTTRN\b)|(?=<\s*/\s*BANKTRANLIST\s*>)|$)",
        flags=re.IGNORECASE | re.DOTALL,
    )
    return [match.group(1) for match in pattern.finditer(text)]


def format_ofx_date(value: str) -> str:
    match = re.match(r"\s*(\d{4})(\d{2})(\d{2})", value or "")
    if not match:
        return clean_value(value)

    year, month, day = match.groups()
    return f"{day}/{month}/{year}"


def format_brazilian_amount(value: str) -> str:
    raw = clean_value(value)
    if not raw:
        return ""

    normalized = raw.replace(" ", "")
    if "," in normalized and "." not in normalized:
        normalized = normalized.replace(",", ".")

    try:
        amount = Decimal(normalized).quantize(Decimal("0.01"))
    except InvalidOperation:
        return raw.replace(".", ",")

    return f"{amount:.2f}".replace(".", ",")


def transaction_to_row(block: str, account: dict[str, str]) -> dict[str, str]:
    name = extract_tag(block, "NAME")
    memo = extract_tag(block, "MEMO")
    checknum = extract_tag(block, "CHECKNUM")
    refnum = extract_tag(block, "REFNUM")

    return {
        "data": format_ofx_date(extract_tag(block, "DTPOSTED")),
        "tipo": extract_tag(block, "TRNTYPE"),
        "valor": format_brazilian_amount(extract_tag(block, "TRNAMT")),
        "descricao": name or memo,
        "nome": name,
        "memo": memo,
        "id_transacao": extract_tag(block, "FITID"),
        "documento": checknum or refnum,
        "banco": account["banco"],
        "agencia": account["agencia"],
        "conta": account["conta"],
        "moeda": account["moeda"],
    }


def extract_account_data(text: str) -> dict[str, str]:
    return {
        "banco": extract_tag(text, "BANKID"),
        "agencia": extract_tag(text, "BRANCHID"),
        "conta": extract_tag(text, "ACCTID"),
        "moeda": extract_tag(text, "CURDEF"),
    }


def convert_ofx_to_csv(input_path: Path, output_path: Path) -> tuple[int, str]:
    text, encoding = read_ofx(input_path)
    account = extract_account_data(text)
    transaction_blocks = extract_transactions(text)

    rows = [transaction_to_row(block, account) for block in transaction_blocks]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=COLUMNS,
            delimiter=";",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(rows)

    return len(rows), encoding


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Converte um arquivo OFX bancario em CSV para Excel BR.",
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
        help="Caminho do arquivo .csv de saida. Padrao: mesmo nome do OFX.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    input_path = args.ofx_file
    output_path = args.output or input_path.with_suffix(".csv")

    if not input_path.exists():
        print(f"Erro: arquivo OFX nao encontrado: {input_path}", file=sys.stderr)
        return 1

    if input_path.suffix.lower() != ".ofx":
        print(f"Aviso: o arquivo de entrada nao tem extensao .ofx: {input_path}")

    transaction_count, encoding = convert_ofx_to_csv(input_path, output_path)

    print(f"CSV gerado: {output_path}")
    print(f"Transacoes exportadas: {transaction_count}")
    print(f"Encoding lido: {encoding}")

    if transaction_count == 0:
        print("Aviso: nenhuma transacao <STMTTRN> foi encontrada.", file=sys.stderr)
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
