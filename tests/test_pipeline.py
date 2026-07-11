from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from build_monthly_report_v2 import build_report_data, report_data_to_json
from build_monthly_html_report import (
    analyze_balance,
    grouped_monthly_totals,
    read_sheet_rows,
    render_detail_table,
    render_html_report,
)
from classify_transactions import classify_transactions_file
from ofx_to_csv import convert_ofx_to_csv


ROOT = Path(__file__).resolve().parents[1]


class PipelineTest(unittest.TestCase):
    def test_known_public_rules_do_not_create_review_items(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            workdir = Path(tmpdir)
            ofx_path = workdir / "sample.ofx"
            csv_path = workdir / "sample.csv"
            xlsx_path = workdir / "sample_classificado.xlsx"
            pending_path = workdir / "sample_operacoes_a_classificar.csv"

            ofx_path.write_text(
                """OFX
<BANKMSGSRSV1><STMTTRNRS><STMTRS><CURDEF>BRL
<BANKACCTFROM><BANKID>756<BRANCHID>1234<ACCTID>5678</BANKACCTFROM>
<BANKTRANLIST><DTSTART>20260601<DTEND>20260630
<STMTTRN><TRNTYPE>CREDIT<DTPOSTED>20260626000000[0:GMT]<TRNAMT>93.00<FITID>2026062693001<NAME>CREDITO REEMBOLSO ALUGUEL SIPAG<MEMO>CREDITO REEMBOLSO ALUGUEL SIPAG</STMTTRN>
<STMTTRN><TRNTYPE>DEBIT<DTPOSTED>20260623000000[0:GMT]<TRNAMT>-76.00<FITID>2026062376001<NAME>TAR. FORNEC. FOLHA CHEQUE TALÃO<MEMO>TAR. FORNEC. FOLHA CHEQUE TALÃO</STMTTRN>
</BANKTRANLIST><LEDGERBAL><BALAMT>17.00<DTASOF>20260630</LEDGERBAL></STMTRS></STMTTRNRS></BANKMSGSRSV1>
""",
                encoding="utf-8",
            )

            transaction_count, encoding = convert_ofx_to_csv(ofx_path, csv_path)
            self.assertEqual(transaction_count, 2)
            self.assertEqual(encoding, "utf-8-sig")

            total, classified, pending, rules, actual_pending_path = (
                classify_transactions_file(
                    input_path=csv_path,
                    output_path=xlsx_path,
                    rules_path=ROOT / "classification_rules_completed.csv",
                    pending_output_path=pending_path,
                )
            )
            self.assertEqual(total, 2)
            self.assertEqual(classified, 2)
            self.assertEqual(pending, 0)
            self.assertGreaterEqual(rules, 32)
            self.assertTrue(actual_pending_path.exists())

            rows = read_sheet_rows(xlsx_path, "extrato_classificado")
            classifications = {
                row["memo"]: (
                    row["status_classificacao"],
                    row["grupo"],
                    row["categoria"],
                    row["tipo_fluxo"],
                )
                for row in rows
            }
            self.assertEqual(
                classifications["CREDITO REEMBOLSO ALUGUEL SIPAG"],
                ("classificado", "Reembolso", "Aluguel", "Entrada"),
            )
            self.assertEqual(
                classifications["TAR. FORNEC. FOLHA CHEQUE TALÃO"],
                ("classificado", "Despesas", "Tarifas bancárias", "Saida"),
            )

            balance = analyze_balance(ofx_path, rows)
            self.assertIsNotNone(balance)
            self.assertEqual(balance.status, "ok")

            report = render_html_report(
                xlsx_path,
                rows,
                grouped_monthly_totals(rows),
                balance,
            )
            self.assertIn("Content-Security-Policy", report)
            self.assertIn("Movimentações Para Avaliação", report)
            self.assertIn("Nenhuma movimentação nova ou sem regra", report)

            report_data = build_report_data(
                xlsx_path,
                rows,
                grouped_monthly_totals(rows),
                balance,
            )
            self.assertEqual(report_data["version"], 2)
            self.assertEqual(report_data["metrics"]["transaction_count"], 2)
            self.assertEqual(report_data["review"]["quantity"], 0)
            self.assertEqual(report_data["balance"]["status"], "ok")
            self.assertEqual(len(report_data["groups"]), 2)
            self.assertIn("CREDITO REEMBOLSO ALUGUEL SIPAG", report_data_to_json(report_data))


class PublicRulesTest(unittest.TestCase):
    rule_files = [
        ROOT / "classification_rules.csv",
        ROOT / "classification_rules_completed.csv",
        ROOT / "docs/python/classification_rules.csv",
        ROOT / "docs/python/classification_rules_completed.csv",
    ]

    def test_public_rules_are_synchronized_and_sanitized(self) -> None:
        contents = [path.read_text(encoding="utf-8-sig") for path in self.rule_files]
        self.assertTrue(all(content == contents[0] for content in contents))

        forbidden_fragments = [
            "REM.:",
            "FAV.:",
            "CPF",
            "CNPJ",
            "SENHA",
            "TOKEN",
            "SECRET",
            "PASSWORD",
            "API_KEY",
        ]
        upper_content = contents[0].upper()
        for fragment in forbidden_fragments:
            self.assertNotIn(fragment, upper_content)

        rows = list(csv.DictReader(contents[0].splitlines(), delimiter=";"))
        self.assertGreaterEqual(len(rows), 32)
        for row in rows:
            self.assertEqual(
                row.get("observacao"),
                "Regra publica generica",
                msg=f"Regra publica sem observacao padrao: {row}",
            )


class StaticSiteTest(unittest.TestCase):
    def test_detail_tables_paginate_after_twenty_rows(self) -> None:
        rows = [
            {
                "data": f"01/01/2026",
                "descricao": f"linha {index}",
                "valor": "1,00",
            }
            for index in range(21)
        ]

        paginated = render_detail_table(rows, ["data", "descricao", "valor"])
        self.assertIn('class="table-pagination"', paginated)
        self.assertIn("Mostrando 1-20 de 21 · Página 1 de 2", paginated)
        self.assertIn('data-page-size="20"', paginated)
        self.assertIn('aria-label="Página anterior"', paginated)
        self.assertIn('aria-label="Próxima página"', paginated)

        short_table = render_detail_table(rows[:20], ["data", "descricao", "valor"])
        self.assertNotIn('class="table-pagination"', short_table)

    def test_pages_entrypoints_and_python_copies_exist(self) -> None:
        root_index = (ROOT / "index.html").read_text(encoding="utf-8")
        docs_index = (ROOT / "docs/index.html").read_text(encoding="utf-8")
        app_js = (ROOT / "docs/app.js").read_text(encoding="utf-8")
        report_v2_html = (ROOT / "docs/report-v2.html").read_text(encoding="utf-8")
        report_v2_js = (ROOT / "docs/report-v2.js").read_text(encoding="utf-8")

        self.assertIn('url=docs/', root_index)
        self.assertIn('id="clear-button"', docs_index)
        self.assertIn('role="status"', docs_index)
        self.assertIn('id="status-steps"', docs_index)
        self.assertIn("MAX_UPLOAD_BYTES", app_js)
        self.assertIn("setProcessingStep", app_js)
        self.assertIn("Abrir relatório interativo", app_js)
        self.assertIn("report-v2.html?dataKey=", app_js)
        self.assertIn('id="groups"', report_v2_html)
        self.assertIn("fetch(dataUrl)", report_v2_js)
        self.assertIn("readStoredReport(dataKey)", report_v2_js)
        self.assertIn("storage.removeItem(key)", report_v2_js)

        for file_name in [
            "ofx_to_csv.py",
            "classify_transactions.py",
            "build_monthly_html_report.py",
            "build_monthly_report_v2.py",
        ]:
            self.assertEqual(
                (ROOT / file_name).read_text(encoding="utf-8"),
                (ROOT / "docs/python" / file_name).read_text(encoding="utf-8"),
                msg=f"docs/python/{file_name} esta fora de sincronia",
            )


if __name__ == "__main__":
    unittest.main()
