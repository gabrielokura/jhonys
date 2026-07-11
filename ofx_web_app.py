#!/usr/bin/env python
"""Local drag-and-drop OFX analyzer."""

from __future__ import annotations

import argparse
import html
import mimetypes
import re
import secrets
import shutil
import sys
import time
import webbrowser
from http.cookies import SimpleCookie
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from build_monthly_html_report import (
    analyze_balance,
    grouped_monthly_totals,
    read_sheet_rows,
    render_html_report,
)
from classify_transactions import classify_transactions_file
from ofx_to_csv import convert_ofx_to_csv


APP_ROOT = Path(__file__).resolve().parent
RUN_ROOT = APP_ROOT / "web_runs"
SESSION_ROOT = RUN_ROOT / "sessions"
RULES_FILES = ("classification_rules_completed.csv", "classification_rules.csv")
SESSION_COOKIE = "ofx_session"
SESSION_ID_RE = re.compile(r"^[A-Za-z0-9_-]{24,80}$")
APP_CONFIG = {
    "ttl_seconds": 60 * 60,
    "max_upload_bytes": 20 * 1024 * 1024,
}


def rules_path() -> Path:
    for file_name in RULES_FILES:
        path = APP_ROOT / file_name
        if path.exists():
            return path
    return APP_ROOT / RULES_FILES[-1]


def safe_filename(filename: str) -> str:
    cleaned = Path(filename or "extrato.ofx").name
    cleaned = "".join(
        char if char.isalnum() or char in {"-", "_", ".", " "} else "_"
        for char in cleaned
    ).strip()
    return cleaned or "extrato.ofx"


def new_session_id() -> str:
    return secrets.token_urlsafe(32)


def valid_session_id(session_id: str | None) -> bool:
    return bool(session_id and SESSION_ID_RE.match(session_id))


def session_path(session_id: str) -> Path:
    if not valid_session_id(session_id):
        raise ValueError("Sessao invalida.")
    return SESSION_ROOT / session_id


def delete_session(session_id: str | None) -> None:
    if not valid_session_id(session_id):
        return
    path = session_path(session_id)
    if path.exists():
        shutil.rmtree(path)


def cleanup_old_sessions() -> None:
    SESSION_ROOT.mkdir(parents=True, exist_ok=True)
    now = time.time()
    ttl_seconds = int(APP_CONFIG["ttl_seconds"])
    for child in SESSION_ROOT.iterdir():
        if not child.is_dir():
            continue
        try:
            if now - child.stat().st_mtime > ttl_seconds:
                shutil.rmtree(child)
        except FileNotFoundError:
            continue


def touch_session(session_id: str) -> None:
    try:
        path = session_path(session_id)
        if path.exists():
            now = time.time()
            Path(path).touch()
            for output in path.iterdir():
                output.touch(exist_ok=True)
            path.touch()
            try:
                import os

                os.utime(path, (now, now))
            except OSError:
                pass
    except ValueError:
        return


def session_file_path(session_id: str, file_name: str) -> Path:
    root = session_path(session_id).resolve()
    requested = (root / file_name).resolve()
    if requested != root and root not in requested.parents:
        raise ValueError("Caminho invalido.")
    return requested


def process_ofx_upload(
    file_name: str,
    content: bytes,
    previous_session_id: str | None = None,
) -> dict[str, object]:
    delete_session(previous_session_id)
    session_id = new_session_id()
    run_path = session_path(session_id)
    run_path.mkdir(parents=True, exist_ok=True)

    original_name = safe_filename(file_name)
    stem = Path(original_name).stem or "extrato"
    ofx_path = run_path / original_name
    csv_path = run_path / f"{stem}.csv"
    xlsx_path = run_path / f"{stem}_classificado.xlsx"
    pending_path = run_path / f"{stem}_operacoes_a_classificar.csv"
    html_path = run_path / f"{stem}_relatorio_mensal.html"

    ofx_path.write_bytes(content)

    transaction_count, encoding = convert_ofx_to_csv(ofx_path, csv_path)
    (
        classified_total,
        classified_count,
        pending_count,
        rule_count,
        actual_pending_path,
    ) = classify_transactions_file(
        input_path=csv_path,
        output_path=xlsx_path,
        rules_path=rules_path(),
        pending_output_path=pending_path,
    )

    rows = read_sheet_rows(xlsx_path, "extrato_classificado")
    grouped = grouped_monthly_totals(rows)
    balance = analyze_balance(ofx_path, rows)
    html_report = render_html_report(xlsx_path, rows, grouped, balance)
    html_path.write_text(html_report, encoding="utf-8")
    touch_session(session_id)

    return {
        "session_id": session_id,
        "original_name": original_name,
        "encoding": encoding,
        "transaction_count": transaction_count,
        "classified_total": classified_total,
        "classified_count": classified_count,
        "pending_count": pending_count,
        "rule_count": rule_count,
        "table_count": len(grouped),
        "balance_status": balance.status if balance is not None else "sem_ofx",
        "csv": csv_path.name,
        "xlsx": xlsx_path.name,
        "pending": actual_pending_path.name,
        "html": html_path.name,
    }


def latest_outputs(session_id: str | None) -> dict[str, str]:
    if not valid_session_id(session_id):
        return {}
    session = session_path(session_id)
    if not session.exists():
        return {}
    output = {}
    for key, pattern in [
        ("html", "*_relatorio_mensal.html"),
        ("xlsx", "*_classificado.xlsx"),
        ("csv", "*.csv"),
    ]:
        matches = sorted(session.glob(pattern), key=lambda path: path.stat().st_mtime)
        if matches:
            output[key] = matches[-1].name
    return output


def page_shell(title: str, body: str) -> bytes:
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    :root {{
      --bg: #f5f7fa;
      --panel: #ffffff;
      --text: #1f2933;
      --muted: #637083;
      --line: #d8dee8;
      --primary: #245f9e;
      --danger: #a73535;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: Arial, Helvetica, sans-serif;
    }}
    main {{
      width: min(860px, calc(100% - 28px));
      margin: 32px auto;
    }}
    h1 {{ margin: 0 0 8px; font-size: 28px; }}
    p {{ color: var(--muted); }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 22px;
      margin: 18px 0;
    }}
    .dropzone {{
      border: 2px dashed #9db2c7;
      border-radius: 8px;
      padding: 32px 22px;
      text-align: center;
      background: #fbfdff;
    }}
    .dropzone.dragover {{
      background: #eef6ff;
      border-color: var(--primary);
    }}
    input[type=file] {{ margin: 14px 0; }}
    button, .button {{
      border: 1px solid var(--primary);
      background: var(--primary);
      color: white;
      border-radius: 6px;
      padding: 10px 14px;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-block;
      margin: 4px 6px 4px 0;
    }}
    .secondary {{
      background: white;
      color: var(--primary);
    }}
    .danger {{
      border-color: var(--danger);
      background: var(--danger);
    }}
    .meta {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
      margin-top: 12px;
    }}
    .meta div {{
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px;
      background: #fbfcfe;
    }}
    .meta span {{ display: block; color: var(--muted); font-size: 12px; }}
    .meta strong {{ display: block; margin-top: 3px; }}
    @media (max-width: 640px) {{
      .meta {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <main>{body}</main>
  <script>
    const dropzone = document.querySelector('.dropzone');
    const input = document.querySelector('input[type=file]');
    if (dropzone && input) {{
      ['dragenter', 'dragover'].forEach((name) => {{
        dropzone.addEventListener(name, (event) => {{
          event.preventDefault();
          dropzone.classList.add('dragover');
        }});
      }});
      ['dragleave', 'drop'].forEach((name) => {{
        dropzone.addEventListener(name, (event) => {{
          event.preventDefault();
          dropzone.classList.remove('dragover');
        }});
      }});
      dropzone.addEventListener('drop', (event) => {{
        if (event.dataTransfer.files.length) {{
          input.files = event.dataTransfer.files;
          input.form.submit();
        }}
      }});
    }}
  </script>
</body>
</html>
""".encode("utf-8")


def session_cookie_header(session_id: str) -> str:
    return (
        f"{SESSION_COOKIE}={session_id}; "
        f"Path=/; HttpOnly; SameSite=Lax; Max-Age={int(APP_CONFIG['ttl_seconds'])}"
    )


def expired_session_cookie_header() -> str:
    return f"{SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0"


def session_from_cookie(cookie_header: str | None) -> str | None:
    if not cookie_header:
        return None
    cookie = SimpleCookie()
    cookie.load(cookie_header)
    morsel = cookie.get(SESSION_COOKIE)
    if morsel is None or not valid_session_id(morsel.value):
        return None
    return morsel.value


def home_page(session_id: str | None = None, message: str = "") -> bytes:
    outputs = latest_outputs(session_id)
    message_html = f'<div class="panel"><p>{html.escape(message)}</p></div>' if message else ""
    outputs_html = ""
    if outputs:
        touch_session(session_id or "")
        outputs_html = f"""
        <div class="panel">
          <h2>Ultima analise</h2>
          <p>Abra o HTML para ver o resumo mensal, a conferencia de saldo e os detalhes por grupo/categoria.</p>
          {'<a class="button" href="/files/' + html.escape(outputs["html"]) + '" target="_blank" rel="noopener">Abrir relatorio HTML</a>' if "html" in outputs else ''}
          {'<a class="button secondary" href="/files/' + html.escape(outputs["xlsx"]) + '">Baixar Excel classificado</a>' if "xlsx" in outputs else ''}
          <form method="post" action="/delete" style="margin-top: 12px;">
            <button class="danger" type="submit">Deletar informacoes e enviar outro arquivo</button>
          </form>
        </div>
        """

    return page_shell(
        "Analisador OFX",
        f"""
        <h1>Analisador OFX</h1>
        <p>Arraste um arquivo .ofx para gerar CSV, Excel classificado e HTML mensal agrupado. Os arquivos gerados ficam temporarios e isolados na sessao deste navegador.</p>
        {message_html}
        <div class="panel">
          <form method="post" action="/upload" enctype="multipart/form-data">
            <div class="dropzone">
              <strong>Arraste o arquivo OFX aqui</strong>
              <p>ou selecione manualmente</p>
              <input type="file" name="ofx_file" accept=".ofx" required>
              <br>
              <button type="submit">Processar OFX</button>
            </div>
          </form>
        </div>
        {outputs_html}
        """,
    )


def result_page(result: dict[str, object]) -> bytes:
    return page_shell(
        "Analise gerada",
        f"""
        <h1>Analise gerada</h1>
        <p>Arquivo processado: {html.escape(str(result["original_name"]))}</p>
        <div class="panel">
          <a class="button" href="/files/{html.escape(str(result["html"]))}" target="_blank" rel="noopener">Abrir relatorio HTML</a>
          <a class="button secondary" href="/files/{html.escape(str(result["xlsx"]))}">Baixar Excel classificado</a>
          <a class="button secondary" href="/files/{html.escape(str(result["pending"]))}">Baixar pendencias</a>
          <form method="post" action="/delete" style="margin-top: 12px;">
            <button class="danger" type="submit">Deletar informacoes e enviar outro arquivo</button>
          </form>
        </div>
        <div class="panel">
          <h2>Resumo do processamento</h2>
          <div class="meta">
            <div><span>Transacoes OFX</span><strong>{result["transaction_count"]}</strong></div>
            <div><span>Transacoes classificadas</span><strong>{result["classified_count"]} / {result["classified_total"]}</strong></div>
            <div><span>Pendencias</span><strong>{result["pending_count"]}</strong></div>
            <div><span>Tabelas grupo/categoria</span><strong>{result["table_count"]}</strong></div>
            <div><span>Regras carregadas</span><strong>{result["rule_count"]}</strong></div>
            <div><span>Conferencia de saldo</span><strong>{html.escape(str(result["balance_status"]))}</strong></div>
          </div>
        </div>
        """,
    )


def parse_upload(headers, body: bytes) -> tuple[str, bytes]:
    content_type = headers.get("Content-Type", "")
    message = BytesParser(policy=default).parsebytes(
        b"Content-Type: "
        + content_type.encode("utf-8")
        + b"\r\nMIME-Version: 1.0\r\n\r\n"
        + body
    )

    for part in message.iter_parts():
        if part.get_param("name", header="content-disposition") != "ofx_file":
            continue
        file_name = part.get_filename() or "extrato.ofx"
        payload = part.get_payload(decode=True) or b""
        return file_name, payload

    raise ValueError("Arquivo OFX nao encontrado no upload.")


class OFXRequestHandler(BaseHTTPRequestHandler):
    def request_session_id(self) -> str | None:
        return session_from_cookie(self.headers.get("Cookie"))

    def send_bytes(
        self,
        data: bytes,
        content_type: str = "text/html; charset=utf-8",
        extra_headers: dict[str, str] | None = None,
    ) -> None:
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        for header, value in (extra_headers or {}).items():
            self.send_header(header, value)
        self.end_headers()
        self.wfile.write(data)

    def send_redirect(self, location: str) -> None:
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        self.end_headers()

    def do_GET(self) -> None:
        cleanup_old_sessions()
        parsed = urlparse(self.path)
        session_id = self.request_session_id()
        if parsed.path == "/":
            self.send_bytes(home_page(session_id=session_id))
            return

        if parsed.path.startswith("/files/"):
            if not valid_session_id(session_id):
                self.send_error(HTTPStatus.FORBIDDEN, "Sessao nao encontrada")
                return
            file_name = unquote(parsed.path.removeprefix("/files/"))
            try:
                path = session_file_path(session_id or "", file_name)
            except ValueError:
                self.send_error(HTTPStatus.BAD_REQUEST, "Caminho invalido")
                return
            if not path.exists() or not path.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "Arquivo nao encontrado")
                return
            content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            data = path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Content-Disposition", f'inline; filename="{path.name}"')
            self.end_headers()
            self.wfile.write(data)
            touch_session(session_id or "")
            return

        self.send_error(HTTPStatus.NOT_FOUND, "Pagina nao encontrada")

    def do_POST(self) -> None:
        cleanup_old_sessions()
        parsed = urlparse(self.path)
        session_id = self.request_session_id()
        if parsed.path == "/delete":
            delete_session(session_id)
            self.send_bytes(
                home_page(message="Informacoes apagadas. Voce ja pode enviar outro arquivo."),
                extra_headers={"Set-Cookie": expired_session_cookie_header()},
            )
            return

        if parsed.path == "/upload":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length > int(APP_CONFIG["max_upload_bytes"]):
                    raise ValueError("Arquivo muito grande. Envie um OFX de ate 20 MB.")
                body = self.rfile.read(length)
                file_name, payload = parse_upload(self.headers, body)
                if not payload:
                    raise ValueError("Arquivo vazio.")
                if not file_name.lower().endswith(".ofx"):
                    raise ValueError("Envie um arquivo com extensao .ofx.")
                result = process_ofx_upload(file_name, payload, session_id)
                self.send_bytes(
                    result_page(result),
                    extra_headers={
                        "Set-Cookie": session_cookie_header(str(result["session_id"]))
                    },
                )
            except Exception as exc:
                self.send_bytes(home_page(session_id=session_id, message=f"Erro ao processar arquivo: {exc}"))
            return

        self.send_error(HTTPStatus.NOT_FOUND, "Pagina nao encontrada")

    def log_message(self, format: str, *args) -> None:
        print(f"{self.address_string()} - {format % args}")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Abre um app local para analisar arquivos OFX.")
    parser.add_argument("--host", default="127.0.0.1", help="Host do servidor local.")
    parser.add_argument("--port", type=int, default=8765, help="Porta do servidor local.")
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Nao abrir o navegador automaticamente.",
    )
    parser.add_argument(
        "--ttl-minutes",
        type=int,
        default=60,
        help="Minutos para manter arquivos temporarios de cada sessao.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    APP_CONFIG["ttl_seconds"] = max(args.ttl_minutes, 1) * 60
    RUN_ROOT.mkdir(parents=True, exist_ok=True)
    SESSION_ROOT.mkdir(parents=True, exist_ok=True)
    cleanup_old_sessions()

    server = ThreadingHTTPServer((args.host, args.port), OFXRequestHandler)
    url = f"http://{args.host}:{args.port}/"
    print(f"Analisador OFX rodando em {url}")
    print(f"Arquivos temporarios expiram em {args.ttl_minutes} minuto(s).")
    print("Pressione Ctrl+C para encerrar.")

    if not args.no_open:
        webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.")
    finally:
        server.server_close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
