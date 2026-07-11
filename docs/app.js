const PYODIDE_VERSION = "314.0.2";
const PYODIDE_INDEX_URL = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;
const PYTHON_ROOT = "python";
const PYTHON_TEXT_FILES = [
  "ofx_to_csv.py",
  "classify_transactions.py",
  "build_monthly_html_report.py",
];
const PYTHON_BINARY_FILES = [
  "classification_rules_completed.csv",
  "classification_rules.csv",
];

const form = document.querySelector("#ofx-form");
const fileInput = document.querySelector("#ofx-file");
const fileLabel = document.querySelector("#file-label");
const processButton = document.querySelector("#process-button");
const statusDot = document.querySelector("#status-dot");
const statusTitle = document.querySelector("#status-title");
const statusMessage = document.querySelector("#status-message");
const summary = document.querySelector("#summary");
const downloads = document.querySelector("#downloads");

let pyodideReadyPromise = null;
let objectUrls = [];

function setStatus(kind, title, message) {
  statusDot.className = `status-dot ${kind}`;
  statusTitle.textContent = title;
  statusMessage.textContent = message;
}

function clearOutputs() {
  for (const url of objectUrls) {
    URL.revokeObjectURL(url);
  }
  objectUrls = [];
  summary.hidden = true;
  summary.innerHTML = "";
  downloads.hidden = true;
  downloads.innerHTML = "";
}

function loadPyodideScript() {
  if (globalThis.loadPyodide) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = `${PYODIDE_INDEX_URL}pyodide.js`;
    script.async = true;
    script.addEventListener("load", resolve, { once: true });
    script.addEventListener("error", () => reject(new Error("Nao foi possivel carregar o Pyodide.")), { once: true });
    document.head.append(script);
  });
}

async function fetchRequiredFile(path, asBinary = false) {
  const response = await fetch(path);
  if (!response.ok) {
    throw new Error(`Arquivo nao encontrado: ${path}`);
  }
  return asBinary ? new Uint8Array(await response.arrayBuffer()) : await response.text();
}

async function preparePyodide() {
  if (!pyodideReadyPromise) {
    pyodideReadyPromise = (async () => {
      setStatus("loading", "Carregando Python", "Baixando o runtime Pyodide. Na primeira vez, isso pode levar alguns segundos.");
      await loadPyodideScript();
      const pyodide = await globalThis.loadPyodide({ indexURL: PYODIDE_INDEX_URL });

      setStatus("loading", "Preparando conversor", "Montando os modulos Python e as regras de classificacao no navegador.");
      pyodide.FS.mkdirTree("/work/python");
      pyodide.FS.mkdirTree("/work/runs");

      for (const fileName of PYTHON_TEXT_FILES) {
        const source = await fetchRequiredFile(`${PYTHON_ROOT}/${fileName}`);
        pyodide.FS.writeFile(`/work/python/${fileName}`, source, { encoding: "utf8" });
      }

      for (const fileName of PYTHON_BINARY_FILES) {
        const bytes = await fetchRequiredFile(`${PYTHON_ROOT}/${fileName}`, true);
        pyodide.FS.writeFile(`/work/python/${fileName}`, bytes);
      }

      pyodide.runPython(`
import shutil
import sys
from pathlib import Path

sys.path.insert(0, "/work/python")

from build_monthly_html_report import (
    analyze_balance,
    grouped_monthly_totals,
    read_sheet_rows,
    render_html_report,
)
from classify_transactions import classify_transactions_file
from ofx_to_csv import convert_ofx_to_csv

RUN_ROOT = Path("/work/runs")
UPLOAD_PATH = Path("/work/upload.ofx")
RULES_PATH = Path("/work/python/classification_rules_completed.csv")


def browser_safe_filename(filename):
    cleaned = Path(filename or "extrato.ofx").name
    cleaned = "".join(
        char if char.isalnum() or char in {"-", "_", ".", " "} else "_"
        for char in cleaned
    ).strip()
    if not cleaned.lower().endswith(".ofx"):
        cleaned = f"{cleaned or 'extrato'}.ofx"
    return cleaned or "extrato.ofx"


def process_browser_file(original_name):
    if not UPLOAD_PATH.exists():
        raise ValueError("Arquivo OFX nao foi carregado.")

    if RUN_ROOT.exists():
        shutil.rmtree(RUN_ROOT)
    RUN_ROOT.mkdir(parents=True, exist_ok=True)

    original_name = browser_safe_filename(original_name)
    stem = Path(original_name).stem or "extrato"
    ofx_path = RUN_ROOT / original_name
    csv_path = RUN_ROOT / f"{stem}.csv"
    xlsx_path = RUN_ROOT / f"{stem}_classificado.xlsx"
    pending_path = RUN_ROOT / f"{stem}_operacoes_a_classificar.csv"
    html_path = RUN_ROOT / f"{stem}_relatorio_mensal.html"

    ofx_path.write_bytes(UPLOAD_PATH.read_bytes())

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
        rules_path=RULES_PATH,
        pending_output_path=pending_path,
    )

    rows = read_sheet_rows(xlsx_path, "extrato_classificado")
    grouped = grouped_monthly_totals(rows)
    balance = analyze_balance(ofx_path, rows)
    html_report = render_html_report(xlsx_path, rows, grouped, balance)
    html_path.write_text(html_report, encoding="utf-8")

    return {
        "original_name": original_name,
        "encoding": encoding,
        "transaction_count": transaction_count,
        "classified_total": classified_total,
        "classified_count": classified_count,
        "pending_count": pending_count,
        "rule_count": rule_count,
        "table_count": len(grouped),
        "balance_status": balance.status if balance is not None else "sem_ofx",
        "csv": str(csv_path),
        "xlsx": str(xlsx_path),
        "pending": str(actual_pending_path),
        "html": str(html_path),
    }
`);
      return pyodide;
    })();
  }

  return pyodideReadyPromise;
}

function blobUrl(bytes, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  objectUrls.push(url);
  return url;
}

function downloadName(filePath) {
  return filePath.split("/").pop();
}

function renderMetric(label, value) {
  return `<div class="metric"><span>${label}</span><strong>${value}</strong></div>`;
}

function renderDownloads(pyodide, result) {
  const outputs = [
    ["Relatorio HTML", result.html, "text/html;charset=utf-8", false],
    ["Excel classificado", result.xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", true],
    ["CSV bruto", result.csv, "text/csv;charset=utf-8", true],
    ["CSV pendencias", result.pending, "text/csv;charset=utf-8", true],
  ];

  downloads.innerHTML = outputs
    .map(([label, filePath, type, secondary]) => {
      const bytes = pyodide.FS.readFile(filePath);
      const url = blobUrl(bytes, type);
      const className = secondary ? "download-link secondary" : "download-link";
      return `<a class="${className}" href="${url}" download="${downloadName(filePath)}">${label}</a>`;
    })
    .join("");
  downloads.hidden = false;
}

function renderSummary(result) {
  summary.innerHTML = [
    renderMetric("Transacoes", result.transaction_count),
    renderMetric("Classificadas", `${result.classified_count} / ${result.classified_total}`),
    renderMetric("Pendencias", result.pending_count),
    renderMetric("Regras", result.rule_count),
    renderMetric("Tabelas", result.table_count),
    renderMetric("Saldo", result.balance_status),
  ].join("");
  summary.hidden = false;
}

async function processFile(file) {
  clearOutputs();
  setStatus("loading", "Processando", "Lendo o OFX e preparando os arquivos de saida.");
  processButton.disabled = true;

  try {
    const pyodide = await preparePyodide();
    const bytes = new Uint8Array(await file.arrayBuffer());
    pyodide.FS.writeFile("/work/upload.ofx", bytes);
    pyodide.globals.set("browser_original_name", file.name);

    const pyResult = pyodide.runPython("process_browser_file(browser_original_name)");
    const result = pyResult.toJs({ dict_converter: Object.fromEntries });
    pyResult.destroy();

    renderSummary(result);
    renderDownloads(pyodide, result);

    const message = result.transaction_count === 0
      ? "Nenhuma transacao STMTTRN foi encontrada, mas os arquivos foram gerados para conferencia."
      : `Arquivo processado: ${result.original_name}.`;
    setStatus("success", "Arquivos prontos", message);
  } catch (error) {
    console.error(error);
    setStatus("error", "Falha no processamento", error.message || String(error));
  } finally {
    processButton.disabled = !fileInput.files.length;
  }
}

fileInput.addEventListener("change", () => {
  clearOutputs();
  const file = fileInput.files[0];
  processButton.disabled = !file;
  fileLabel.textContent = file ? file.name : "ou selecione um arquivo";
  setStatus("idle", file ? "Arquivo selecionado" : "Aguardando arquivo", file ? "Pronto para processar." : "Os arquivos bancarios ficam no seu navegador durante o processamento.");
});

for (const eventName of ["dragenter", "dragover"]) {
  form.addEventListener(eventName, (event) => {
    event.preventDefault();
    form.classList.add("dragover");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  form.addEventListener(eventName, (event) => {
    event.preventDefault();
    form.classList.remove("dragover");
  });
}

form.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) {
    return;
  }
  fileInput.files = event.dataTransfer.files;
  fileInput.dispatchEvent(new Event("change"));
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const file = fileInput.files[0];
  if (!file) {
    return;
  }
  await processFile(file);
});
