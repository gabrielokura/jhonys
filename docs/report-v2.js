const PAGE_SIZE = 20;
const params = new URLSearchParams(window.location.search);
const hashParams = new URLSearchParams(window.location.hash.replace(/^#/, ""));
const dataKey = params.get("dataKey");
const dataUrl = hashParams.get("data");
const emptyState = document.querySelector("#empty-state");
const reportContent = document.querySelector("#report-content");
const reportSource = document.querySelector("#report-source");
const metrics = document.querySelector("#metrics");
const balanceSection = document.querySelector("#balance-section");
const reviewSection = document.querySelector("#review-section");
const groupCount = document.querySelector("#group-count");
const groupsContainer = document.querySelector("#groups");
const groupSearch = document.querySelector("#group-search");

let reportData = null;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function moneyClass(money) {
  return money?.kind || "neutral";
}

function moneyText(money) {
  return money?.formatted || "-";
}

function renderMetric(label, value, className = "") {
  return `<div class="metric"><span>${escapeHtml(label)}</span><strong class="${className}">${escapeHtml(value)}</strong></div>`;
}

function renderMoneyMetric(label, money) {
  return renderMetric(label, moneyText(money), moneyClass(money));
}

function renderBalance(balance) {
  if (!balance) {
    balanceSection.hidden = true;
    return;
  }

  const statusLabel = {
    ok: "Batendo",
    atencao: "Atenção",
    sem_saldo_final: "Sem saldo final",
    sem_ofx: "Sem OFX",
  }[balance.status] || balance.status;

  balanceSection.innerHTML = `
    <h2>Conferência de saldo</h2>
    <div class="balance-grid">
      <div><span>Status</span><strong>${escapeHtml(statusLabel)}</strong></div>
      <div><span>Período</span><strong>${escapeHtml(balance.start_date)} a ${escapeHtml(balance.end_date)}</strong></div>
      <div><span>Data do saldo</span><strong>${escapeHtml(balance.balance_date)}</strong></div>
      <div><span>Saldo inicial inferido</span><strong>${escapeHtml(balance.opening_balance_formatted)}</strong></div>
      <div><span>Movimentações</span><strong class="${moneyClass(balance.movement_total)}">${moneyText(balance.movement_total)}</strong></div>
      <div><span>Saldo final OFX</span><strong>${escapeHtml(balance.final_balance_formatted)}</strong></div>
      <div><span>Diferença extração x OFX</span><strong class="${moneyClass(balance.extraction_difference)}">${moneyText(balance.extraction_difference)}</strong></div>
      <div><span>Diferença conciliação</span><strong>${balance.reconciliation_difference ? moneyText(balance.reconciliation_difference) : "-"}</strong></div>
      <div><span>Transações</span><strong>${balance.transaction_count} / ${balance.ofx_transaction_count}</strong></div>
    </div>
    <p>${escapeHtml(balance.note)}</p>
  `;
}

function renderRowsTable(rows, columns, tableClass = "") {
  const header = columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("");
  const body = rows.map((row) => {
    const cells = columns.map((column) => `<td>${escapeHtml(row[column] || "")}</td>`).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  return `
    <div class="table-wrap">
      <table class="${tableClass}">
        <thead><tr>${header}</tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  `;
}

function attachPagination(container, rows) {
  const tableRows = rows.length ? rows : Array.from(container.querySelectorAll("tbody tr"));
  const pagination = container.querySelector(".pagination");
  if (!pagination || tableRows.length <= PAGE_SIZE) {
    return;
  }

  const prev = pagination.querySelector("[data-prev]");
  const next = pagination.querySelector("[data-next]");
  const status = pagination.querySelector("[data-status]");
  const totalPages = Math.ceil(tableRows.length / PAGE_SIZE);
  let page = 1;

  function renderPage() {
    const start = (page - 1) * PAGE_SIZE;
    const end = start + PAGE_SIZE;
    tableRows.forEach((row, index) => {
      row.hidden = index < start || index >= end;
    });
    status.textContent = `Mostrando ${start + 1}-${Math.min(end, rows.length)} de ${rows.length} · Página ${page} de ${totalPages}`;
    prev.disabled = page === 1;
    next.disabled = page === totalPages;
  }

  prev.addEventListener("click", () => {
    if (page > 1) {
      page -= 1;
      renderPage();
    }
  });

  next.addEventListener("click", () => {
    if (page < totalPages) {
      page += 1;
      renderPage();
    }
  });

  renderPage();
}

function paginationHtml(rows) {
  if (rows.length <= PAGE_SIZE) {
    return "";
  }
  const totalPages = Math.ceil(rows.length / PAGE_SIZE);
  return `
    <div class="pagination">
      <button type="button" data-prev aria-label="Página anterior">Anterior</button>
      <span data-status aria-live="polite">Mostrando 1-${Math.min(PAGE_SIZE, rows.length)} de ${rows.length} · Página 1 de ${totalPages}</span>
      <button type="button" data-next aria-label="Próxima página">Próxima</button>
    </div>
  `;
}

function renderReview(review) {
  if (!review || review.quantity === 0) {
    reviewSection.innerHTML = `
      <h2>Movimentações para avaliação</h2>
      <p>Nenhuma movimentação nova ou sem regra foi encontrada.</p>
    `;
    return;
  }

  const columns = [
    "data",
    "tipo",
    "valor",
    "descricao",
    "memo",
    "id_transacao",
    "grupo",
    "categoria",
    "motivo_classificacao",
  ];
  reviewSection.innerHTML = `
    <h2>Movimentações para avaliação</h2>
    <p>${review.quantity} linha(s) precisam ser avaliadas para futura incorporação nas regras.</p>
    <div class="review-metrics">
      <div><span>Entradas</span><strong class="${moneyClass(review.inflows)}">${moneyText(review.inflows)}</strong></div>
      <div><span>Saídas</span><strong class="${moneyClass(review.outflows)}">${moneyText(review.outflows)}</strong></div>
      <div><span>Saldo</span><strong class="${moneyClass(review.total)}">${moneyText(review.total)}</strong></div>
    </div>
    ${renderRowsTable(review.rows, columns, "detail-table")}
    ${paginationHtml(review.rows)}
  `;
  attachPagination(reviewSection, review.rows);
}

function renderMonthTable(months) {
  const rows = months.map((month) => `
    <tr>
      <td>${escapeHtml(month.month_label)}</td>
      <td class="number">${month.quantity}</td>
      <td class="money positive">${moneyText(month.inflows)}</td>
      <td class="money negative">${moneyText(month.outflows)}</td>
      <td class="money ${moneyClass(month.total)}">${moneyText(month.total)}</td>
    </tr>
  `).join("");

  return `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Mês</th><th>Qtde.</th><th>Entradas</th><th>Saídas</th><th>Total mensal</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderGroupCard(group, index) {
  const detailId = `group-details-${index}`;
  const columns = group.rows.length ? Object.keys(group.rows[0]) : [];
  return `
    <article class="group-card" data-search="${escapeHtml(`${group.group} ${group.category} ${JSON.stringify(group.rows)}`.toLowerCase())}">
      <div class="group-card-header">
        <div>
          <h3>${escapeHtml(group.group)}</h3>
          <p>${escapeHtml(group.category)} · ${group.quantity} transação(ões) · <strong class="${moneyClass(group.total)}">${moneyText(group.total)}</strong></p>
        </div>
        <button type="button" data-toggle="${detailId}" aria-controls="${detailId}" aria-expanded="false">Ver detalhes</button>
      </div>
      ${renderMonthTable(group.months)}
      <div id="${detailId}" hidden>
        ${renderRowsTable(group.rows, columns, "detail-table")}
        ${paginationHtml(group.rows)}
      </div>
    </article>
  `;
}

function attachGroupInteractions() {
  groupsContainer.querySelectorAll("[data-toggle]").forEach((button) => {
    button.addEventListener("click", () => {
      const target = document.getElementById(button.dataset.toggle);
      const isHidden = target.hasAttribute("hidden");
      target.toggleAttribute("hidden", !isHidden);
      button.textContent = isHidden ? "Ocultar detalhes" : "Ver detalhes";
      button.setAttribute("aria-expanded", String(isHidden));
    });
  });

  groupsContainer.querySelectorAll("[id^='group-details-']").forEach((detailContainer) => {
    const detailRows = Array.from(detailContainer.querySelectorAll(".detail-table tbody tr"));
    attachPagination(detailContainer, detailRows);
  });
}

function renderGroups(groups) {
  groupsContainer.innerHTML = groups.map(renderGroupCard).join("");
  groupCount.textContent = `${groups.length} grupo(s)/categoria(s)`;
  attachGroupInteractions();
}

function applySearch() {
  const query = groupSearch.value.trim().toLowerCase();
  let visibleCount = 0;
  groupsContainer.querySelectorAll(".group-card").forEach((card) => {
    const isHidden = query && !card.dataset.search.includes(query);
    card.hidden = isHidden;
    if (!isHidden) {
      visibleCount += 1;
    }
  });
  groupCount.textContent = query
    ? `${visibleCount} de ${reportData.groups.length} grupo(s)/categoria(s)`
    : `${reportData.groups.length} grupo(s)/categoria(s)`;
}

function readStoredReport(key) {
  let raw = null;
  for (const storageName of ["localStorage", "sessionStorage"]) {
    try {
      const storage = window[storageName];
      raw = raw || storage.getItem(key);
      storage.removeItem(key);
    } catch (error) {
      console.warn("Nao foi possivel acessar o armazenamento do relatorio v2.", error);
    }
  }
  return raw;
}

async function loadReportData() {
  if (dataUrl) {
    const response = await fetch(dataUrl);
    if (!response.ok) {
      return null;
    }
    return await response.json();
  }

  if (!dataKey) {
    return null;
  }
  const raw = readStoredReport(dataKey);
  if (!raw) {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch (error) {
    console.error(error);
    return null;
  }
}

function renderReport(data) {
  reportSource.textContent = `${data.source} · Gerado em ${data.generated_at}`;
  metrics.innerHTML = [
    renderMetric("Transações consideradas", data.metrics.transaction_count),
    renderMetric("Grupos/Categorias", data.metrics.group_count),
    renderMoneyMetric("Total geral", data.metrics.total),
  ].join("");
  renderBalance(data.balance);
  renderReview(data.review);
  renderGroups(data.groups);
  reportContent.hidden = false;
}

(async () => {
  reportData = await loadReportData();
  if (!reportData) {
    emptyState.hidden = false;
  } else {
    renderReport(reportData);
  }
})();

groupSearch.addEventListener("input", applySearch);
