/**
 * BACKEND - Ordens de Serviço - Ponto dos Fogões
 * Cole este código no Apps Script vinculado (Extensões > Apps Script)
 * à planilha "Ordem de serviço - PDF - NOVA".
 *
 * NÃO altera nenhuma coluna existente da aba principal. Adiciona apenas
 * colunas novas no final (FOTOS, PRAZO GARANTIA) e uma ABA NOVA
 * (HISTORICO_OS) para o histórico de movimentações — criada automaticamente
 * na primeira execução, sem tocar nas abas já existentes.
 */

/* ================= CONFIGURAÇÃO ================= */
// Nome exato da aba com os dados das ordens de serviço.
// Confirme o nome da aba na parte de baixo da planilha (ex: "Ordem de Serviços").
const SHEET_NAME = "Ordem de Serviços";

// Nome da aba nova de histórico (criada automaticamente se não existir).
const HIST_SHEET_NAME = "HISTORICO_OS";

// Cole aqui o ID de uma pasta do Google Drive onde as fotos serão salvas.
// Pegue o ID na URL da pasta: drive.google.com/drive/folders/ESTE_TRECHO_AQUI
const DRIVE_FOLDER_ID = "COLE_O_ID_DA_PASTA_DO_DRIVE_AQUI";

const HEADER_ROW = 1;
const DATA_START_ROW = 2;

/* ================= STATUS / MÁQUINA DE ESTADOS ================= */
// Fluxo principal. Cada status só pode avançar para os status listados aqui
// (evita, por exemplo, pular de ABERTO direto para ENTREGUE).
// CANCELADO e DESCARTE podem ser aplicados a partir de qualquer status não-terminal.
const STATUS_FLOW = {
  "ABERTO": ["EM AVALIAÇÃO"],
  "EM AVALIAÇÃO": ["AGUARDANDO ORÇAMENTO", "AGUARDANDO PEÇA"],
  "AGUARDANDO ORÇAMENTO": ["AGUARDANDO APROVAÇÃO"],
  "AGUARDANDO APROVAÇÃO": ["APROVADO", "NÃO APROVADO"],
  "APROVADO": ["EM EXECUÇÃO"],
  "AGUARDANDO PEÇA": ["EM EXECUÇÃO"],
  "EM EXECUÇÃO": ["FINALIZADO", "AGUARDANDO PEÇA"],
  "FINALIZADO": ["PRONTO PARA RETIRADA"],
  "PRONTO PARA RETIRADA": ["ENTREGUE"],
  "NÃO APROVADO": ["EM AVALIAÇÃO"],
  "ENTREGUE": [],
  "CANCELADO": [],
  "DESCARTE": []
};
const TERMINAL_STATUS = ["ENTREGUE", "CANCELADO", "DESCARTE"];

function isValidTransition(from, to) {
  if (!from || from === to) return true; // criação, ou "alteração" sem mudança real
  if (to === "CANCELADO" || to === "DESCARTE") {
    return TERMINAL_STATUS.indexOf(from) === -1; // sempre permitido, exceto a partir de status já terminal
  }
  return (STATUS_FLOW[from] || []).indexOf(to) !== -1;
}

// Mapeamento de colunas (1-indexed), igual à estrutura atual da planilha.
const COL = {
  OS: 1,
  CLIENTE: 2,
  DATA_ENTRADA: 3,
  TELEFONE: 4,
  SETOR: 5,
  // 6 = coluna em branco (não usar)
  ATENDENTE: 7,
  ENTREGUE_EM: 8,
  // blocos de item: ITEM, GARANTIA, SERVICO, HISTORICO, VALOR
  ITEM_BASE: 9, // item 1 começa na coluna 9 (I)
  ENTRADA_TAXA: 39,
  STATUS: 40,
  TECNICO: 41,
  DATA_APROVACAO: 42,
  VALOR_TOTAL: 43,
  GARANTIAS_ORDEM: 44,
  PREVISAO_ORCAMENTO: 45,
  PREVISAO_MENOR_ORCAMENTO: 46,
  FOTOS: 47, // coluna nova (AU) - criada automaticamente se não existir
  PRAZO_GARANTIA_BASE: 48 // colunas novas (AV..BA), 1 por item: prazo de garantia (ex: "90 dias")
};
const ITEM_BLOCK_SIZE = 5; // ITEM, GARANTIA, SERVICO, HISTORICO, VALOR
const MAX_ITEMS = 6;

function itemCol(itemIndex, offset) {
  // itemIndex: 1..6 ; offset: 0=ITEM,1=GARANTIA,2=SERVICO,3=HISTORICO,4=VALOR
  return COL.ITEM_BASE + (itemIndex - 1) * ITEM_BLOCK_SIZE + offset;
}
function prazoGarantiaCol(itemIndex) {
  // itemIndex: 1..6 - coluna nova no final da planilha, uma por item
  return COL.PRAZO_GARANTIA_BASE + (itemIndex - 1);
}

/* ================= HELPERS ================= */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error("Aba '" + SHEET_NAME + "' não encontrada. Verifique SHEET_NAME no código.");
  return sh;
}

// Cria a aba de histórico automaticamente na primeira vez que for necessária.
// Não mexe em nenhuma aba já existente.
function getHistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HIST_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(HIST_SHEET_NAME);
    sh.getRange(1, 1, 1, 6).setValues([[
      "OS", "DATA/HORA", "USUÁRIO", "EVENTO", "STATUS ANTERIOR", "STATUS NOVO"
    ]]);
    sh.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  return sh;
}

function logHistory(os, usuario, evento, statusAnterior, statusNovo) {
  const sh = getHistSheet();
  sh.appendRow([
    os,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy HH:mm"),
    usuario || "",
    evento || "",
    statusAnterior || "",
    statusNovo || ""
  ]);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getNextOsNumber(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 1;
  const values = sheet.getRange(DATA_START_ROW, COL.OS, lastRow - DATA_START_ROW + 1, 1).getValues();
  let max = 0;
  values.forEach(r => {
    const n = parseInt(r[0], 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function findRowByOs(sheet, osNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return -1;
  const values = sheet.getRange(DATA_START_ROW, COL.OS, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(osNumber).trim()) {
      return DATA_START_ROW + i;
    }
  }
  return -1;
}

function saveImageToDrive(dataUrl, osNumber, index) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const matches = dataUrl.match(/^data:(image\/\w+);base64,(.+)$/);
  if (!matches) return null;
  const mime = matches[1];
  const base64 = matches[2];
  const bytes = Utilities.base64Decode(base64);
  const ext = mime.split('/')[1] || 'jpg';
  const blob = Utilities.newBlob(bytes, mime, "OS_" + osNumber + "_" + index + "_" + Date.now() + "." + ext);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

/* ================= GET (leitura) ================= */
function doGet(e) {
  const action = e.parameter.action;
  const sheet = getSheet();

  if (action === "ping") {
    return jsonOut({ ok: true });
  }

  if (action === "nextos") {
    return jsonOut({ ok: true, nextOs: getNextOsNumber(sheet) });
  }

  if (action === "search") {
    const q = (e.parameter.q || "").trim().toLowerCase();
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return jsonOut({ ok: true, results: [] });
    const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, COL.FOTOS);
    const values = range.getValues();
    const results = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const os = String(row[COL.OS - 1] || "");
      const cliente = String(row[COL.CLIENTE - 1] || "");
      const telefone = String(row[COL.TELEFONE - 1] || "");
      if (!os && !cliente) continue;
      if (os.toLowerCase() === q ||
          cliente.toLowerCase().indexOf(q) !== -1 ||
          telefone.replace(/\D/g,'').indexOf(q.replace(/\D/g,'')) !== -1) {
        results.push({
          os: os,
          cliente: cliente,
          data: row[COL.DATA_ENTRADA - 1],
          setor: row[COL.SETOR - 1],
          status: row[COL.STATUS - 1],
          valorTotal: row[COL.VALOR_TOTAL - 1]
        });
      }
      if (results.length >= 30) break;
    }
    return jsonOut({ ok: true, results: results });
  }

  if (action === "get") {
    const os = e.parameter.os;
    const rowIdx = findRowByOs(sheet, os);
    if (rowIdx === -1) return jsonOut({ ok: false, error: "Ordem não encontrada" });
    const lastCol = prazoGarantiaCol(MAX_ITEMS);
    const row = sheet.getRange(rowIdx, 1, 1, lastCol).getValues()[0];

    const items = [];
    for (let i = 1; i <= MAX_ITEMS; i++) {
      const itemVal = row[itemCol(i, 0) - 1];
      if (!itemVal) continue;
      items.push({
        item: itemVal,
        garantia: row[itemCol(i, 1) - 1],
        servico: row[itemCol(i, 2) - 1],
        historico: row[itemCol(i, 3) - 1],
        valor: row[itemCol(i, 4) - 1],
        prazoGarantia: row[prazoGarantiaCol(i) - 1] || ""
      });
    }

    const fotosRaw = row[COL.FOTOS - 1] || "";
    const fotos = fotosRaw ? String(fotosRaw).split("\n").filter(Boolean) : [];
    const status = row[COL.STATUS - 1] || "ABERTO";

    return jsonOut({
      ok: true,
      data: {
        os: row[COL.OS - 1],
        cliente: row[COL.CLIENTE - 1],
        telefone: row[COL.TELEFONE - 1],
        dataEntrada: row[COL.DATA_ENTRADA - 1],
        setor: row[COL.SETOR - 1],
        atendente: row[COL.ATENDENTE - 1],
        tecnico: row[COL.TECNICO - 1],
        status: status,
        proximosStatus: STATUS_FLOW[status] || [],
        entrada: row[COL.ENTRADA_TAXA - 1],
        valorTotal: row[COL.VALOR_TOTAL - 1],
        items: items,
        fotos: fotos
      }
    });
  }

  if (action === "history") {
    const os = e.parameter.os;
    const hsh = getHistSheet();
    const lastRow = hsh.getLastRow();
    if (lastRow < 2) return jsonOut({ ok: true, events: [] });
    const values = hsh.getRange(2, 1, lastRow - 1, 6).getValues();
    const events = values
      .filter(r => String(r[0]).trim() === String(os).trim())
      .map(r => ({
        dataHora: r[1], usuario: r[2], evento: r[3], statusAnterior: r[4], statusNovo: r[5]
      }));
    return jsonOut({ ok: true, events: events });
  }

  if (action === "dashboard") {
    const lastRow = sheet.getLastRow();
    const counts = {};
    const recent = [];
    if (lastRow >= DATA_START_ROW) {
      const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, COL.FOTOS);
      const values = range.getValues();
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const os = row[COL.OS - 1];
        if (!os) continue;
        const status = String(row[COL.STATUS - 1] || "ABERTO").trim() || "ABERTO";
        counts[status] = (counts[status] || 0) + 1;
        recent.push({
          os: os,
          cliente: row[COL.CLIENTE - 1],
          data: row[COL.DATA_ENTRADA - 1],
          status: status
        });
      }
    }
    recent.reverse(); // mais recentes primeiro (última linha = última criada)
    return jsonOut({ ok: true, counts: counts, recent: recent.slice(0, 10) });
  }

  return jsonOut({ ok: false, error: "Ação inválida" });
}

/* ================= POST (escrita) ================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    if (payload.action === "create") {
      return jsonOut(createOrder(sheet, payload));
    }
    if (payload.action === "update") {
      return jsonOut(updateOrder(sheet, payload));
    }
    return jsonOut({ ok: false, error: "Ação inválida" });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function createOrder(sheet, payload) {
  const osNumber = getNextOsNumber(sheet);
  const lastRow = sheet.getLastRow();
  const newRow = lastRow < DATA_START_ROW ? DATA_START_ROW : lastRow + 1;

  const status = "ABERTO"; // toda O.S. nova começa em ABERTO (ponto de partida da máquina de estados)

  sheet.getRange(newRow, COL.OS).setValue(osNumber);
  sheet.getRange(newRow, COL.CLIENTE).setValue(payload.cliente || "");
  sheet.getRange(newRow, COL.DATA_ENTRADA).setValue(payload.dataEntrada || "");
  sheet.getRange(newRow, COL.TELEFONE).setValue(payload.telefone || "");
  sheet.getRange(newRow, COL.SETOR).setValue(payload.setor || "");
  sheet.getRange(newRow, COL.ATENDENTE).setValue(payload.atendente || "");
  sheet.getRange(newRow, COL.TECNICO).setValue(payload.tecnico || "");
  sheet.getRange(newRow, COL.ENTRADA_TAXA).setValue(payload.entrada || 0);
  sheet.getRange(newRow, COL.STATUS).setValue(status);

  const items = payload.items || [];
  let total = 0;
  items.slice(0, MAX_ITEMS).forEach((it, idx) => {
    const i = idx + 1;
    sheet.getRange(newRow, itemCol(i, 0)).setValue(it.item || "");
    sheet.getRange(newRow, itemCol(i, 1)).setValue(it.garantia || "NÃO");
    sheet.getRange(newRow, itemCol(i, 2)).setValue(it.servico || "");
    sheet.getRange(newRow, itemCol(i, 3)).setValue(it.historico || "");
    sheet.getRange(newRow, itemCol(i, 4)).setValue(it.valor || 0);
    sheet.getRange(newRow, prazoGarantiaCol(i)).setValue(it.garantia === "SIM" ? (it.prazoGarantia || "") : "");
    total += Number(it.valor) || 0;
  });
  sheet.getRange(newRow, COL.VALOR_TOTAL).setValue(total);

  const fotoUrls = [];
  (payload.photos || []).forEach((dataUrl, idx) => {
    const url = saveImageToDrive(dataUrl, osNumber, idx + 1);
    if (url) fotoUrls.push(url);
  });
  if (fotoUrls.length > 0) {
    sheet.getRange(newRow, COL.FOTOS).setValue(fotoUrls.join("\n"));
  }

  logHistory(osNumber, payload.atendente || "", "O.S. criada", "", status);

  return { ok: true, os: osNumber, row: newRow };
}

function updateOrder(sheet, payload) {
  const rowIdx = findRowByOs(sheet, payload.os);
  if (rowIdx === -1) return { ok: false, error: "Ordem não encontrada" };
  const usuario = payload.usuario || "";

  if (payload.status) {
    const statusAtual = sheet.getRange(rowIdx, COL.STATUS).getValue() || "ABERTO";
    if (payload.status !== statusAtual) {
      if (!isValidTransition(statusAtual, payload.status)) {
        return {
          ok: false,
          error: "Transição de status inválida: " + statusAtual + " → " + payload.status
        };
      }
      sheet.getRange(rowIdx, COL.STATUS).setValue(payload.status);
      if (payload.status === "ENTREGUE") {
        sheet.getRange(rowIdx, COL.ENTREGUE_EM).setValue(payload.dataEntrega || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy"));
      }
      logHistory(payload.os, usuario, "Status alterado", statusAtual, payload.status);
    }
  }

  if (payload.newItem && payload.newItem.item) {
    for (let i = 1; i <= MAX_ITEMS; i++) {
      const cell = sheet.getRange(rowIdx, itemCol(i, 0));
      if (!cell.getValue()) {
        sheet.getRange(rowIdx, itemCol(i, 0)).setValue(payload.newItem.item || "");
        sheet.getRange(rowIdx, itemCol(i, 1)).setValue(payload.newItem.garantia || "NÃO");
        sheet.getRange(rowIdx, itemCol(i, 2)).setValue(payload.newItem.servico || "");
        sheet.getRange(rowIdx, itemCol(i, 3)).setValue(payload.newItem.historico || "");
        sheet.getRange(rowIdx, itemCol(i, 4)).setValue(payload.newItem.valor || 0);
        sheet.getRange(rowIdx, prazoGarantiaCol(i)).setValue(
          payload.newItem.garantia === "SIM" ? (payload.newItem.prazoGarantia || "") : ""
        );
        break;
      }
    }
    // recalcula valor total somando os 6 slots
    let total = 0;
    for (let i = 1; i <= MAX_ITEMS; i++) {
      total += Number(sheet.getRange(rowIdx, itemCol(i, 4)).getValue()) || 0;
    }
    sheet.getRange(rowIdx, COL.VALOR_TOTAL).setValue(total);
    logHistory(payload.os, usuario, "Item adicionado: " + payload.newItem.item, "", "");
  }

  if (payload.newPhotos && payload.newPhotos.length > 0) {
    const osNumber = sheet.getRange(rowIdx, COL.OS).getValue();
    const existing = sheet.getRange(rowIdx, COL.FOTOS).getValue();
    const existingUrls = existing ? String(existing).split("\n").filter(Boolean) : [];
    payload.newPhotos.forEach((dataUrl, idx) => {
      const url = saveImageToDrive(dataUrl, osNumber, existingUrls.length + idx + 1);
      if (url) existingUrls.push(url);
    });
    sheet.getRange(rowIdx, COL.FOTOS).setValue(existingUrls.join("\n"));
    logHistory(payload.os, usuario, payload.newPhotos.length + " foto(s) adicionada(s)", "", "");
  }

  return { ok: true, row: rowIdx };
}
