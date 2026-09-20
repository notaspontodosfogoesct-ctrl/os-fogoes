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

// Nome da aba nova de registro de mensagens WhatsApp enviadas (criada automaticamente).
const WPP_SHEET_NAME = "WHATSAPP_ENVIOS";

// Nome da aba de clientes já cadastrados (só leitura — o app nunca escreve aqui).
const CADASTRO_SHEET_NAME = "Cadastro";

// Cole aqui o ID de uma pasta do Google Drive onde as fotos serão salvas.
// Pegue o ID na URL da pasta: drive.google.com/drive/folders/ESTE_TRECHO_AQUI
const DRIVE_FOLDER_ID = "COLE_O_ID_DA_PASTA_DO_DRIVE_AQUI";

const HEADER_ROW = 1;
const DATA_START_ROW = 2;

/* ================= STATUS / MÁQUINA DE ESTADOS ================= */
// Fluxo real da empresa. Cada status só pode avançar para os status listados aqui.
// CANCELADO pode ser aplicado a partir de qualquer status não-terminal (botão dedicado).
const STATUS_FLOW = {
  "ABERTA": ["AGUARDA APROVAÇÃO"],
  "AGUARDA APROVAÇÃO": ["APROVADO", "NÃO APROVADO"],
  "APROVADO": ["PRONTO", "AGUARDANDO PEÇA"],
  "AGUARDANDO PEÇA": ["PRONTO"],
  "NÃO APROVADO": ["RETIRADO"],
  "PRONTO": ["RETIRADO"],
  "RETIRADO": [],
  "CANCELADO": []
};
const TERMINAL_STATUS = ["RETIRADO", "CANCELADO"];

// ===== Prazos (SLA) por setor, em dias corridos =====
// Orçamento: contado a partir da Data de Entrada, enquanto a O.S. ainda não foi aprovada/reprovada.
const SLA_ORCAMENTO_DIAS = { "PANELAS": 5, "LIQUIDIFICADORES": 15, "ELÉTRICOS": 15 };
const SLA_ORCAMENTO_PADRAO = 15;
// Finalização: contado a partir da Data de Aprovação, só enquanto está em APROVADO
// (pausa automaticamente se o status for AGUARDANDO PEÇA).
const SLA_FINALIZACAO_DIAS = { "PANELAS": 5, "LIQUIDIFICADORES": 10, "ELÉTRICOS": 10 };
const SLA_FINALIZACAO_PADRAO = 10;

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
  ATENDENTE: 6,
  ENTREGUE_EM: 7,
  // blocos de item: ITEM, GARANTIA, SERVICO, HISTORICO, VALOR
  ITEM_BASE: 8, // item 1 começa na coluna 8 (H)
  ENTRADA_TAXA: 38,
  STATUS: 39,
  TECNICO: 40,
  DATA_APROVACAO: 41,
  VALOR_TOTAL: 42,
  GARANTIAS_ORDEM: 43,
  PREVISAO_ORCAMENTO: 44,
  PREVISAO_MENOR_ORCAMENTO: 45,
  FOTOS: 46, // coluna nova (AT) - criada automaticamente se não existir
  PRAZO_GARANTIA_BASE: 47 // colunas novas (AU..AZ), 1 por item: prazo de garantia (ex: "90 dias")
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

// Converte número de coluna (1-indexed) para letra (ex: 1->A, 28->AB).
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// A planilha já usa uma FÓRMULA nessa célula (soma dos 6 itens). Nunca
// escrevemos um número fixo aqui — sempre a fórmula, para que continue
// recalculando sozinha se algum valor de item for editado depois.
function valorTotalFormula(row) {
  const letras = [1, 2, 3, 4, 5, 6].map(i => columnToLetter(itemCol(i, 4)));
  return "=" + letras.slice().reverse().map(l => l + row).join("+");
}

// Garante o telefone no mesmo padrão já usado na planilha: só dígitos,
// com o 55 (Brasil) na frente quando faltar.
function normalizePhone(v) {
  let digits = String(v || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.length <= 11 && !digits.startsWith("55")) {
    digits = "55" + digits;
  }
  return digits;
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

// Registro de mensagens WhatsApp enviadas pelo app (aba nova, criada sozinha).
function getWppSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(WPP_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(WPP_SHEET_NAME);
    sh.getRange(1, 1, 1, 6).setValues([[
      "OS", "DATA/HORA", "TIPO", "MENSAGEM ENVIADA", "TELEFONE", "USUÁRIO"
    ]]);
    sh.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  return sh;
}
function logWhatsapp(os, tipo, mensagem, telefone, usuario) {
  const sh = getWppSheet();
  sh.appendRow([
    os,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy HH:mm"),
    tipo || "",
    mensagem || "",
    telefone || "",
    usuario || ""
  ]);
}

// Compara só os últimos 9 dígitos do telefone, pra não falhar por causa de
// +55, DDD faltando, 9º dígito etc.
function telefoneChave(v) {
  const digits = String(v || "").replace(/\D/g, "");
  return digits.slice(-9);
}

// Busca um cliente na aba Cadastro pelo telefone. SÓ LEITURA — nunca escreve
// nem cria linha nova aqui, pra não gerar cadastro duplicado.
function buscarClientePorTelefone(telefone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CADASTRO_SHEET_NAME);
  if (!sh) return null;
  const chave = telefoneChave(telefone);
  if (!chave) return null;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  // Colunas do Cadastro: A DATA, B TELEFONE, C NOME, D ENDEREÇO, E E-MAIL
  const values = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < values.length; i++) {
    if (telefoneChave(values[i][1]) === chave) {
      return {
        nome: values[i][2] || "",
        endereco: values[i][3] || "",
        email: values[i][4] || ""
      };
    }
  }
  return null;
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

// Converte valor de célula (Date real ou texto "dd/MM/yyyy") num objeto Date.
function toDateSafe(v) {
  if (v instanceof Date) return v;
  if (typeof v === "string" && v.trim()) {
    const p = v.trim().split("/");
    if (p.length === 3) {
      return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
    }
  }
  return null;
}
function diffDias(prazo, hoje) {
  const p = new Date(prazo); p.setHours(0,0,0,0);
  const h = new Date(hoje); h.setHours(0,0,0,0);
  return Math.round((p.getTime() - h.getTime()) / 86400000);
}
// vencido = já passou do prazo | proximo = 2 dias ou menos | null = tranquilo (não exibe alerta)
function nivelAlerta(diasRestantes) {
  if (diasRestantes < 0) return "vencido";
  if (diasRestantes <= 2) return "proximo";
  return null;
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

  if (action === "buscar_cliente") {
    const telefone = e.parameter.telefone || "";
    const cliente = buscarClientePorTelefone(telefone);
    if (!cliente) return jsonOut({ ok: true, encontrado: false });
    return jsonOut({ ok: true, encontrado: true, nome: cliente.nome, endereco: cliente.endereco, email: cliente.email });
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
    const status = row[COL.STATUS - 1] || "ABERTA";

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
        dataAprovacao: row[COL.DATA_APROVACAO - 1],
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
    const ativos = [];
    const hoje = new Date();
    hoje.setHours(0,0,0,0);

    if (lastRow >= DATA_START_ROW) {
      const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, COL.FOTOS);
      const values = range.getValues();
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const os = row[COL.OS - 1];
        if (!os) continue;
        const status = String(row[COL.STATUS - 1] || "ABERTA").trim() || "ABERTA";
        const setor = String(row[COL.SETOR - 1] || "").trim().toUpperCase();
        counts[status] = (counts[status] || 0) + 1;

        if (TERMINAL_STATUS.indexOf(status) !== -1) continue; // não entram na lista de acompanhamento

        const item = {
          os: os,
          cliente: row[COL.CLIENTE - 1],
          setor: row[COL.SETOR - 1],
          atendente: row[COL.ATENDENTE - 1],
          status: status,
          dataEntrada: row[COL.DATA_ENTRADA - 1],
          alerta: null
        };

        const dataEntrada = toDateSafe(row[COL.DATA_ENTRADA - 1]);
        const dataAprovacao = toDateSafe(row[COL.DATA_APROVACAO - 1]);

        if ((status === "ABERTA" || status === "AGUARDA APROVAÇÃO") && dataEntrada) {
          const dias = SLA_ORCAMENTO_DIAS[setor] || SLA_ORCAMENTO_PADRAO;
          const prazo = new Date(dataEntrada);
          prazo.setDate(prazo.getDate() + dias);
          const restam = diffDias(prazo, hoje);
          item.alerta = { tipo: "orçamento", diasRestantes: restam, nivel: nivelAlerta(restam) };
        } else if (status === "APROVADO" && dataAprovacao) {
          const dias = SLA_FINALIZACAO_DIAS[setor] || SLA_FINALIZACAO_PADRAO;
          const prazo = new Date(dataAprovacao);
          prazo.setDate(prazo.getDate() + dias);
          const restam = diffDias(prazo, hoje);
          item.alerta = { tipo: "finalização", diasRestantes: restam, nivel: nivelAlerta(restam) };
        }
        // AGUARDANDO PEÇA: prazo pausado de propósito (sem item.alerta)

        ativos.push(item);
      }
    }
    ativos.reverse(); // mais recentes primeiro
    return jsonOut({ ok: true, counts: counts, ativos: ativos });
  }

  // ===== ROTAS DA TELA RESTRITA DO TÉCNICO (sem telefone / sem valores) =====

  if (action === "list_setor") {
    // Lista as O.S. do setor do técnico, sem telefone nem valores.
    const setor = (e.parameter.setor || "").trim().toUpperCase();
    const lastRow = sheet.getLastRow();
    const list = [];
    if (lastRow >= DATA_START_ROW) {
      const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, COL.FOTOS);
      const values = range.getValues();
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const os = row[COL.OS - 1];
        if (!os) continue;
        const rowSetor = String(row[COL.SETOR - 1] || "").trim().toUpperCase();
        if (rowSetor !== setor) continue;
        const status = String(row[COL.STATUS - 1] || "ABERTA").trim() || "ABERTA";
        if (TERMINAL_STATUS.indexOf(status) !== -1) continue; // não mostra O.S. já encerradas
        list.push({ os: os, cliente: row[COL.CLIENTE - 1], data: row[COL.DATA_ENTRADA - 1], status: status });
      }
    }
    list.reverse();
    return jsonOut({ ok: true, list: list });
  }

  if (action === "get_tech") {
    // Igual ao "get", mas nunca inclui telefone, valores dos itens nem valor total.
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
        index: i,
        item: itemVal,
        garantia: row[itemCol(i, 1) - 1],
        servico: row[itemCol(i, 2) - 1],
        historico: row[itemCol(i, 3) - 1],
        prazoGarantia: row[prazoGarantiaCol(i) - 1] || ""
      });
    }
    const status = row[COL.STATUS - 1] || "ABERTA";

    return jsonOut({
      ok: true,
      data: {
        os: row[COL.OS - 1],
        cliente: row[COL.CLIENTE - 1],
        dataEntrada: row[COL.DATA_ENTRADA - 1],
        setor: row[COL.SETOR - 1],
        status: status,
        podeFinalizar: status === "APROVADO",
        items: items
      }
    });
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
    if (payload.action === "update_tech") {
      return jsonOut(updateOrderTech(sheet, payload));
    }
    if (payload.action === "log_whatsapp") {
      logWhatsapp(payload.os, payload.tipo, payload.mensagem, payload.telefone, payload.usuario);
      logHistory(payload.os, payload.usuario, "Mensagem WhatsApp enviada (" + (payload.tipo || "") + ")", "", "");
      return jsonOut({ ok: true });
    }
    return jsonOut({ ok: false, error: "Ação inválida" });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Verifica se as últimas 3 linhas já têm exatamente o mesmo cliente+telefone+setor+atendente
// — sinal de clique duplicado no Salvar, não de um cliente novo trazendo outro aparelho.
function isDuplicadoRecente(sheet, payload) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return false;
  const qtd = Math.min(3, lastRow - DATA_START_ROW + 1);
  const startRow = lastRow - qtd + 1;
  const values = sheet.getRange(startRow, 1, qtd, COL.ATENDENTE).getValues();
  const cli = String(payload.cliente || "").trim().toUpperCase();
  const tel = String(payload.telefone || "").replace(/\D/g, "");
  const set = String(payload.setor || "").trim().toUpperCase();
  const ate = String(payload.atendente || "").trim().toUpperCase();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (
      String(row[COL.CLIENTE - 1] || "").trim().toUpperCase() === cli &&
      String(row[COL.TELEFONE - 1] || "").replace(/\D/g, "") === tel &&
      String(row[COL.SETOR - 1] || "").trim().toUpperCase() === set &&
      String(row[COL.ATENDENTE - 1] || "").trim().toUpperCase() === ate
    ) {
      return row[COL.OS - 1]; // devolve o número da O.S. já existente
    }
  }
  return false;
}

function createOrder(sheet, payload) {
  const duplicada = isDuplicadoRecente(sheet, payload);
  if (duplicada) {
    return {
      ok: false,
      error: "Essa O.S. já foi criada como nº " + duplicada + " (mesmo cliente, telefone, setor e atendente). Se for outro aparelho do mesmo cliente, mude o setor ou aguarde antes de tentar de novo."
    };
  }

  const osNumber = getNextOsNumber(sheet);
  const lastRow = sheet.getLastRow();
  const newRow = lastRow < DATA_START_ROW ? DATA_START_ROW : lastRow + 1;

  const status = "ABERTA"; // toda O.S. nova começa em ABERTA (ponto de partida da máquina de estados)

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
  items.slice(0, MAX_ITEMS).forEach((it, idx) => {
    const i = idx + 1;
    sheet.getRange(newRow, itemCol(i, 0)).setValue(it.item || "");
    sheet.getRange(newRow, itemCol(i, 1)).setValue(it.garantia || "NÃO");
    sheet.getRange(newRow, itemCol(i, 2)).setValue(it.servico || "");
    sheet.getRange(newRow, itemCol(i, 3)).setValue(it.historico || "");
    sheet.getRange(newRow, itemCol(i, 4)).setValue(it.valor || 0);
    sheet.getRange(newRow, prazoGarantiaCol(i)).setValue(it.garantia === "SIM" ? (it.prazoGarantia || "") : "");
  });
  // VALOR TOTAL não é escrito aqui: a própria planilha já calcula essa coluna
  // por fórmula (soma dos 6 itens). Escrever um número por cima destruiria a fórmula.

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

  // ===== Edição dos dados cadastrais (cliente, telefone, setor, atendente) =====
  if (payload.cadastro) {
    const c = payload.cadastro;
    const mudancas = [];
    if (typeof c.cliente === "string" && c.cliente.trim()) {
      sheet.getRange(rowIdx, COL.CLIENTE).setValue(c.cliente.trim());
      mudancas.push("cliente");
    }
    if (typeof c.telefone === "string" && c.telefone.trim()) {
      sheet.getRange(rowIdx, COL.TELEFONE).setValue(c.telefone.trim());
      mudancas.push("telefone");
    }
    if (typeof c.setor === "string" && c.setor.trim()) {
      sheet.getRange(rowIdx, COL.SETOR).setValue(c.setor.trim());
      mudancas.push("setor");
    }
    if (typeof c.atendente === "string" && c.atendente.trim()) {
      sheet.getRange(rowIdx, COL.ATENDENTE).setValue(c.atendente.trim());
      mudancas.push("atendente");
    }
    if (typeof c.tecnico === "string" && c.tecnico.trim()) {
      sheet.getRange(rowIdx, COL.TECNICO).setValue(c.tecnico.trim());
      mudancas.push("responsável técnico");
    }
    if (c.entrada !== undefined && c.entrada !== null && c.entrada !== "") {
      sheet.getRange(rowIdx, COL.ENTRADA_TAXA).setValue(Number(c.entrada) || 0);
      mudancas.push("entrada/taxa");
    }
    if (mudancas.length > 0) {
      logHistory(payload.os, usuario, "Dados cadastrais editados (" + mudancas.join(", ") + ")", "", "");
    }
  }

  if (payload.status) {
    const statusAtual = sheet.getRange(rowIdx, COL.STATUS).getValue() || "ABERTA";
    if (payload.status !== statusAtual) {
      if (!isValidTransition(statusAtual, payload.status)) {
        return {
          ok: false,
          error: "Transição de status inválida: " + statusAtual + " → " + payload.status
        };
      }
      sheet.getRange(rowIdx, COL.STATUS).setValue(payload.status);
      if (payload.status === "APROVADO") {
        sheet.getRange(rowIdx, COL.DATA_APROVACAO).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy"));
      }
      if (payload.status === "RETIRADO") {
        sheet.getRange(rowIdx, COL.ENTREGUE_EM).setValue(payload.dataEntrega || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy"));
      }
      const motivo = payload.status === "CANCELADO" && payload.motivo ? (" — Motivo: " + payload.motivo) : "";
      logHistory(payload.os, usuario, "Status alterado" + motivo, statusAtual, payload.status);
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
    // VALOR TOTAL não é recalculado/escrito aqui: a planilha já soma os itens
    // por fórmula própria (coluna AP). Mexer nela apagaria a fórmula.
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

/* =========================================================
 * ATUALIZAÇÃO PELO TÉCNICO (acesso restrito)
 * Só altera ITEM / SERVIÇO / HISTÓRICO dos itens informados
 * (nunca GARANTIA nem VALOR), e só permite mudar o status
 * de APROVADO para PRONTO — nada além disso.
 * ========================================================= */
function updateOrderTech(sheet, payload) {
  const rowIdx = findRowByOs(sheet, payload.os);
  if (rowIdx === -1) return { ok: false, error: "Ordem não encontrada" };
  const usuario = payload.usuario || "(técnico não identificado)";

  const itens = payload.items || [];
  itens.forEach(it => {
    const i = parseInt(it.index, 10);
    if (!i || i < 1 || i > MAX_ITEMS) return;
    // só atualiza se já existir um item nesse slot (técnico edita, não cria itens novos)
    if (!sheet.getRange(rowIdx, itemCol(i, 0)).getValue()) return;
    if (typeof it.item === "string") sheet.getRange(rowIdx, itemCol(i, 0)).setValue(it.item);
    if (typeof it.servico === "string") sheet.getRange(rowIdx, itemCol(i, 2)).setValue(it.servico);
    if (typeof it.historico === "string") sheet.getRange(rowIdx, itemCol(i, 3)).setValue(it.historico);
  });
  if (itens.length > 0) {
    logHistory(payload.os, usuario, "Itens atualizados pelo técnico", "", "");
  }

  if (payload.finalizar) {
    const statusAtual = sheet.getRange(rowIdx, COL.STATUS).getValue() || "";
    if (statusAtual !== "APROVADO") {
      return { ok: false, error: "Só é possível finalizar uma O.S. que esteja em APROVADO (status atual: " + statusAtual + ")" };
    }
    sheet.getRange(rowIdx, COL.STATUS).setValue("PRONTO");
    logHistory(payload.os, usuario, "Status alterado (técnico)", statusAtual, "PRONTO");
  }

  return { ok: true, row: rowIdx };
}

/* =========================================================
 * VERIFICAÇÃO — RODE ESTA FUNÇÃO DEPOIS DE COLAR O CÓDIGO
 * Mostra no "Registro de execução" exatamente quais colunas
 * este código vai usar. Confira se bate com o esperado abaixo
 * ANTES de criar qualquer O.S. de teste.
 * ========================================================= */
function verificarMapeamento() {
  Logger.log("=== MAPEAMENTO ATUAL (deveria ser sempre igual a este) ===");
  Logger.log("ATENDENTE deve ser 6 → é: " + COL.ATENDENTE);
  Logger.log("ITEM_BASE deve ser 8 → é: " + COL.ITEM_BASE);
  Logger.log("ENTRADA_TAXA deve ser 38 → é: " + COL.ENTRADA_TAXA);
  Logger.log("STATUS deve ser 39 → é: " + COL.STATUS);
  Logger.log("TECNICO deve ser 40 → é: " + COL.TECNICO);
  Logger.log("DATA_APROVACAO deve ser 41 → é: " + COL.DATA_APROVACAO);
  Logger.log("VALOR_TOTAL deve ser 42 → é: " + COL.VALOR_TOTAL);
  Logger.log("FOTOS deve ser 46 → é: " + COL.FOTOS);
  Logger.log("Status inicial de uma O.S. nova: ABERTA");
  Logger.log("Fluxo de status: " + JSON.stringify(STATUS_FLOW));
  Logger.log("Prazo ORÇAMENTO (dias): " + JSON.stringify(SLA_ORCAMENTO_DIAS) + " padrão=" + SLA_ORCAMENTO_PADRAO);
  Logger.log("Prazo FINALIZAÇÃO (dias): " + JSON.stringify(SLA_FINALIZACAO_DIAS) + " padrão=" + SLA_FINALIZACAO_PADRAO);
  Logger.log("=== Se algum número acima não bater, o código não foi colado por inteiro. ===");
}

/* =========================================================
 * FUNÇÃO DE LIMPEZA — RODAR SÓ UMA VEZ, MANUALMENTE
 * Remove linhas de sobra (sem número de O.S.) que ficaram
 * depois da última ordem real. Não mexe em nenhuma linha
 * que tenha um número de O.S. válido na coluna A.
 * ========================================================= */
function limparLinhasSobra() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  let lastRealRow = DATA_START_ROW - 1;

  for (let r = DATA_START_ROW; r <= lastRow; r++) {
    const v = sheet.getRange(r, COL.OS).getValue();
    if (v !== '' && v !== null) lastRealRow = r;
  }

  if (lastRealRow < lastRow) {
    const qtd = lastRow - lastRealRow;
    sheet.deleteRows(lastRealRow + 1, qtd);
    Logger.log('Removidas ' + qtd + ' linha(s) de sobra, da linha ' + (lastRealRow + 1) + ' até ' + lastRow + '.');
  } else {
    Logger.log('Nenhuma linha de sobra encontrada — nada foi alterado.');
  }
}
