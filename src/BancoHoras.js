// =========================================================================
// BANCO DE HORAS — importação do relatório do RH, saldo em horas/dias,
// uso em folgas/emendas, lembrete ao gestor e envio do saldo ao time.
// -------------------------------------------------------------------------
// Fonte: exportação do Looker "Banco de Horas TOTVS e Regras_Saldo Banco de
// Horas_Tabela dinâmica" (PARTICIPANTE | COMPENSAR ATÉ | BANCO DE HORAS em
// segundos). Cada importação substitui a foto anterior na aba "Banco de Horas".
// Uso: folgas com benefício "BANCO DE HORAS" registradas DEPOIS da última
// importação são descontadas (o próximo relatório do RH já virá com o desconto).
// =========================================================================

const BH_CONFIG = {
  aba: "Banco de Horas",
  tituloExportacao: "Saldo Banco de Horas",
  assuntoEmailRH: "Dashboard de Banco de Horas",
  jornadaPadrao: 8,
  diasAlertaVencimento: 30
};

function jornadaHoras_() {
  const v = parseFloat(PropertiesService.getScriptProperties().getProperty("BH_JORNADA"));
  return v > 0 ? v : BH_CONFIG.jornadaPadrao;
}

function fmtHoras_(h) {
  const neg = h < 0, m = Math.round(Math.abs(h) * 60);
  return (neg ? "-" : "") + Math.floor(m / 60) + "h" + String(m % 60).padStart(2, "0");
}

// "31 de jan. de 2027", "31/01/2027", Date
function parseDataBH_(v) {
  const d = parseData_(v);
  if (d) return d;
  const m = String(v || "").toLowerCase().match(/(\d{1,2})\s*de\s*([a-zç]+)\.?\s*de\s*(\d{4})/);
  if (!m) return null;
  const meses = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  const i = meses.indexOf(m[2].slice(0, 3));
  return i < 0 ? null : new Date(+m[3], i, +m[1]);
}

// ---------- Localizar a exportação ----------
function exportacaoBancoHorasMaisRecente_() {
  const q = 'title contains "' + BH_CONFIG.tituloExportacao + '" and mimeType = "application/vnd.google-apps.spreadsheet" and trashed = false and \'me\' in owners';
  const it = DriveApp.searchFiles(q);
  let melhor = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!melhor || f.getDateCreated() > melhor.getDateCreated()) melhor = f;
  }
  return melhor ? { id: melhor.getId(), nome: melhor.getName(), criada: Utilities.formatDate(melhor.getDateCreated(), tz_(), "dd/MM/yyyy HH:mm") } : null;
}

// ---------- Importar ----------
function importarBancoHoras(idOpcional) {
  let origem;
  if (idOpcional && String(idOpcional).trim()) origem = { id: String(idOpcional).trim(), nome: "ID informado", criada: "" };
  else {
    origem = exportacaoBancoHorasMaisRecente_();
    if (!origem) throw new Error("Nenhuma planilha '" + BH_CONFIG.tituloExportacao + "' sua encontrada no Drive. Exporte a tabela do relatório de banco de horas (⋮ → Exportar → Google Planilhas).");
  }
  const ssRH = SpreadsheetApp.openById(origem.id);
  const abaRH = ssRH.getSheets()[0];
  if (abaRH.getLastRow() < 2) throw new Error("A exportação de banco de horas está vazia.");

  const cab = abaRH.getRange(1, 1, 1, abaRH.getLastColumn()).getValues()[0].map(c => normalizar_(c));
  const col = (n, p) => { const i = cab.indexOf(normalizar_(n)); return i >= 0 ? i : p; };
  const C = { nome: col("PARTICIPANTE", 0), ate: col("COMPENSAR ATE", 1), seg: col("BANCO DE HORAS", 2) };

  const doTime = nomesDoTime_();
  const agora = new Date();
  const linhas = [], fora = {};
  abaRH.getRange(2, 1, abaRH.getLastRow() - 1, abaRH.getLastColumn()).getValues().forEach(r => {
    const nome = String(r[C.nome] || "").trim().toUpperCase();
    const ate = parseDataBH_(r[C.ate]);
    const seg = parseFloat(String(r[C.seg]).replace(",", "."));
    if (!nome || !ate || isNaN(seg)) return;
    if (doTime && !doTime[normalizar_(nome)]) { fora[nome] = 1; return; }
    const horas = Math.round(seg / 36) / 100; // segundos -> horas (2 casas)
    linhas.push([nome, ate, horas, fmtHoras_(horas), agora, origem.nome]);
  });
  if (!linhas.length) throw new Error("Nenhuma linha de banco de horas de pessoas do time foi encontrada na exportação.");

  const ss = planilha_();
  const aba = ss.getSheetByName(BH_CONFIG.aba) || ss.insertSheet(BH_CONFIG.aba);
  if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, Math.max(6, aba.getLastColumn())).clearContent();
  aba.getRange(1, 1, 1, 6).setValues([["Colaborador", "Compensar até", "Saldo (horas)", "Saldo (h:mm)", "Importado em", "Origem"]])
    .setFontWeight("bold").setBackground("#1f4e78").setFontColor("#ffffff");
  aba.setFrozenRows(1);
  linhas.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]);
  aba.getRange(2, 1, linhas.length, 6).setValues(linhas);
  aba.getRange(2, 2, linhas.length, 1).setNumberFormat("dd/mm/yyyy");
  aba.getRange(2, 5, linhas.length, 1).setNumberFormat("dd/mm/yyyy hh:mm");
  SpreadsheetApp.flush();

  const props = PropertiesService.getScriptProperties();
  props.setProperty("BH_ULTIMA_IMPORTACAO", agora.toISOString());
  props.setProperty("BH_ULTIMO_ID", origem.id);

  const pessoas = {};
  linhas.forEach(l => pessoas[l[0]] = 1);
  return "Banco de horas importado: " + Object.keys(pessoas).length + " pessoa(s), " + linhas.length + " linha(s)" +
    (origem.criada ? " (" + origem.nome + " de " + origem.criada + ")" : "") + "." +
    (Object.keys(fora).length ? " 🚫 Ignoradas " + Object.keys(fora).length + " pessoa(s) fora do time: " + Object.keys(fora).map(primeiroNome_).join(", ") + "." : "");
}

// ---------- Saldo ----------
// Retorna o saldo disponível de uma pessoa (ou de todos, se nome vazio).
function saldosBancoHoras_() {
  const ss = planilha_();
  const jornada = jornadaHoras_();
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const res = {};
  const base = n => (res[n] = res[n] || { nome: n, faixas: [], totalValido: 0, vencidas: 0, usadoHoras: 0, usos: [] });

  const aba = ss.getSheetByName(BH_CONFIG.aba);
  if (aba && aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues().forEach(r => {
      const n = normalizar_(r[0]); const ate = parseDataBH_(r[1]); const h = parseFloat(r[2]) || 0;
      if (!n || !ate) return;
      const p = base(n);
      const vencida = ate < hoje;
      p.faixas.push({ ate: iso_(ate), horas: h, vencida: vencida });
      if (vencida) p.vencidas += h; else p.totalValido += h;
    });
  }

  // Usos registrados depois da última importação
  const ultIso = PropertiesService.getScriptProperties().getProperty("BH_ULTIMA_IMPORTACAO");
  const ult = ultIso ? new Date(ultIso) : null;
  const folgas = ss.getSheetByName("Folgas e Emendas - 2026");
  if (folgas && folgas.getLastRow() > 1) {
    folgas.getRange(2, 1, folgas.getLastRow() - 1, 6).getValues().forEach(r => {
      if (normalizar_(r[1]).indexOf("BANCO") === -1) return;
      const reg = r[5] instanceof Date ? r[5] : parseData_(r[5]);
      if (ult && reg && reg < ult) return;
      const n = normalizar_(r[0]); if (!n) return;
      const dias = Math.max(1, parseInt(r[3], 10) || 1);
      const p = base(n);
      p.usadoHoras += dias * jornada;
      p.usos.push({ data: iso_(parseData_(r[2])), dias: dias });
    });
  }

  Object.keys(res).forEach(n => {
    const p = res[n];
    p.disponivel = Math.max(0, Math.round((p.totalValido - p.usadoHoras) * 100) / 100);
    p.dias = Math.floor(p.disponivel / jornada);
    p.disponivelFmt = fmtHoras_(p.disponivel);
    p.faixas.sort((a, b) => a.ate < b.ate ? -1 : 1);
    // Próxima faixa positiva que vence em breve
    const prox = p.faixas.filter(f => !f.vencida && f.horas > 0)[0];
    if (prox) {
      const dd = Math.round((parseData_(prox.ate) - hoje) / 86400000);
      if (dd <= BH_CONFIG.diasAlertaVencimento) p.vencendo = { ate: prox.ate, horas: prox.horas, fmt: fmtHoras_(prox.horas), emDias: dd };
    }
  });
  return { porNome: res, jornada: jornada, ultimaImportacao: ultIso || "" };
}

function saldoBancoHorasDe_(nome) {
  const s = saldosBancoHoras_();
  return s.porNome[normalizar_(nome)] || { disponivel: 0, dias: 0, disponivelFmt: "0h00", faixas: [], usos: [] };
}

// ---------- Dashboard ----------
function obterBancoHoras() {
  vincularPlanilha_();
  const s = saldosBancoHoras_();
  let nomes = [];
  try { nomes = buscarNomesTime().map(n => String(n).trim().toUpperCase()); } catch (e) {}
  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  let ex = null;
  try { ex = exportacaoBancoHorasMaisRecente_(); } catch (e) {}
  const props = PropertiesService.getScriptProperties();
  return {
    jornada: s.jornada,
    ultimaImportacao: s.ultimaImportacao ? Utilities.formatDate(new Date(s.ultimaImportacao), tz_(), "dd/MM/yyyy HH:mm") : "",
    exportacao: ex,
    exportacaoNova: !!ex && props.getProperty("BH_ULTIMO_ID") !== ex.id,
    emailRH: (function() { try { return ultimoEmailBancoHoras_(); } catch (e) { return null; } })(),
    ehGestor: !!usuario && String(usuario).toLowerCase() === String(emailLider_()).toLowerCase(),
    pessoas: nomes.map(n => {
      const p = s.porNome[normalizar_(n)];
      return p ? { nome: n, disponivel: p.disponivel, fmt: p.disponivelFmt, dias: p.dias, usado: p.usadoHoras, vencendo: p.vencendo || null,
        faixas: p.faixas.map(f => ({ ate: f.ate, fmt: fmtHoras_(f.horas), horas: f.horas, vencida: f.vencida })) }
        : { nome: n, disponivel: 0, fmt: "–", dias: 0, usado: 0, vencendo: null, faixas: [] };
    })
  };
}

function salvarJornadaBH(horas) {
  if (emailLider_()) validarGestor_();
  const h = parseFloat(horas);
  if (!(h > 0 && h <= 12)) throw new Error("Informe a jornada em horas (ex.: 8).");
  PropertiesService.getScriptProperties().setProperty("BH_JORNADA", String(h));
  return "Jornada salva: " + h + "h por dia.";
}

// ---------- E-mail do RH e lembrete ----------
function ultimoEmailBancoHoras_() {
  const th = GmailApp.search('subject:"' + BH_CONFIG.assuntoEmailRH + '" newer_than:60d', 0, 1);
  if (!th.length) return null;
  const d = th[0].getLastMessageDate();
  return { id: th[0].getId(), assunto: th[0].getFirstMessageSubject(), data: Utilities.formatDate(d, tz_(), "dd/MM/yyyy"), ms: d.getTime() };
}

// Chamado de hora em hora pela rotina automática:
//  - se apareceu uma exportação nova de banco de horas sua, importa;
//  - se chegou o e-mail do RH e ainda não houve importação depois dele, lembra o gestor (uma vez por e-mail).
function verificarBancoHoras_() {
  const props = PropertiesService.getScriptProperties();
  try {
    const ex = exportacaoBancoHorasMaisRecente_();
    if (ex && props.getProperty("BH_ULTIMO_ID") !== ex.id) {
      const r = importarBancoHoras();
      registrarUltimaSinc_("Banco de horas (automático)", r);
    }
  } catch (e) { Logger.log("Banco de horas: " + e.message); }

  try {
    const em = ultimoEmailBancoHoras_();
    if (!em || !emailLider_()) return;
    const ult = props.getProperty("BH_ULTIMA_IMPORTACAO");
    if (ult && new Date(ult).getTime() >= em.ms) return;
    if (props.getProperty("BH_LEMBRETE_ENVIADO") === em.id) return;
    enviarEmail_(emailLider_(), "⏰ Banco de horas: hora de exportar e importar (" + em.assunto + ")", "Lembrete do banco de horas",
      "<p>O RH enviou o <b>" + esc_(em.assunto) + "</b> em <b>" + em.data + "</b>, e o controle de férias ainda não recebeu esse saldo.</p>" +
      "<ol><li>Abra o relatório de banco de horas pelo link do e-mail do RH.</li>" +
      "<li>Na tabela <b>Banco de horas por TOTVER</b>, clique em <b>⋮ → Exportar → Google Planilhas</b>.</li>" +
      "<li>Pronto: em até 1 hora o script importa sozinho. Ou abra o dashboard → <b>Banco de horas → Importar agora</b>.</li></ol>");
    props.setProperty("BH_LEMBRETE_ENVIADO", em.id);
  } catch (e) { Logger.log("Lembrete banco de horas: " + e.message); }
}

// ---------- Enviar saldo ao time ----------
function enviarSaldosBancoHoras() {
  if (emailLider_()) validarGestor_();
  const s = saldosBancoHoras_();
  let enviados = 0;
  const semEmail = [];
  buscarNomesTime().forEach(n => {
    const p = s.porNome[normalizar_(n)];
    if (!p || !p.faixas.length) return;
    const email = obterEmail_(n);
    if (!email) { semEmail.push(primeiroNome_(n)); return; }
    const linhas = p.faixas.map(f =>
      "<tr><td style='padding:6px 8px;border-bottom:1px solid #eee'>" + fmt_(parseData_(f.ate)) + (f.vencida ? " <span style='color:#999'>(vencido)</span>" : "") +
      "</td><td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right'><b>" + fmtHoras_(f.horas) + "</b></td></tr>").join("");
    enviarEmail_(email, "⏱️ Seu banco de horas - " + nomeSquad_(), "Seu banco de horas",
      "<p>Olá, <b>" + esc_(primeiroNome_(n)) + "</b>!</p>" +
      "<p>Saldo disponível: <b>" + p.disponivelFmt + "</b>, o que dá <b>" + p.dias + " dia(s)</b> de folga (jornada de " + s.jornada + "h)." +
      (p.usadoHoras ? " Já descontamos " + fmtHoras_(p.usadoHoras) + " usadas desde o último relatório do RH." : "") + "</p>" +
      (p.vencendo ? "<p style='color:#b42328'>⚠️ <b>" + p.vencendo.fmt + "</b> precisam ser compensadas até <b>" + fmt_(parseData_(p.vencendo.ate)) + "</b>.</p>" : "") +
      "<table style='border-collapse:collapse;font-size:13px;margin:10px 0'><tr><th style='text-align:left;padding:6px 8px;background:#f8f9fa'>Compensar até</th><th style='text-align:right;padding:6px 8px;background:#f8f9fa'>Horas</th></tr>" + linhas + "</table>" +
      "<p>Você pode usar essas horas para <b>emendar as férias</b> ou para uma <b>folga avulsa</b>, pelo painel de férias (benefício <i>Banco de Horas</i>).</p>");
    enviados++;
  });
  return "Saldo enviado para " + enviados + " pessoa(s)." + (semEmail.length ? " Sem e-mail cadastrado: " + semEmail.join(", ") + "." : "");
}
