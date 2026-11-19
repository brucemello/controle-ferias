// =========================================================================
// CONFIGURAÇÃO POR TIME
// -------------------------------------------------------------------------
// Tudo o que é específico de um time (nome, gestor, agenda, destinatários,
// webhook do Chat) fica nas Propriedades do Script — que NÃO são copiadas
// quando alguém faz uma cópia da planilha. Assim, uma cópia nasce "em branco"
// e o dashboard abre a tela "Configurar para o meu time".
//
// A planilha original (Squad Novas Apurações) usa os valores de
// PADRAO_ORIGINAL enquanto não for configurada pela tela.
// =========================================================================

const ID_PLANILHA_ORIGINAL = "1ox7bG8Z1NqSJVINDF2EjAq00XpjnbvSyD4SapgXvjD8";

const PADRAO_ORIGINAL = {
  NOME_SQUAD: "Squad Novas Apurações",
  EMAIL_GESTOR: "bruce.mello@totvs.com.br",
  ID_AGENDA: "c_309f2f359160bb1d732268dbb2e33b80dda091c626cbb2f997c8066a87c004af@group.calendar.google.com",
  EMAILS_AVISO: "bruce.mello@totvs.com.br,squad.tributos@totvs.com.br",
  ID_PLANILHA_RH: "1xhmo4t63liWnr-GDQyNPbEIOHjfF-va_GxQ2uy4e-aE"
};

const CHAVES_CONFIG = ["NOME_SQUAD", "EMAIL_GESTOR", "ID_AGENDA", "EMAILS_AVISO", "WEBHOOK_CHAT"];

function ehPlanilhaOriginal_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss.getId() === ID_PLANILHA_ORIGINAL;
  } catch (e) {}
  return PropertiesService.getScriptProperties().getProperty("ID_PLANILHA") === ID_PLANILHA_ORIGINAL;
}

function cfg_(chave) {
  const v = PropertiesService.getScriptProperties().getProperty(chave);
  if (v !== null && v !== "") return v;
  return ehPlanilhaOriginal_() ? (PADRAO_ORIGINAL[chave] || "") : "";
}

function nomeSquad_() { return cfg_("NOME_SQUAD") || "Meu time"; }
function emailLider_() { return cfg_("EMAIL_GESTOR"); }
function idAgenda_() { return String(cfg_("ID_AGENDA") || "").trim(); }
function emailsAviso_() { return cfg_("EMAILS_AVISO"); }
function webhookChat_() { return cfg_("WEBHOOK_CHAT"); }

function timeConfigurado_() {
  return !!(emailLider_() && idAgenda_());
}

// ---------- usado pela tela de configuração do dashboard ----------
function obterConfiguracaoTime() {
  vincularPlanilha_();
  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  const gestor = emailLider_();
  let agendas = [];
  try {
    agendas = CalendarApp.getAllOwnedCalendars().map(c => ({ id: c.getId(), nome: c.getName() }))
      .sort((a, b) => a.nome.localeCompare(b.nome));
  } catch (e) {}
  const atual = idAgenda_();
  if (atual && !agendas.some(a => a.id === atual)) {
    let nome = "Agenda configurada";
    try { const c = CalendarApp.getCalendarById(atual); if (c) nome = c.getName(); } catch (e) {}
    agendas.unshift({ id: atual, nome: nome });
  }
  let time = [];
  try { time = buscarNomesTime(); } catch (e) {}
  return {
    configurado: timeConfigurado_(),
    original: ehPlanilhaOriginal_(),
    podeEditar: !gestor || String(usuario).toLowerCase() === String(gestor).toLowerCase(),
    usuario: usuario,
    nomeSquad: cfg_("NOME_SQUAD"),
    emailGestor: gestor,
    idAgenda: atual,
    emailsAviso: emailsAviso_(),
    temWebhook: !!webhookChat_(),
    agendas: agendas,
    time: time,
    contatos: lerContatos_()
  };
}

// dados: { nomeSquad, emailGestor, idAgenda ("NOVA" cria), emailsAviso, webhook,
//          membros: [{nome, email}], limparDados: bool, confirmacao: "APAGAR" }
function salvarConfiguracaoTime(dados) {
  const atualGestor = emailLider_();
  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  if (atualGestor && String(usuario).toLowerCase() !== String(atualGestor).toLowerCase()) {
    throw new Error("Só o gestor configurado (" + atualGestor + ") pode alterar a configuração.");
  }

  const nomeSquad = String(dados.nomeSquad || "").trim();
  const emailGestor = String(dados.emailGestor || "").trim().toLowerCase();
  if (!nomeSquad) throw new Error("Informe o nome do time.");
  if (!/^[^@\s]+@[^@\s]+$/.test(emailGestor)) throw new Error("Informe um e-mail de gestor válido.");

  const ss = planilha_();
  const props = PropertiesService.getScriptProperties();

  // Agenda
  let idAgenda = String(dados.idAgenda || "").trim();
  if (idAgenda === "NOVA") {
    const nova = CalendarApp.createCalendar("Ausências " + nomeSquad, { color: CalendarApp.Color.TEAL });
    idAgenda = nova.getId();
  }
  if (!idAgenda) throw new Error("Escolha a agenda do time.");

  // Limpeza dos dados do time anterior (exige confirmação)
  if (dados.limparDados) {
    if (String(dados.confirmacao || "").trim().toUpperCase() !== "APAGAR") {
      throw new Error("Para apagar os dados, digite APAGAR no campo de confirmação.");
    }
    limparDadosTime_(ss);
    // Desliga rotinas e esquece origem do RH/histórico do time anterior
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
    ["ID_PLANILHA_RH", "RH_MODO", "RH_ULTIMO_ID", "SINC_ULTIMA", "URL_WEBAPP"].forEach(k => props.deleteProperty(k));
  }

  props.setProperties({
    NOME_SQUAD: nomeSquad,
    EMAIL_GESTOR: emailGestor,
    ID_AGENDA: idAgenda,
    EMAILS_AVISO: String(dados.emailsAviso || "").trim(),
    ID_PLANILHA: ss.getId()
  });
  if (dados.webhook === "") props.deleteProperty("WEBHOOK_CHAT");
  else if (dados.webhook) props.setProperty("WEBHOOK_CHAT", String(dados.webhook).trim());

  // Time e contatos
  if (Array.isArray(dados.membros)) gravarMembros_(ss, dados.membros);

  return "Configuração salva para " + nomeSquad + "." + (dados.limparDados ? " Dados do time anterior apagados." : "");
}

// Transforma uma CÓPIA em modelo em branco: apaga dados, configurações e rotinas.
// Bloqueado na planilha original.
function transformarEmModelo(confirmacao) {
  if (ehPlanilhaOriginal_()) throw new Error("Isso não pode ser feito na planilha original. Faça uma cópia primeiro.");
  if (String(confirmacao || "").trim().toUpperCase() !== "APAGAR") throw new Error("Digite APAGAR para confirmar.");
  const ss = planilha_();
  limparDadosTime_(ss);
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().deleteAllProperties();
  return "Pronto: esta planilha virou um modelo em branco. Quem fizer uma cópia dela começa do zero.";
}

function lerContatos_() {
  const aba = planilha_().getSheetByName("Contatos");
  if (!aba || aba.getLastRow() < 2) return {};
  const m = {};
  aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) m[normalizar_(r[0])] = String(r[1] || ""); });
  return m;
}

function gravarMembros_(ss, membros) {
  const lista = membros
    .map(m => ({ nome: String(m.nome || "").trim(), email: String(m.email || "").trim() }))
    .filter(m => m.nome);

  const time = ss.getSheetByName("TIME") || ss.insertSheet("TIME");
  if (time.getLastRow() > 1) time.getRange(2, 1, time.getLastRow() - 1, Math.max(1, time.getLastColumn())).clearContent();
  time.getRange(1, 1).setValue("Nome");
  if (lista.length) time.getRange(2, 1, lista.length, 1).setValues(lista.map(m => [m.nome]));

  let contatos = ss.getSheetByName("Contatos");
  if (!contatos) { contatos = ss.insertSheet("Contatos"); }
  if (contatos.getLastRow() > 1) contatos.getRange(2, 1, contatos.getLastRow() - 1, Math.max(2, contatos.getLastColumn())).clearContent();
  contatos.getRange(1, 1, 1, 2).setValues([["Nome", "Email"]]).setFontWeight("bold");
  if (lista.length) contatos.getRange(2, 1, lista.length, 2).setValues(lista.map(m => [m.nome.toUpperCase(), m.email]));
}

// Mantém cabeçalhos e estrutura; apaga só as linhas de dados
function limparDadosTime_(ss) {
  const abas = {
    "Controle Consolidado - 2026": 2, "Folgas e Emendas - 2026": 2, "Aprovações": 2,
    "Controle de Conflitos": 6, "Admisssão": 2, "MVT": 2, "TIME": 2, "Contatos": 2
  };
  Object.keys(abas).forEach(nome => {
    const aba = ss.getSheetByName(nome);
    if (!aba) return;
    const ini = abas[nome], ult = aba.getLastRow();
    if (ult >= ini) {
      const rng = aba.getRange(ini, 1, ult - ini + 1, Math.max(1, aba.getLastColumn()));
      rng.clearContent();
      rng.setBackground(null);
    }
  });
}
