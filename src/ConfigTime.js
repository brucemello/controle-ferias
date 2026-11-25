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
  ID_PLANILHA_RH: "1xhmo4t63liWnr-GDQyNPbEIOHjfF-va_GxQ2uy4e-aE",
  LIDER_NOME: "BRUCE EGNOR MAGALHAES DE MELLO",
  GERENTE_NOME: "ROBSON ALVES BEZERRA",
  GERENTE_EMAIL: "robson.bezerra@totvs.com.br",
  PARES_LIDER: "LEANDRO HEILIG PASSOS"
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
  if (v === " ") return "";               // apagado de propósito na Configuração
  if (v !== null && v !== "") return v;
  return ehPlanilhaOriginal_() ? (PADRAO_ORIGINAL[chave] || "") : "";
}

function nomeSquad_() { return cfg_("NOME_SQUAD") || "Meu time"; }
function emailLider_() { return cfg_("EMAIL_GESTOR"); }
function idAgenda_() { return String(cfg_("ID_AGENDA") || "").trim(); }
function emailsAviso_() { return cfg_("EMAILS_AVISO"); }
function webhookChat_() { return cfg_("WEBHOOK_CHAT"); }

// Hierarquia: as férias do líder vão para aprovação do gerente
function liderNome_() { return String(cfg_("LIDER_NOME") || "").trim().toUpperCase(); }
function gerenteNome_() { return String(cfg_("GERENTE_NOME") || "").trim(); }
function gerenteEmail_() { return String(cfg_("GERENTE_EMAIL") || "").trim().toLowerCase(); }
function idAgendaTribe_() { return String(cfg_("ID_AGENDA_TRIBE") || "").trim(); }
// PO e outros TLs: as férias do líder só conflitam com eles (não com o time).
// Uma pessoa por linha: "NOME" ou "NOME; ID ou link da planilha de controle dela".
function paresLider_() {
  return String(cfg_("PARES_LIDER") || "").split(/\r?\n/).map(l => {
    const [nome, plan] = l.split(";");
    const n = String(nome || "").trim().toUpperCase();
    if (!n) return null;
    const m = String(plan || "").match(/[-\w]{25,}/);
    return { nome: n, planilha: m ? m[0] : "" };
  }).filter(Boolean);
}
function ehNomeLider_(nome) { return !!liderNome_() && normalizar_(nome) === normalizar_(liderNome_()); }
function ehParLider_(nome) { const n = normalizar_(nome); return paresLider_().some(p => normalizar_(p.nome) === n); }
// Regra de conflito: líder × (PO/TLs) apenas; líder não entra no conflito dos devs.
function conflitoRelevante_(a, b) {
  if (!liderNome_()) return true;
  const la = ehNomeLider_(a), lb = ehNomeLider_(b);
  if (la || lb) return ehParLider_(la ? b : a);
  return true;
}
// Férias do PO/TLs lidas das planilhas de controle deles (cache por 10 min)
function periodosExternosPares_() {
  const pares = paresLider_().filter(p => p.planilha);
  if (!pares.length) return [];
  const cache = CacheService.getScriptCache();
  const chave = "PARES_EXT_" + pares.map(p => p.nome + p.planilha).join("|").length + "_" + pares.length;
  try { const c = cache.get(chave); if (c) return JSON.parse(c).map(x => ({ nome: x.nome, inicio: new Date(x.i), fim: new Date(x.f), linha: 0, externo: true })); } catch (e) {}
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const lista = [];
  pares.forEach(par => {
    try {
      const ss = SpreadsheetApp.openById(par.planilha);
      const aba = ss.getSheetByName(CONFIG_APROVACAO.abaControle) || ss.getSheets().filter(s => /controle consolidado/i.test(s.getName()))[0];
      if (!aba || aba.getLastRow() < 2) return;
      const alvo = normalizar_(par.nome), vistos = {};
      aba.getRange(2, 1, aba.getLastRow() - 1, 11).getValues().forEach(r => {
        if (normalizar_(r[COL_CTRL.NOME - 1]) !== alvo) return;
        const sit = normalizar_(r[COL_CTRL.SITUACAO - 1]);
        if (situacaoIgnorada_(sit) || sit.indexOf("REPROVADO") === 0 || sit.indexOf("AJUSTE SOLICITADO") === 0) return;
        const ini = parseData_(r[COL_CTRL.INICIO - 1]), fim = parseData_(r[COL_CTRL.FIM - 1]);
        if (!ini || !fim || fim < hoje) return;
        const k = ini.getTime() + "|" + fim.getTime();
        if (vistos[k]) return; vistos[k] = 1;
        lista.push({ nome: par.nome, inicio: ini, fim: fim, linha: 0, externo: true });
      });
    } catch (e) { Logger.log("Planilha de " + par.nome + ": " + e.message); }
  });
  try { cache.put(chave, JSON.stringify(lista.map(x => ({ nome: x.nome, i: x.inicio.getTime(), f: x.fim.getTime() }))), 600); } catch (e) {}
  return lista;
}
function ehLider_(nome) { return !!liderNome_() && !!gerenteEmail_() && normalizar_(nome) === normalizar_(liderNome_()); }

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
    urlWebApp: PropertiesService.getScriptProperties().getProperty("URL_WEBAPP") || "",
    liderNome: liderNome_(), gerenteNome: gerenteNome_(), gerenteEmail: gerenteEmail_(), idAgendaTribe: idAgendaTribe_(),
    paresLider: String(cfg_("PARES_LIDER") || "").trim(),
    idFotos: (function() { try { return idPlanilhaFotos_(); } catch (e) { return ""; } })(),
    urlWebAppAuto: (function() { try { const s = ScriptApp.getService(); return s.isEnabled() ? (s.getUrl() || "") : ""; } catch (e) { return ""; } })(),
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
  // Líder / gerente / agenda da Tribe
  if (typeof dados.liderNome === "string") {
    const ge = String(dados.gerenteEmail || "").trim().toLowerCase();
    if (ge && !/^[^@\s]+@[^@\s]+$/.test(ge)) throw new Error("Informe um e-mail de gerente válido.");
    if (ge && ge === emailGestor) throw new Error("O e-mail do gerente não pode ser o mesmo do gestor do time: informe o e-mail de quem aprova as férias do líder.");
    const pares = { LIDER_NOME: String(dados.liderNome || "").trim().toUpperCase(), GERENTE_NOME: String(dados.gerenteNome || "").trim(),
      GERENTE_EMAIL: ge, ID_AGENDA_TRIBE: String(dados.idAgendaTribe || "").trim(),
      PARES_LIDER: String(dados.paresLider || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean).join("\n") };
    Object.keys(pares).forEach(k => { if (pares[k]) props.setProperty(k, pares[k]); else props.setProperty(k, " "); });
  }
  if (typeof dados.urlWebApp === "string") {
    const u = dados.urlWebApp.trim();
    if (!u) props.deleteProperty("URL_WEBAPP");
    else if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(u)) throw new Error("O link do App da Web deve começar com https://script.google.com/ e terminar com /exec.");
    else props.setProperty("URL_WEBAPP", u);
  }
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

// Lê os nomes das pessoas na exportação do RH (Looker) mais recente, sem gravar nada.
// A tela de Configuração junta com a lista atual para o gestor revisar e salvar.
function lerPessoasDoRh() {
  const origem = idRhParaImportar_();
  const ssRH = SpreadsheetApp.openById(origem.id);
  const aba = ssRH.getSheetByName("Sheet1") || ssRH.getSheets()[0];
  if (aba.getLastRow() < 2) throw new Error("A exportação do RH está vazia.");
  const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => normalizar_(c));
  let iNome = cab.indexOf("NOME"); if (iNome < 0) iNome = 1;
  const vistos = {}, nomes = [];
  aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues().forEach(r => {
    const n = String(r[iNome] || "").trim().toUpperCase();
    if (n && !vistos[n]) { vistos[n] = true; nomes.push(n); }
  });
  return { nomes: nomes.sort(), origem: origem.nome + (origem.criada ? " (" + origem.criada + ")" : "") };
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
    "Controle de Conflitos": 6, "Admisssão": 2, "MVT": 2, "TIME": 2, "Contatos": 2,
    "Fotos": 2, "Banco de Horas": 2, "Férias Incentivadas": 2, "Adesões Incentivadas": 2
  };
  Object.keys(abas).forEach(nome => {
    const aba = ss.getSheetByName(nome);
    if (!aba) return;
    // Tira critérios de filtro herdados e reexibe linhas ocultas
    try {
      const f = aba.getFilter();
      if (f) {
        const c0 = f.getRange().getColumn(), n = f.getRange().getNumColumns();
        for (let c = c0; c < c0 + n; c++) { if (f.getColumnFilterCriteria(c)) f.removeColumnFilterCriteria(c); }
      }
      aba.showRows(1, aba.getMaxRows());
    } catch (e) {}
    const ini = abas[nome], ult = aba.getLastRow();
    if (ult >= ini) {
      const rng = aba.getRange(ini, 1, ult - ini + 1, Math.max(1, aba.getLastColumn()));
      rng.clearContent();
      rng.setBackground(null);
    }
  });
}
