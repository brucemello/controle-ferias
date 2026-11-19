// =========================================================================
// PRIMEIROS PASSOS — assistente guiado de configuração e acessos
// -------------------------------------------------------------------------
// verificarSetup() confere cada etapa e devolve o status para a tela
// "Primeiros passos" do dashboard. As funções ativar...() executam as
// etapas que podem ser feitas com um clique.
// =========================================================================

const HORA_AVISO_DIARIO = 8; // aviso diário de férias (e-mail/Chat) entre 8h e 9h

function verificarSetup() {
  vincularPlanilha_();
  const passos = [];
  const add = (id, ok, detalhe, extra) => passos.push(Object.assign({ id: id, ok: ok, detalhe: detalhe }, extra || {}));

  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}

  // 1. Permissões
  const faltando = [];
  try { DriveApp.searchFiles('title contains "' + SINC_CONFIG.tituloExportacao + '"').hasNext(); } catch (e) { faltando.push("Google Drive"); }
  try { ScriptApp.getProjectTriggers(); } catch (e) { faltando.push("Rotinas automáticas"); }
  try { CalendarApp.getAllOwnedCalendars(); } catch (e) { faltando.push("Google Agenda"); }
  try { MailApp.getRemainingDailyQuota(); } catch (e) { faltando.push("Envio de e-mail"); }
  add("permissoes", faltando.length === 0,
    faltando.length ? "Falta autorizar: " + faltando.join(", ") + "." : "Drive, Agenda, e-mail e rotinas autorizados para " + (usuario || "você") + ".");

  // 2. Dados do time
  const nome = PropertiesService.getScriptProperties().getProperty("NOME_SQUAD") || (ehPlanilhaOriginal_() ? nomeSquad_() : "");
  const gestor = emailLider_();
  add("time", !!(nome && gestor),
    nome && gestor ? nome + " · gestor(a): " + gestor : "Nome do time e e-mail do gestor ainda não informados.");

  // 3. Pessoas
  let pessoas = [];
  try { pessoas = buscarNomesTime(); } catch (e) {}
  const contatos = lerContatos_();
  const semEmail = pessoas.filter(n => !contatos[normalizar_(n)]);
  add("pessoas", pessoas.length > 0 && semEmail.length === 0,
    !pessoas.length ? "Nenhuma pessoa cadastrada." :
      pessoas.length + " pessoa(s)" + (semEmail.length ? " · sem e-mail: " + semEmail.slice(0, 4).map(n => primeiroNome_(n)).join(", ") + (semEmail.length > 4 ? "…" : "") : ", todas com e-mail."),
    { parcial: pessoas.length > 0 && semEmail.length > 0 });

  // 4. Agenda
  let agendaNome = "", agendaOk = false;
  if (idAgenda_()) {
    try { const c = CalendarApp.getCalendarById(idAgenda_()); if (c) { agendaOk = true; agendaNome = c.getName(); } } catch (e) {}
  }
  add("agenda", agendaOk,
    agendaOk ? "Publicando em: " + agendaNome : idAgenda_() ? "A agenda configurada não está acessível para você." : "Nenhuma agenda escolhida.");

  // 5. App da Web (botão de aprovação dos e-mails)
  const url = urlWebApp_();
  add("webapp", !!url, url ? "Link de aprovação ativo." : "Ainda não implantado: os e-mails de conflito vão sem o botão de decisão.", { url: url });

  // 6. Exportação do RH
  let ex = null;
  try { ex = exportacaoMaisRecente_(); } catch (e) {}
  const modo = modoRh_();
  const idFixo = cfg_("ID_PLANILHA_RH");
  add("rh", modo === "auto" ? !!ex : !!idFixo,
    modo === "auto" ? (ex ? "Última exportação: " + ex.nome + " (" + ex.criada + ")" : "Nenhuma planilha " + SINC_CONFIG.tituloExportacao + " encontrada no seu Drive.")
      : (idFixo ? "Usando ID fixo." : "Nenhum ID informado."));

  // 7-9. Rotinas (gatilhos de quem está usando)
  let gat = [];
  try { gat = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()); } catch (e) {}
  const tem = f => gat.indexOf(f) !== -1;
  add("sinc", tem(SINC_CONFIG.funcaoAutomatica),
    tem(SINC_CONFIG.funcaoAutomatica) ? "Ligada: importa após cada exportação e todo dia às " + SINC_CONFIG.horaAutomatica + "h." : "Desligada.");
  add("aviso", tem("notificarProximasFerias"),
    tem("notificarProximasFerias") ? "Ligado: todo dia às " + HORA_AVISO_DIARIO + "h para " + (emailsAviso_() || "(sem destinatários)") + (webhookChat_() ? " + Chat" : "") + "."
      : "Desligado.", { opcional: true });
  add("abertura", tem("gatilhoAberturaAutomatica") || tem("abrirDashboard"),
    tem("gatilhoAberturaAutomatica") || tem("abrirDashboard") ? "O dashboard abre sozinho ao abrir a planilha." : "Desligada: abra pelo menu Apoio Gestão.", { opcional: true });

  return {
    passos: passos,
    usuario: usuario,
    ehGestor: !!usuario && String(usuario).toLowerCase() === String(gestor || "").toLowerCase(),
    semGestor: !gestor,
    idScript: ScriptApp.getScriptId(),
    original: ehPlanilhaOriginal_()
  };
}

function exigirGestorOuSemConfig_() {
  const g = emailLider_();
  if (!g) return;
  validarGestor_();
}

function ativarAvisoDiario(ligar) {
  exigirGestorOuSemConfig_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "notificarProximasFerias").forEach(t => ScriptApp.deleteTrigger(t));
  if (ligar) {
    if (!emailsAviso_() && !webhookChat_()) throw new Error("Informe os destinatários do aviso (ou o webhook do Chat) em Configuração antes de ligar.");
    ScriptApp.newTrigger("notificarProximasFerias").timeBased().everyDays(1).atHour(HORA_AVISO_DIARIO).inTimezone("America/Sao_Paulo").create();
  }
  return verificarSetup();
}

function ativarAberturaAutomatica(ligar) {
  ScriptApp.getProjectTriggers()
    .filter(t => ["gatilhoAberturaAutomatica", "abrirDashboard"].indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (ligar) ScriptApp.newTrigger("gatilhoAberturaAutomatica").forSpreadsheet(planilha_()).onOpen().create();
  return verificarSetup();
}

// Envia um e-mail de teste para quem está configurando (confere e-mail + nome do time)
function enviarEmailTeste() {
  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!usuario) throw new Error("Não consegui identificar o seu e-mail.");
  enviarEmail_(usuario, "✅ Teste do Controle de Férias - " + nomeSquad_(), "Tudo certo! 🎉",
    "<p>Este é um e-mail de teste do <b>Controle de Férias</b> do time <b>" + esc_(nomeSquad_()) + "</b>.</p>" +
    "<p>Gestor(a) configurado(a): <b>" + esc_(emailLider_() || "-") + "</b><br>Agenda: <b>" + esc_(idAgenda_() ? "configurada" : "não configurada") + "</b></p>" +
    (urlWebApp_() ? "<p style='text-align:center'><a href='" + urlWebApp_() + "' style='background:#14b8a6;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold'>Abrir dashboard pelo link</a></p>" : ""));
  return "E-mail de teste enviado para " + usuario + ".";
}
