// =========================================================================
// SINCRONIZAÇÃO — importação do RH, envio em lote para o Calendar,
// atualização da aba de conflitos e rotina automática diária.
// Usado pela tela "Sincronizar" do dashboard.
// =========================================================================

const SINC_CONFIG = {
  idRhPadrao: "1xhmo4t63liWnr-GDQyNPbEIOHjfF-va_GxQ2uy4e-aE", // mesmo padrão de importarDadosPorIdPlanilha
  horaAutomatica: 7, // a rotina diária roda entre 7h e 8h
  funcaoAutomatica: "sincronizacaoAutomatica"
};

function obterStatusSincronizacao() {
  vincularPlanilha_();
  const props = PropertiesService.getScriptProperties();
  const gatilho = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === SINC_CONFIG.funcaoAutomatica)[0];
  let ultima = null;
  try { ultima = JSON.parse(props.getProperty("SINC_ULTIMA") || "null"); } catch (e) {}
  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  return {
    idRH: props.getProperty("ID_PLANILHA_RH") || SINC_CONFIG.idRhPadrao,
    automatico: !!gatilho,
    hora: SINC_CONFIG.horaAutomatica,
    ultima: ultima,
    podeConfigurar: String(usuario).toLowerCase() === String(emailGestor_()).toLowerCase()
  };
}

function importarRhDashboard(idRH) {
  idRH = String(idRH || "").trim();
  if (!idRH) throw new Error("Informe o ID da planilha do RH.");
  PropertiesService.getScriptProperties().setProperty("ID_PLANILHA_RH", idRH);
  const res = importarDadosPorIdPlanilha(idRH);
  registrarUltimaSinc_("Importação manual", res);
  return res;
}

function sincronizarCalendarDashboard() {
  const total = publicarLoteCalendario_();
  const res = total ? "Foram publicados " + total + " novos eventos na agenda." : "Nada novo para publicar: a agenda já está em dia.";
  registrarUltimaSinc_("Envio manual para o Calendar", res);
  return res;
}

function atualizarConflitosDashboard() {
  atualizarAbaControleConflitos(false);
  return "Aba 'Controle de Conflitos' atualizada.";
}

// Liga/desliga a rotina diária (somente o gestor)
function configurarSincronizacaoAutomatica(ativar) {
  validarGestor_();
  vincularPlanilha_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === SINC_CONFIG.funcaoAutomatica)
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (ativar) {
    ScriptApp.newTrigger(SINC_CONFIG.funcaoAutomatica)
      .timeBased().everyDays(1).atHour(SINC_CONFIG.horaAutomatica).inTimezone("America/Sao_Paulo")
      .create();
  }
  return obterStatusSincronizacao();
}

// Rotina diária: importa do RH, publica o lote na agenda e atualiza conflitos.
// Em caso de erro, avisa o gestor por e-mail.
function sincronizacaoAutomatica() {
  const etapas = [];
  const erros = [];
  const idRH = PropertiesService.getScriptProperties().getProperty("ID_PLANILHA_RH") || SINC_CONFIG.idRhPadrao;

  try { etapas.push(String(importarDadosPorIdPlanilha(idRH)).replace(/\n/g, " ")); }
  catch (e) { erros.push("Importação do RH: " + e.message); }

  try { etapas.push("Agenda: " + publicarLoteCalendario_() + " evento(s) novo(s)."); }
  catch (e) { erros.push("Envio para o Calendar: " + e.message); }

  try { atualizarAbaControleConflitos(false); etapas.push("Conflitos atualizados."); }
  catch (e) { erros.push("Controle de conflitos: " + e.message); }

  registrarUltimaSinc_("Automática", etapas.concat(erros.map(x => "❌ " + x)).join(" · "));

  if (erros.length) {
    enviarEmail_(emailGestor_(), "❌ Falha na sincronização automática de férias", "Sincronização automática",
      "<p>A rotina diária terminou com erro:</p><ul>" + erros.map(x => "<li>" + esc_(x) + "</li>").join("") + "</ul>" +
      (etapas.length ? "<p>Concluído: " + etapas.map(esc_).join(" · ") + "</p>" : ""));
  }
}

function registrarUltimaSinc_(origem, resumo) {
  PropertiesService.getScriptProperties().setProperty("SINC_ULTIMA", JSON.stringify({
    quando: Utilities.formatDate(new Date(), tz_(), "dd/MM/yyyy HH:mm"),
    origem: origem,
    resumo: String(resumo || "").replace(/\n/g, " ").slice(0, 400)
  }));
}
