// =========================================================================
// CANCELAMENTO DE FÉRIAS
// -------------------------------------------------------------------------
// O gestor cancela um lançamento (pendente ou já aprovado) com um motivo:
//  - a linha do Controle Consolidado fica como "Cancelado" (não conta mais
//    em conflito, avisos, saldos nem volta para a agenda);
//  - a solicitação pendente, se houver, fica "Cancelada" na aba Aprovações;
//  - o evento sai da agenda do time (e da Tribe, se for férias do líder);
//  - a pessoa e, se havia aprovação pendente, o aprovador recebem e-mail.
// =========================================================================

const SITUACAO_CANCELADO = "Cancelado";

function ehCancelado_(situacao) { return normalizar_(situacao).indexOf("CANCELAD") === 0; }

// linha: linha no Controle Consolidado; inicioIso/fimIso conferem que é a mesma férias
function cancelarFerias(linha, inicioIso, fimIso, motivo) {
  validarGestor_();
  motivo = String(motivo || "").trim();
  if (!motivo) throw new Error("Informe o motivo do cancelamento.");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const controle = planilha_().getSheetByName(CONFIG_APROVACAO.abaControle);
    const alvoIni = parseData_(inicioIso), alvoFim = parseData_(fimIso);
    linha = localizarLinhaPorDatas_(controle, Number(linha), alvoIni, alvoFim);
    if (!linha) throw new Error("Não achei essas férias no Controle Consolidado. Atualize o dashboard e tente de novo.");

    const row = controle.getRange(linha, 1, 1, 17).getValues()[0];
    const nome = String(row[COL_CTRL.NOME - 1]).trim();
    const ini = parseData_(row[COL_CTRL.INICIO - 1]);
    const fim = parseData_(row[COL_CTRL.FIM - 1]);
    const retorno = parseData_(row[COL_CTRL.RETORNO_AFAST - 1]);
    const situacao = String(row[COL_CTRL.SITUACAO - 1] || "");
    if (ehCancelado_(situacao)) throw new Error("Essas férias já estão canceladas.");
    const estavaNoRH = normalizar_(row[8]) === "AGENDADOS" || /MARCAD/i.test(situacao);
    const estavaNaAgenda = String(row[COL_CTRL.STATUS - 1]).indexOf("Espelhado") !== -1;
    const quem = Session.getActiveUser().getEmail();
    const quando = Utilities.formatDate(new Date(), tz_(), "dd/MM/yyyy");

    // Solicitação pendente na aba Aprovações
    let aprovadorPendente = "";
    const abaApr = planilha_().getSheetByName(CONFIG_APROVACAO.abaAprovacoes);
    if (abaApr && abaApr.getLastRow() > 1) {
      const dados = abaApr.getRange(2, 1, abaApr.getLastRow() - 1, COL_APR.TIPO).getValues();
      dados.forEach((r, i) => {
        if (r[COL_APR.STATUS - 1] !== STATUS_APR.PENDENTE) return;
        const ri = parseData_(r[COL_APR.INICIO - 1]), rf = parseData_(r[COL_APR.FIM - 1]);
        if (normalizar_(r[COL_APR.NOME - 1]) !== normalizar_(nome) || !ri || !rf ||
            ri.getTime() !== ini.getTime() || rf.getTime() !== fim.getTime()) return;
        abaApr.getRange(i + 2, COL_APR.STATUS, 1, 4).setValues([["Cancelada", quem, new Date(), "Cancelada: " + motivo]]);
        aprovadorPendente = String(r[COL_APR.APROVADOR - 1] || emailGestor_()).toLowerCase();
      });
    }

    // Agenda do time e da Tribe
    if (estavaNaAgenda || estavaNoRH) removerEventoAgenda_(nome.toUpperCase(), ini, retorno || fim, true);
    if (ehNomeLider_(nome) && idAgendaTribe_()) removerEventoDeAgenda_(idAgendaTribe_(), nome.toUpperCase(), ini, retorno || fim);

    // Linha do Controle
    controle.getRange(linha, COL_CTRL.SITUACAO).setValue(SITUACAO_CANCELADO);
    controle.getRange(linha, 9).setValue("CANCELADO");
    controle.getRange(linha, COL_CTRL.STATUS).setValue("Cancelado por " + quem + " em " + quando + ": " + motivo);
    controle.getRange(linha, 1, 1, 17).setBackground("#eeeeee").setFontLine("line-through");
    SpreadsheetApp.flush();
    try { atualizarAbaControleConflitos(false); } catch (e) { Logger.log(e); }

    // E-mails
    const periodo = fmt_(ini) + " a " + fmt_(fim);
    const corpo = "<p>As férias de <b>" + esc_(nome) + "</b> de <b>" + periodo + "</b> foram <b>canceladas</b>.</p>" +
      "<p><b>Motivo:</b> " + esc_(motivo) + "</p>" +
      (estavaNoRH ? "<p>Essas férias estavam marcadas no sistema do RH: o cancelamento também precisa ser feito lá.</p>" : "") +
      "<p style='color:#5f6368;font-size:12px'>Cancelado por " + esc_(quem) + " em " + quando + ".</p>";
    const emailPessoa = obterEmail_(nome);
    if (emailPessoa && emailPessoa.toLowerCase() !== String(quem).toLowerCase()) {
      enviarEmail_(emailPessoa, "🗓️ Férias canceladas: " + periodo, "Férias canceladas", corpo);
    }
    if (aprovadorPendente && aprovadorPendente !== String(quem).toLowerCase()) {
      enviarEmail_(aprovadorPendente, "🗓️ Solicitação de férias cancelada: " + nome, "Solicitação cancelada",
        "<p>Não é mais preciso decidir esta solicitação.</p>" + corpo);
    }

    return "Férias de " + nome + " (" + periodo + ") canceladas" +
      (estavaNaAgenda ? " e retiradas da agenda" : "") + "." +
      (emailPessoa ? " Aviso enviado por e-mail." : "") +
      (estavaNoRH ? " ⚠️ Estavam marcadas no RH: cancele também no sistema do RH." : "");
  } finally {
    lock.releaseLock();
  }
}

function localizarLinhaPorDatas_(sheet, linha, ini, fim) {
  if (!ini || !fim) return null;
  const dados = sheet.getDataRange().getValues();
  const confere = i => {
    const a = parseData_(dados[i][COL_CTRL.INICIO - 1]), b = parseData_(dados[i][COL_CTRL.FIM - 1]);
    return a && b && a.getTime() === ini.getTime() && b.getTime() === fim.getTime();
  };
  if (linha >= 2 && linha <= dados.length && confere(linha - 1)) return linha;
  return null;
}

// Remove eventos "Ausência: NOME" de uma agenda qualquer (ex.: Tribe)
function removerEventoDeAgenda_(idAgenda, nome, ini, fim) {
  try {
    const agenda = CalendarApp.getCalendarById(idAgenda);
    if (!agenda || !ini) return;
    const ate = new Date((fim || ini).getFullYear(), (fim || ini).getMonth(), (fim || ini).getDate() + 2);
    agenda.getEvents(ini, ate).forEach(ev => {
      if (ev.getTitle().toUpperCase().indexOf("AUSÊNCIA: " + nome) === 0) ev.deleteEvent();
    });
  } catch (e) { Logger.log("Agenda " + idAgenda + ": " + e.message); }
}
