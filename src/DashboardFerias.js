// =========================================================================
// DASHBOARD DE FÉRIAS
// -------------------------------------------------------------------------
// Abre pelo menu "Apoio Gestão > 📊 Abrir Dashboard de Férias".
// Reaproveita as regras de AprovacaoFerias.js (parseData_, situacaoIgnorada_,
// sobreposicao_) e grava lançamentos com salvarSugestaoDev (Código.js), que
// já dispara a aprovação quando há conflito.
// =========================================================================

const DASH_CONFIG = {
  diasHistorico: 60,   // "voltaram recentemente": quem retornou nos últimos N dias
  diasProximos: 30     // "saem em breve": quem sai nos próximos N dias
};

// Versão pelo link do App da Web (página inteira, sem botão Fechar)
function doGetDashboard_() {
  let eu = "";
  try { eu = String(Session.getActiveUser().getEmail()).toLowerCase(); } catch (e) {}
  if (emailLider_() && eu !== String(emailLider_()).toLowerCase()) {
    return paginaSimples_("Acesso pelo link restrito", "Abra o Dashboard de Férias pelo menu Apoio Gestão da planilha do time.");
  }
  return HtmlService.createHtmlOutputFromFile("dashboard")
    .append("<script>window.MODO_WEB=true;document.querySelectorAll('.so-dialogo').forEach(function(el){el.style.display='none';});</script>")
    .setTitle("Controle de Férias · " + nomeSquad_())
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function abrirDashboard() {
  vincularPlanilha_();
  const html = HtmlService.createHtmlOutputFromFile("dashboard")
    .setWidth(1200)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, " ");
}

// Chamado pelo dashboard. Datas sempre como "yyyy-MM-dd" (google.script.run
// não devolve objetos Date).
function obterDadosDashboard() {
  const ss = planilha_();
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const limiteHistorico = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - DASH_CONFIG.diasHistorico);

  // ---- Férias (Controle Consolidado) ----
  const periodos = [];
  const vistos = {};
  const controle = ss.getSheetByName(CONFIG_APROVACAO.abaControle);
  if (controle && controle.getLastRow() > 1) {
    const dados = controle.getRange(2, 1, controle.getLastRow() - 1, 17).getValues();
    dados.forEach((r, i) => {
      const nome = String(r[1] || "").trim().toUpperCase();
      const ini = parseData_(r[9]);
      const fim = parseData_(r[10]);
      if (!nome || !ini || !fim || fim < ini) return;
      if (fim < limiteHistorico) return;

      const situacao = String(r[4] || "").trim();
      const sitNorm = normalizar_(situacao);
      if (sitNorm.indexOf("REPROVADO") === 0 || sitNorm.indexOf("AJUSTE SOLICITADO") === 0 || sitNorm.indexOf("CANCELAD") === 0) return;

      const chave = nome + "|" + ini.getTime() + "|" + fim.getTime();
      if (vistos[chave]) return;
      vistos[chave] = true;

      const retorno = parseData_(r[13]) || parseData_(r[12]) ||
        new Date(fim.getFullYear(), fim.getMonth(), fim.getDate() + 1);

      periodos.push({
        linha: i + 2,
        nome: nome,
        tipo: "ferias",
        inicio: iso_(ini),
        fim: iso_(fim),
        retorno: iso_(retorno),
        dias: diasEntre_(ini, fim),
        situacao: situacao.replace(PREFIXO_PENDENTE, "") || "-",
        agendamento: String(r[8] || "").trim(),
        pendente: sitNorm.indexOf("PENDENTE") === 0,
        contaConflito: !situacaoIgnorada_(sitNorm)
      });
    });
  }

  // ---- Folgas avulsas (GiftCard, ResetDay, MVT...) ----
  const folgas = [];
  const abaFolgas = ss.getSheetByName("Folgas e Emendas - 2026");
  if (abaFolgas && abaFolgas.getLastRow() > 1) {
    abaFolgas.getRange(2, 1, abaFolgas.getLastRow() - 1, 5).getValues().forEach(r => {
      const nome = String(r[0] || "").trim().toUpperCase();
      const ini = parseData_(r[2]);
      if (!nome || !ini) return;
      const qtd = Math.max(1, parseInt(r[3], 10) || 1);
      const fim = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate() + qtd - 1);
      if (fim < limiteHistorico) return;
      folgas.push({
        nome: nome, tipo: "folga", beneficio: String(r[1] || "").trim(),
        inicio: iso_(ini), fim: iso_(fim), dias: qtd,
        retorno: iso_(parseData_(r[4]) || new Date(fim.getFullYear(), fim.getMonth(), fim.getDate() + 1))
      });
    });
  }

  // ---- Conflitos entre férias vigentes/futuras ----
  const ativos = periodos.filter(p => p.contaConflito && parseData_(p.fim) >= hoje);
  const conflitos = [];
  for (let a = 0; a < ativos.length; a++) {
    for (let b = a + 1; b < ativos.length; b++) {
      const p1 = ativos[a], p2 = ativos[b];
      if (p1.nome === p2.nome || !conflitoRelevante_(p1.nome, p2.nome)) continue;
      const s = sobreposicao_(parseData_(p1.inicio), parseData_(p1.fim), parseData_(p2.inicio), parseData_(p2.fim));
      if (!s) continue;
      conflitos.push({
        a: { nome: p1.nome, inicio: p1.inicio, fim: p1.fim, pendente: p1.pendente },
        b: { nome: p2.nome, inicio: p2.inicio, fim: p2.fim, pendente: p2.pendente },
        inicio: iso_(s.inicio), fim: iso_(s.fim), dias: s.dias
      });
    }
  }
  // Líder × PO/TLs de outras planilhas
  let externos = [];
  try { if (liderNome_()) externos = periodosExternosPares_(); } catch (e) {}
  ativos.filter(p => ehNomeLider_(p.nome)).forEach(p1 => {
    externos.forEach(p2 => {
      const s = sobreposicao_(parseData_(p1.inicio), parseData_(p1.fim), p2.inicio, p2.fim);
      if (!s) return;
      conflitos.push({
        a: { nome: p1.nome, inicio: p1.inicio, fim: p1.fim, pendente: p1.pendente },
        b: { nome: p2.nome, inicio: iso_(p2.inicio), fim: iso_(p2.fim), pendente: false, externo: true },
        inicio: iso_(s.inicio), fim: iso_(s.fim), dias: s.dias
      });
    });
  });
  conflitos.sort((x, y) => x.inicio < y.inicio ? -1 : x.inicio > y.inicio ? 1 : y.dias - x.dias);

  // ---- Pendentes de aprovação ----
  const pendentes = [];
  const abaApr = ss.getSheetByName(CONFIG_APROVACAO.abaAprovacoes);
  if (abaApr && abaApr.getLastRow() > 1) {
    let eu = "";
    try { eu = String(Session.getActiveUser().getEmail()).toLowerCase(); } catch (e) {}
    const urlBase = urlWebApp_();
    abaApr.getRange(2, 1, abaApr.getLastRow() - 1, COL_APR.TIPO).getValues().forEach(r => {
      if (r[COL_APR.STATUS - 1] !== STATUS_APR.PENDENTE) return;
      const aprovador = String(r[COL_APR.APROVADOR - 1] || emailLider_()).toLowerCase();
      const url = eu && eu === aprovador ? urlBase : "";
      pendentes.push({
        id: String(r[0]), nome: String(r[2]),
        inicio: iso_(parseData_(r[3])), fim: iso_(parseData_(r[4])),
        conflitos: String(r[6] || ""),
        tipo: String(r[COL_APR.TIPO - 1] || "CONFLITO"), aprovador: aprovador,
        link: url ? url + "?id=" + encodeURIComponent(String(r[0])) + "&t=" + String(r[COL_APR.TOKEN - 1]) : ""
      });
    });
  }

  let time = [];
  try { time = buscarNomesTime().map(n => String(n).toUpperCase()); } catch (e) {}

  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}

  return {
    hoje: iso_(hoje),
    usuario: usuario,
    time: time,
    periodos: periodos,
    folgas: folgas,
    conflitos: conflitos,
    pendentes: pendentes,
    config: DASH_CONFIG,
    nomeSquad: nomeSquad_(),
    liderNome: liderNome_(),
    pares: paresLider_().map(p => p.nome),
    paresPeriodos: externos.map(p => ({ nome: p.nome, inicio: iso_(p.inicio), fim: iso_(p.fim) })),
    lider: ehLider_(liderNome_()) ? liderNome_() : "", gerente: gerenteNome_() || gerenteEmail_(),
    configurado: timeConfigurado_(),
    ehGestor: !!usuario && String(usuario).toLowerCase() === String(emailLider_()).toLowerCase()
  };
}

// Guarda o ID da planilha para o App da Web (que roda sem planilha aberta)
function vincularPlanilha_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    if (ss && props.getProperty("ID_PLANILHA") !== ss.getId()) props.setProperty("ID_PLANILHA", ss.getId());
  } catch (e) {}
}
