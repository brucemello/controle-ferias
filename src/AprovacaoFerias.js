// =========================================================================
// APROVAÇÃO DE FÉRIAS COM CONFLITO — Squad Novas Apurações
// -------------------------------------------------------------------------
// Fluxo:
//  1. salvarSugestaoDev grava a nova linha e chama processarLancamentoFerias(linha).
//  2. Sem conflito: segue como antes (espelha no Calendar).
//  3. Com conflito:
//     - a Situação vira "Pendente Aprovação | <situação original>";
//     - NÃO espelha no Calendar ainda;
//     - o gestor recebe e-mail com o botão "Analisar e decidir";
//     - o dev que lançou é avisado de que está pendente;
//     - quem está no conflito recebe os dias de sobreposição.
//  4. O gestor abre a página (App da Web) e escolhe Aprovar, Pedir ajuste ou
//     Reprovar. Comentário é obrigatório para ajuste e reprovação.
//     - Aprovar: volta a situação original e espelha no Calendar.
//     - Ajuste/Reprovar: a situação vira "Ajuste Solicitado"/"Reprovado"
//       e a linha deixa de contar como conflito e como saldo usado.
//     O dev recebe a decisão e os envolvidos são atualizados.
// =========================================================================

const CONFIG_APROVACAO = {
  abaControle: "Controle Consolidado - 2026",
  abaContatos: "Contatos",
  abaAprovacoes: "Aprovações",
  // Situações que NÃO entram na checagem de conflito (sem acento, maiúsculas)
  situacoesIgnoradas: ["FINALIZADA", "PAGA", "REPROVADO", "AJUSTE SOLICITADO"],
  janelaBuscaSugestaoDias: 120
};

// E-mail do gestor vem de EMAIL_LIDER (Código.js); lido em tempo de execução
function emailGestor_() { return EMAIL_LIDER; }

const STATUS_APR = {
  PENDENTE: "Pendente Aprovação",
  APROVADO: "Aprovado",
  REPROVADO: "Reprovado",
  AJUSTE: "Ajuste Solicitado"
};
const PREFIXO_PENDENTE = STATUS_APR.PENDENTE + " | ";

// Colunas da aba Controle Consolidado (1 = A)
const COL_CTRL = { NOME: 2, SITUACAO: 5, INICIO: 10, FIM: 11, RETORNO_AFAST: 14, STATUS: 15 };

// Colunas da aba Aprovações (1 = A)
const COL_APR = {
  ID: 1, DATA_SOL: 2, NOME: 3, INICIO: 4, FIM: 5, DIAS: 6, CONFLITOS: 7,
  LINHA: 8, STATUS: 9, DECIDIDO_POR: 10, DATA_DECISAO: 11, COMENTARIO: 12,
  TOKEN: 13, SITUACAO_ORIGINAL: 14
};

// =========================================================================
// 1. CONFIGURAÇÃO INICIAL (menu Apoio Gestão > Configurar aprovação de conflitos)
// =========================================================================
function configurarAprovacaoFerias() {
  const ss = planilha_();
  PropertiesService.getScriptProperties().setProperty("ID_PLANILHA", ss.getId());

  let contatos = ss.getSheetByName(CONFIG_APROVACAO.abaContatos);
  if (!contatos) {
    contatos = ss.insertSheet(CONFIG_APROVACAO.abaContatos);
    contatos.appendRow(["Nome", "Email"]);
    const time = ss.getSheetByName("TIME");
    if (time && time.getLastRow() > 1) {
      const nomes = time.getRange(2, 1, time.getLastRow() - 1, 1).getValues()
        .filter(r => r[0])
        .map(r => [String(r[0]).trim().toUpperCase(), ""]);
      if (nomes.length) contatos.getRange(2, 1, nomes.length, 2).setValues(nomes);
    }
    contatos.getRange(1, 1, 1, 2).setFontWeight("bold");
    contatos.setFrozenRows(1);
  }

  obterAbaAprovacoes_();

  const msg = "Abas 'Contatos' e 'Aprovações' prontas.\n\nPreencha os e-mails na aba 'Contatos'." +
    (urlWebApp_() ? "" : "\n\n⚠️ URL do App da Web não encontrada: confira o deploy no GitHub.");
  try { SpreadsheetApp.getUi().alert("Aprovação de conflitos", msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }
}

// =========================================================================
// 2. PONTO DE INTEGRAÇÃO (chamado por salvarSugestaoDev após gravar a linha)
//    Retorno: { status: "sem_conflito" | "pendente" | "ignorado", mensagem }
// =========================================================================
function processarLancamentoFerias(linha) {
  const ss = planilha_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("ID_PLANILHA")) props.setProperty("ID_PLANILHA", ss.getId());

  const controle = ss.getSheetByName(CONFIG_APROVACAO.abaControle);
  const row = controle.getRange(linha, 1, 1, controle.getLastColumn()).getValues()[0];

  const nome = String(row[COL_CTRL.NOME - 1] || "").trim();
  const inicio = parseData_(row[COL_CTRL.INICIO - 1]);
  const fim = parseData_(row[COL_CTRL.FIM - 1]);
  if (!nome || !inicio || !fim) return { status: "ignorado" };

  const conflitos = encontrarConflitos_(nome, inicio, fim, listarPeriodosAtivos_(controle, linha));
  if (conflitos.length === 0) return { status: "sem_conflito" };

  const situacaoOriginal = String(row[COL_CTRL.SITUACAO - 1] || "").trim();
  controle.getRange(linha, COL_CTRL.SITUACAO).setValue(PREFIXO_PENDENTE + situacaoOriginal);
  controle.getRange(linha, COL_CTRL.STATUS).setValue("Aguardando aprovação do gestor");
  controle.getRange(linha, 1, 1, 17).setBackground("#fff3e0");

  const sol = registrarSolicitacao_(linha, nome, inicio, fim, conflitos, situacaoOriginal);

  notificarGestor_(sol, conflitos);
  notificarSolicitante_(sol, conflitos);
  notificarEnvolvidos_(sol, conflitos);

  const nomes = agruparPorNome_(conflitos).map(g => g.nome).join(", ");
  return {
    status: "pendente",
    id: sol.id,
    mensagem: "⚠️ Férias registradas, mas coincidem com: " + nomes +
      ". O lançamento foi enviado para aprovação do gestor e só vai para a agenda depois de aprovado. Você receberá a resposta por e-mail."
  };
}

// Teste sem enviar e-mail: mostra no log os conflitos de uma linha
function simularConflitos(linha) {
  const controle = planilha_().getSheetByName(CONFIG_APROVACAO.abaControle);
  const row = controle.getRange(linha, 1, 1, controle.getLastColumn()).getValues()[0];
  const nome = row[COL_CTRL.NOME - 1];
  const ini = parseData_(row[COL_CTRL.INICIO - 1]);
  const fim = parseData_(row[COL_CTRL.FIM - 1]);
  const periodos = listarPeriodosAtivos_(controle, linha);
  const conflitos = encontrarConflitos_(nome, ini, fim, periodos);
  Logger.log(nome + " " + fmt_(ini) + " a " + fmt_(fim));
  conflitos.forEach(c => Logger.log("  ↳ " + c.nome + ": " + c.dias + "d (" + fmt_(c.inicio) + " a " + fmt_(c.fim) + ")"));
  const sug = sugerirJanelaLivre_(nome, ini, fim, periodos);
  Logger.log(sug ? "Sugestão livre: " + fmt_(sug.inicio) + " a " + fmt_(sug.fim) : "Sem sugestão livre.");
}

// =========================================================================
// 3. PÁGINA DE APROVAÇÃO (App da Web)
// =========================================================================
function doGet(e) {
  return doGetAprovacao(e);
}

function doGetAprovacao(e) {
  const p = (e && e.parameter) || {};
  const sol = buscarSolicitacao_(p.id);
  if (!sol || sol.token !== p.t) {
    return paginaSimples_("Link inválido", "Solicitação não encontrada ou link inválido.");
  }
  try { validarGestor_(); } catch (err) { return paginaSimples_("Acesso negado", err.message); }

  const controle = planilha_().getSheetByName(CONFIG_APROVACAO.abaControle);
  const linha = localizarLinhaControle_(controle, sol);
  const periodos = listarPeriodosAtivos_(controle, linha);
  const conflitos = encontrarConflitos_(sol.nome, sol.inicio, sol.fim, periodos);
  const sugestao = sugerirJanelaLivre_(sol.nome, sol.inicio, sol.fim, periodos);

  const tpl = HtmlService.createTemplate(HTML_APROVACAO);
  tpl.sol = {
    id: sol.id, token: sol.token, nome: sol.nome, status: sol.status,
    inicioFmt: fmt_(sol.inicio), fimFmt: fmt_(sol.fim), dias: diasEntre_(sol.inicio, sol.fim)
  };
  tpl.podeDecidir = sol.status === STATUS_APR.PENDENTE;
  tpl.conflitos = conflitos.map(c => ({
    nome: c.nome,
    periodo: fmt_(c.inicioDele) + " a " + fmt_(c.fimDele),
    comum: fmt_(c.inicio) + " a " + fmt_(c.fim),
    dias: c.dias
  }));
  tpl.sugestao = sugestao ? {
    texto: fmt_(sugestao.inicio) + " a " + fmt_(sugestao.fim),
    iniIso: iso_(sugestao.inicio),
    fimIso: iso_(sugestao.fim)
  } : null;

  return tpl.evaluate()
    .setTitle("Aprovação de Férias")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Chamado pela página via google.script.run
function registrarDecisao(id, token, acao, comentario, novoInicioIso, novoFimIso) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    validarGestor_();
    const sol = buscarSolicitacao_(id);
    if (!sol || sol.token !== token) throw new Error("Solicitação não encontrada.");
    if (sol.status !== STATUS_APR.PENDENTE) return "Esta solicitação já foi decidida (" + sol.status + ").";

    const mapa = { APROVAR: STATUS_APR.APROVADO, REPROVAR: STATUS_APR.REPROVADO, AJUSTAR: STATUS_APR.AJUSTE };
    const novoStatus = mapa[acao];
    if (!novoStatus) throw new Error("Ação inválida.");
    comentario = String(comentario || "").trim();
    if (acao !== "APROVAR" && !comentario) throw new Error("Comentário obrigatório para ajuste ou reprovação.");

    let novasDatas = null;
    if (acao === "AJUSTAR") {
      const ni = parseData_(novoInicioIso), nf = parseData_(novoFimIso);
      if (ni && nf && nf >= ni) novasDatas = { inicio: ni, fim: nf };
    }

    // Aba Aprovações
    sol.aba.getRange(sol.linhaAba, COL_APR.STATUS, 1, 4).setValues([[
      novoStatus, Session.getActiveUser().getEmail(), new Date(), comentario
    ]]);

    // Aba Controle
    const controle = planilha_().getSheetByName(CONFIG_APROVACAO.abaControle);
    const linha = localizarLinhaControle_(controle, sol);
    if (linha) {
      if (acao === "APROVAR") {
        controle.getRange(linha, COL_CTRL.SITUACAO).setValue(sol.situacaoOriginal || "Sugestão do Dev (Padrão)");
        controle.getRange(linha, 1, 1, 17).setBackground("#fffde7");
        aoAprovarFerias_(controle, linha, sol);
      } else {
        controle.getRange(linha, COL_CTRL.SITUACAO).setValue(novoStatus);
        controle.getRange(linha, COL_CTRL.STATUS).setValue(novoStatus + " pelo gestor em " + fmt_(new Date()));
        controle.getRange(linha, 1, 1, 17).setBackground("#eeeeee");
      }
    }

    const conflitos = encontrarConflitos_(sol.nome, sol.inicio, sol.fim, listarPeriodosAtivos_(controle, linha));
    notificarDecisaoSolicitante_(sol, acao, comentario, novasDatas, conflitos);
    notificarDecisaoEnvolvidos_(sol, acao, conflitos);

    try { atualizarAbaControleConflitosSemEmail_(); } catch (e) { Logger.log(e); }

    return "✅ Decisão registrada: " + novoStatus + ". O colaborador e os envolvidos foram avisados.";
  } finally {
    lock.releaseLock();
  }
}

// Ao aprovar: espelha no Calendar do mesmo jeito que o painel faz
function aoAprovarFerias_(controle, linha, sol) {
  const r = controle.getRange(linha, 1, 1, 17).getValues()[0];
  const nome = String(r[COL_CTRL.NOME - 1] || "").toUpperCase().trim();
  const ini = parseData_(r[COL_CTRL.INICIO - 1]);
  const fim = parseData_(r[COL_CTRL.FIM - 1]);
  const retornoAfast = parseData_(r[COL_CTRL.RETORNO_AFAST - 1]);

  const m = String(sol.situacaoOriginal || "").match(/\(([^)]+)\)/);
  const opcao = m && m[1].toUpperCase() !== "PADRÃO" && m[1].toUpperCase() !== "PADRAO" ? m[1] : "";

  const titulo = "Ausência: " + nome + " [Férias" + (opcao ? " + " + opcao : "") + "]";
  const ok = injetarEventoNoCalendarioGoogle(titulo, iso_(ini), iso_(retornoAfast || fim));
  controle.getRange(linha, COL_CTRL.STATUS).setValue(ok
    ? "Espelhado no Calendar"
    : "Aprovado (falha ao gravar na agenda)");
}

// =========================================================================
// 4. REGRAS DE CONFLITO
// =========================================================================
function listarPeriodosAtivos_(sheet, linhaIgnorar) {
  const dados = sheet.getDataRange().getValues();
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    const linha = i + 1;
    if (linha === linhaIgnorar) continue;
    const nome = dados[i][COL_CTRL.NOME - 1];
    const situacao = normalizar_(dados[i][COL_CTRL.SITUACAO - 1]);
    const ini = parseData_(dados[i][COL_CTRL.INICIO - 1]);
    const fim = parseData_(dados[i][COL_CTRL.FIM - 1]);
    if (!nome || !ini || !fim || fim < hoje) continue;
    if (situacaoIgnorada_(situacao)) continue;
    lista.push({ linha: linha, nome: String(nome).trim(), inicio: ini, fim: fim });
  }
  return lista;
}

function situacaoIgnorada_(situacaoNormalizada) {
  const s = normalizar_(situacaoNormalizada);
  return CONFIG_APROVACAO.situacoesIgnoradas.some(x => s.indexOf(x) !== -1 && s.indexOf("PENDENTE") !== 0);
}

function encontrarConflitos_(nome, ini, fim, periodos) {
  const alvo = normalizar_(nome);
  const vistos = {};
  const conflitos = [];
  periodos.forEach(p => {
    if (normalizar_(p.nome) === alvo) return;
    const chave = normalizar_(p.nome) + "|" + p.inicio.getTime() + "|" + p.fim.getTime();
    if (vistos[chave]) return; // linhas duplicadas da importação
    const s = sobreposicao_(ini, fim, p.inicio, p.fim);
    if (!s) return;
    vistos[chave] = true;
    conflitos.push({
      nome: p.nome, linha: p.linha,
      inicioDele: p.inicio, fimDele: p.fim,
      inicio: s.inicio, fim: s.fim, dias: s.dias
    });
  });
  return conflitos.sort((a, b) => b.dias - a.dias);
}

function sobreposicao_(i1, f1, i2, f2) {
  const ini = i1 > i2 ? i1 : i2;
  const fim = f1 < f2 ? f1 : f2;
  if (ini > fim) return null;
  return { inicio: ini, fim: fim, dias: diasEntre_(ini, fim) };
}

// Próxima janela com a mesma duração e sem conflito, começando de segunda a quarta
function sugerirJanelaLivre_(nome, ini, fim, periodos) {
  if (!ini || !fim) return null;
  const duracao = diasEntre_(ini, fim);
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const base = ini > hoje ? ini : hoje;
  for (let d = 1; d <= CONFIG_APROVACAO.janelaBuscaSugestaoDias; d++) {
    const cIni = new Date(base.getFullYear(), base.getMonth(), base.getDate() + d);
    const dow = cIni.getDay();
    if (dow < 1 || dow > 3) continue;
    const cFim = new Date(cIni.getFullYear(), cIni.getMonth(), cIni.getDate() + duracao - 1);
    const choqueProprio = periodos.some(p =>
      normalizar_(p.nome) === normalizar_(nome) && sobreposicao_(cIni, cFim, p.inicio, p.fim));
    if (!choqueProprio && encontrarConflitos_(nome, cIni, cFim, periodos).length === 0) return { inicio: cIni, fim: cFim };
  }
  return null;
}

// =========================================================================
// 5. E-MAILS
// =========================================================================
function notificarGestor_(sol, conflitos) {
  const url = urlWebApp_();
  const link = url ? url + "?id=" + encodeURIComponent(sol.id) + "&t=" + sol.token : "";
  const semEmail = [sol.nome].concat(conflitos.map(c => c.nome)).filter(n => !obterEmail_(n));
  const unicosSemEmail = semEmail.filter((n, i) => semEmail.indexOf(n) === i);

  const corpo =
    "<p><b>" + esc_(sol.nome) + "</b> lançou férias de <b>" + fmt_(sol.inicio) + " a " + fmt_(sol.fim) +
    "</b> (" + sol.dias + " dias), em conflito com:</p>" +
    tabelaConflitosHtml_(conflitos) +
    (link
      ? "<p style='text-align:center;margin:24px 0'><a href='" + link + "' style='background:#1a73e8;color:#fff;padding:12px 22px;border-radius:6px;text-decoration:none;font-weight:bold'>Analisar e decidir</a></p>"
      : "<p style='color:#d93025'><b>URL do App da Web não configurada:</b> registre a decisão manualmente na aba Aprovações.</p>") +
    (unicosSemEmail.length
      ? "<p style='font-size:12px;color:#b06000'>⚠️ Sem e-mail cadastrado na aba Contatos: " + unicosSemEmail.map(esc_).join(", ") + "</p>"
      : "");

  enviarEmail_(emailGestor_(),
    "⚠️ Aprovação necessária: férias de " + sol.nome + " com conflito",
    "Aprovação de férias", corpo);
}

function notificarSolicitante_(sol, conflitos) {
  const corpo =
    "<p>Olá, <b>" + esc_(primeiroNome_(sol.nome)) + "</b>!</p>" +
    "<p>Seu lançamento de férias de <b>" + fmt_(sol.inicio) + " a " + fmt_(sol.fim) +
    "</b> foi registrado, mas coincide com férias de outras pessoas do time:</p>" +
    tabelaConflitosHtml_(conflitos) +
    "<p>Ele ficou <b>pendente de aprovação do gestor</b> e só vai para a agenda depois de aprovado. " +
    "Você vai receber a resposta por e-mail. Até lá, evite comprar passagens ou fechar reservas.</p>";
  enviarEmail_(obterEmail_(sol.nome),
    "⏳ Suas férias estão aguardando aprovação (" + fmt_(sol.inicio) + " a " + fmt_(sol.fim) + ")",
    "Férias pendentes de aprovação", corpo);
}

function notificarEnvolvidos_(sol, conflitos) {
  agruparPorNome_(conflitos).forEach(g => {
    const itens = g.itens.map(c =>
      "<li>Suas férias de <b>" + fmt_(c.inicioDele) + " a " + fmt_(c.fimDele) + "</b>: " +
      "<b>" + c.dias + " dia(s)</b> em comum (" + fmt_(c.inicio) + " a " + fmt_(c.fim) + ")</li>").join("");
    const corpo =
      "<p>Olá, <b>" + esc_(primeiroNome_(g.nome)) + "</b>!</p>" +
      "<p><b>" + esc_(sol.nome) + "</b> lançou férias de <b>" + fmt_(sol.inicio) + " a " + fmt_(sol.fim) +
      "</b>, que coincidem com as suas:</p><ul>" + itens + "</ul>" +
      "<p>O lançamento está aguardando aprovação do gestor. Você não precisa fazer nada agora. " +
      "Se for preciso ajustar, o gestor vai alinhar com vocês.</p>";
    enviarEmail_(obterEmail_(g.nome),
      "📅 Conflito de férias com " + sol.nome,
      "Aviso de sobreposição de férias", corpo);
  });
}

function notificarDecisaoSolicitante_(sol, acao, comentario, novasDatas, conflitos) {
  const periodo = fmt_(sol.inicio) + " a " + fmt_(sol.fim);
  const blocoComentario = comentario
    ? "<div style='background:#f8f9fa;border-left:4px solid #1a73e8;padding:10px 12px;margin:12px 0'><b>Comentário do gestor:</b><br>" +
      esc_(comentario).replace(/\n/g, "<br>") + "</div>"
    : "";
  let titulo, assunto, texto;

  if (acao === "APROVAR") {
    titulo = "Férias aprovadas ✅";
    assunto = "✅ Férias aprovadas (" + periodo + ")";
    texto = "<p>Suas férias de <b>" + periodo + "</b> foram <b>aprovadas</b> e já estão na agenda do time.</p>" + blocoComentario +
      (conflitos.length
        ? "<p>Elas vão coincidir com as férias abaixo. Combine a cobertura das atividades antes de sair:</p>" + tabelaConflitosHtml_(conflitos)
        : "");
  } else if (acao === "AJUSTAR") {
    titulo = "Ajuste de datas solicitado 🔁";
    assunto = "🔁 Ajuste solicitado nas suas férias (" + periodo + ")";
    texto = "<p>O gestor pediu um <b>ajuste</b> nas suas férias de <b>" + periodo + "</b>.</p>" + blocoComentario +
      (novasDatas
        ? "<p>💡 Sugestão de período: <b>" + fmt_(novasDatas.inicio) + " a " + fmt_(novasDatas.fim) + "</b></p>"
        : "") +
      "<p>Faça um novo lançamento pelo painel com as datas ajustadas.</p>";
  } else {
    titulo = "Férias não aprovadas ❌";
    assunto = "❌ Férias não aprovadas (" + periodo + ")";
    texto = "<p>Suas férias de <b>" + periodo + "</b> <b>não foram aprovadas</b>.</p>" + blocoComentario +
      "<p>Se quiser, converse com o gestor e faça um novo lançamento pelo painel.</p>";
  }

  enviarEmail_(obterEmail_(sol.nome), assunto, titulo,
    "<p>Olá, <b>" + esc_(primeiroNome_(sol.nome)) + "</b>!</p>" + texto);
}

function notificarDecisaoEnvolvidos_(sol, acao, conflitos) {
  const periodo = fmt_(sol.inicio) + " a " + fmt_(sol.fim);
  agruparPorNome_(conflitos).forEach(g => {
    const totalDias = g.itens.reduce((s, c) => s + c.dias, 0);
    let texto;
    if (acao === "APROVAR") {
      texto = "<p>As férias de <b>" + esc_(sol.nome) + "</b> (" + periodo + ") foram <b>aprovadas</b> e vão coincidir " +
        "com as suas em <b>" + totalDias + " dia(s)</b>. Combinem a cobertura das atividades.</p>";
    } else if (acao === "AJUSTAR") {
      texto = "<p>O lançamento de <b>" + esc_(sol.nome) + "</b> (" + periodo + ") que coincidia com suas férias " +
        "foi <b>devolvido para ajuste de datas</b>. Por enquanto, o conflito foi desfeito.</p>";
    } else {
      texto = "<p>O lançamento de <b>" + esc_(sol.nome) + "</b> (" + periodo + ") que coincidia com suas férias " +
        "<b>não foi aprovado</b>. O conflito foi desfeito.</p>";
    }
    enviarEmail_(obterEmail_(g.nome),
      "📅 Atualização do conflito de férias com " + sol.nome,
      "Atualização de conflito de férias",
      "<p>Olá, <b>" + esc_(primeiroNome_(g.nome)) + "</b>!</p>" + texto);
  });
}

function tabelaConflitosHtml_(conflitos) {
  const th = "style='text-align:left;padding:8px;background:#f8f9fa;border-bottom:1px solid #e0e0e0'";
  const td = "style='padding:8px;border-bottom:1px solid #f0f0f0'";
  const linhas = conflitos.map(c =>
    "<tr><td " + td + ">" + esc_(c.nome) + "</td>" +
    "<td " + td + ">" + fmt_(c.inicioDele) + " a " + fmt_(c.fimDele) + "</td>" +
    "<td " + td + ">" + fmt_(c.inicio) + " a " + fmt_(c.fim) + "</td>" +
    "<td " + td + " align='center'><b>" + c.dias + "</b></td></tr>").join("");
  return "<table style='width:100%;border-collapse:collapse;font-size:13px;margin:12px 0'>" +
    "<tr><th " + th + ">Colaborador</th><th " + th + ">Férias</th><th " + th + ">Dias em comum</th><th " + th + ">Qtd</th></tr>" +
    linhas + "</table>";
}

function enviarEmail_(para, assunto, titulo, corpoHtml) {
  if (!para) { Logger.log("Sem e-mail para envio: " + assunto); return; }
  const html =
    "<div style='max-width:600px;margin:0 auto;font-family:Arial,sans-serif;border:1px solid #e0e0e0;border-radius:12px;overflow:hidden;background:#fff'>" +
    "<div style='padding:16px 20px;border-bottom:1px solid #f0f0f0'>" +
    "<h2 style='margin:0;font-size:18px;color:#202124'>" + titulo + "</h2>" +
    "<p style='margin:4px 0 0;font-size:13px;color:#5f6368'>Squad Novas Apurações · Controle de Férias</p></div>" +
    "<div style='padding:20px;font-size:14px;color:#3c4043;line-height:1.5'>" + corpoHtml + "</div></div>";
  try {
    MailApp.sendEmail({
      to: para,
      subject: assunto,
      body: corpoHtml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim(),
      htmlBody: html
    });
  } catch (e) {
    Logger.log("Falha ao enviar e-mail para " + para + ": " + e.message);
  }
}

// =========================================================================
// 6. APOIO
// =========================================================================
// No painel a planilha ativa existe; no App da Web usa o ID salvo
function planilha_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const id = PropertiesService.getScriptProperties().getProperty("ID_PLANILHA");
  if (!id) throw new Error("ID da planilha não configurado. Rode 'Configurar aprovação de conflitos' no menu Apoio Gestão.");
  return SpreadsheetApp.openById(id);
}

// URL gerada no deploy (ConfigDeploy.js) ou definida na propriedade URL_WEBAPP
function urlWebApp_() {
  const prop = PropertiesService.getScriptProperties().getProperty("URL_WEBAPP");
  if (prop) return prop;
  return typeof URL_WEBAPP_APROVACAO !== "undefined" ? URL_WEBAPP_APROVACAO : "";
}

function obterAbaAprovacoes_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(CONFIG_APROVACAO.abaAprovacoes);
  if (!aba) {
    aba = ss.insertSheet(CONFIG_APROVACAO.abaAprovacoes);
    aba.appendRow(["ID", "Data Solicitação", "Colaborador", "Início", "Fim", "Dias", "Conflitos",
      "Linha Controle", "Status", "Decidido por", "Data Decisão", "Comentário", "Token", "Situação Original"]);
    aba.getRange(1, 1, 1, 14).setFontWeight("bold");
    aba.setFrozenRows(1);
    aba.hideColumns(COL_APR.TOKEN);
  }
  return aba;
}

function registrarSolicitacao_(linha, nome, inicio, fim, conflitos, situacaoOriginal) {
  const aba = obterAbaAprovacoes_();
  const token = Utilities.getUuid();
  const id = "FER-" + Utilities.formatDate(new Date(), tz_(), "yyyyMMdd-HHmmss") + "-" + token.slice(0, 4);
  const dias = diasEntre_(inicio, fim);
  const resumo = conflitos.map(c => c.nome + " (" + c.dias + "d)").join(", ");
  aba.appendRow([id, new Date(), nome, inicio, fim, dias, resumo, linha, STATUS_APR.PENDENTE, "", "", "", token, situacaoOriginal]);
  return { id: id, token: token, nome: nome, inicio: inicio, fim: fim, dias: dias, linha: linha, situacaoOriginal: situacaoOriginal };
}

function buscarSolicitacao_(id) {
  if (!id) return null;
  const aba = obterAbaAprovacoes_();
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][COL_APR.ID - 1]) === String(id)) {
      const r = dados[i];
      return {
        aba: aba, linhaAba: i + 1,
        id: r[COL_APR.ID - 1], nome: String(r[COL_APR.NOME - 1]),
        inicio: parseData_(r[COL_APR.INICIO - 1]), fim: parseData_(r[COL_APR.FIM - 1]),
        dias: r[COL_APR.DIAS - 1], linha: r[COL_APR.LINHA - 1],
        status: r[COL_APR.STATUS - 1], token: String(r[COL_APR.TOKEN - 1]),
        situacaoOriginal: String(r[COL_APR.SITUACAO_ORIGINAL - 1] || "")
      };
    }
  }
  return null;
}

// Localiza a linha pelo nome + datas (caso a aba tenha sido reordenada)
function localizarLinhaControle_(sheet, sol) {
  const dados = sheet.getDataRange().getValues();
  const confere = i => {
    const r = dados[i];
    const ini = parseData_(r[COL_CTRL.INICIO - 1]);
    const fim = parseData_(r[COL_CTRL.FIM - 1]);
    return normalizar_(r[COL_CTRL.NOME - 1]) === normalizar_(sol.nome) &&
      ini && fim && ini.getTime() === sol.inicio.getTime() && fim.getTime() === sol.fim.getTime();
  };
  const idx = Number(sol.linha) - 1;
  if (idx >= 1 && idx < dados.length && confere(idx)) return idx + 1;
  for (let i = 1; i < dados.length; i++) if (confere(i)) return i + 1;
  return null;
}

let _cacheEmails = null;
function obterEmail_(nome) {
  if (!_cacheEmails) {
    _cacheEmails = {};
    const aba = planilha_().getSheetByName(CONFIG_APROVACAO.abaContatos);
    if (aba) {
      aba.getDataRange().getValues().slice(1).forEach(r => {
        if (r[0] && r[1]) _cacheEmails[normalizar_(r[0])] = String(r[1]).trim();
      });
    }
  }
  return _cacheEmails[normalizar_(nome)] || "";
}

function validarGestor_() {
  const email = String(Session.getActiveUser().getEmail() || "").toLowerCase();
  if (email !== String(emailGestor_()).toLowerCase()) {
    throw new Error("Apenas o gestor pode registrar a decisão.");
  }
}

function agruparPorNome_(conflitos) {
  const mapa = {};
  conflitos.forEach(c => {
    const k = normalizar_(c.nome);
    if (!mapa[k]) mapa[k] = { nome: c.nome, itens: [] };
    mapa[k].itens.push(c);
  });
  return Object.keys(mapa).map(k => mapa[k]);
}

function parseData_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v || "").trim();
  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function diasEntre_(a, b) { return Math.round((b - a) / 86400000) + 1; } // inclusivo
function tz_() { return Session.getScriptTimeZone() || "America/Sao_Paulo"; }
function fmt_(d) { return d ? Utilities.formatDate(d, tz_(), "dd/MM/yyyy") : "-"; }
function iso_(d) { return d ? Utilities.formatDate(d, tz_(), "yyyy-MM-dd") : ""; }
function normalizar_(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toUpperCase();
}
function primeiroNome_(nome) {
  const p = String(nome || "").trim().split(/\s+/)[0] || "";
  return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
}
function esc_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function paginaSimples_(titulo, msg) {
  return HtmlService.createHtmlOutput(
    "<div style='font-family:Arial,sans-serif;max-width:520px;margin:60px auto;text-align:center'>" +
    "<h2>" + esc_(titulo) + "</h2><p>" + esc_(msg) + "</p></div>").setTitle(titulo);
}

// =========================================================================
// 7. HTML DA PÁGINA DE APROVAÇÃO
// =========================================================================
const HTML_APROVACAO = [
  "<!DOCTYPE html><html><head><base target='_top'>",
  "<style>",
  "body{font-family:Arial,sans-serif;background:#f4f5f7;margin:0;padding:24px;color:#202124}",
  ".card{max-width:680px;margin:0 auto;background:#fff;border:1px solid #e0e0e0;border-radius:12px;padding:24px}",
  "h2{margin:0 0 4px}.sub{color:#5f6368;font-size:13px;margin:0 0 16px}",
  "table{width:100%;border-collapse:collapse;font-size:13px;margin:12px 0}",
  "th,td{border-bottom:1px solid #eee;padding:8px;text-align:left}th{background:#f8f9fa}",
  ".sug{background:#e8f0fe;padding:10px 12px;border-radius:8px;font-size:13px}",
  "label{font-size:13px;font-weight:bold;display:block;margin-top:14px;margin-bottom:6px}",
  "textarea{width:100%;min-height:80px;box-sizing:border-box;padding:8px;border:1px solid #ccc;border-radius:6px;font-family:inherit}",
  "input[type=date]{padding:6px;border:1px solid #ccc;border-radius:6px}",
  ".btns{display:flex;gap:8px;margin-top:18px;flex-wrap:wrap}",
  "button{border:0;border-radius:6px;padding:10px 16px;font-weight:bold;cursor:pointer;color:#fff}",
  "button:disabled{opacity:.5;cursor:default}",
  ".ok{background:#188038}.no{background:#d93025}.adj{background:#e37400}",
  "#msg{margin-top:16px;font-weight:bold}",
  "</style></head><body><div class='card'>",
  "<h2>Solicitação de férias com conflito</h2>",
  "<p class='sub'><?= sol.id ?> · Status: <?= sol.status ?></p>",
  "<p><b><?= sol.nome ?></b>: <?= sol.inicioFmt ?> a <?= sol.fimFmt ?> (<?= sol.dias ?> dias)</p>",
  "<? if (conflitos.length) { ?>",
  "<table><tr><th>Em conflito com</th><th>Férias</th><th>Dias em comum</th><th>Qtd</th></tr>",
  "<? for (var i = 0; i < conflitos.length; i++) { var c = conflitos[i]; ?>",
  "<tr><td><?= c.nome ?></td><td><?= c.periodo ?></td><td><?= c.comum ?></td><td><b><?= c.dias ?></b></td></tr>",
  "<? } ?></table>",
  "<? } else { ?><p>✅ No momento não há mais conflito com outras férias.</p><? } ?>",
  "<? if (sugestao) { ?><div class='sug'>💡 Próxima janela livre com a mesma duração: <b><?= sugestao.texto ?></b></div><? } ?>",
  "<? if (podeDecidir) { ?>",
  "<input type='hidden' id='sid' value='<?= sol.id ?>'><input type='hidden' id='stk' value='<?= sol.token ?>'>",
  "<label>Comentário para o colaborador (obrigatório para ajuste ou reprovação)</label>",
  "<textarea id='coment'></textarea>",
  "<label>Datas sugeridas para o ajuste (opcional)</label>",
  "<input type='date' id='ni' value='<?= sugestao ? sugestao.iniIso : \"\" ?>'> a ",
  "<input type='date' id='nf' value='<?= sugestao ? sugestao.fimIso : \"\" ?>'>",
  "<div class='btns'>",
  "<button class='ok' onclick=\"decidir('APROVAR')\">Aprovar</button>",
  "<button class='adj' onclick=\"decidir('AJUSTAR')\">Pedir ajuste de datas</button>",
  "<button class='no' onclick=\"decidir('REPROVAR')\">Reprovar</button>",
  "</div>",
  "<? } else { ?><p><b>Esta solicitação já foi decidida.</b></p><? } ?>",
  "<div id='msg'></div></div>",
  "<script>",
  "function travar(v){var b=document.querySelectorAll('button');for(var i=0;i<b.length;i++)b[i].disabled=v;}",
  "function decidir(acao){",
  "  var c=document.getElementById('coment').value.trim();",
  "  if(acao!=='APROVAR'&&!c){alert('Escreva um comentário para o colaborador.');return;}",
  "  travar(true);document.getElementById('msg').textContent='Registrando...';",
  "  google.script.run",
  "    .withSuccessHandler(function(r){document.getElementById('msg').textContent=r;})",
  "    .withFailureHandler(function(e){document.getElementById('msg').textContent='Erro: '+e.message;travar(false);})",
  "    .registrarDecisao(document.getElementById('sid').value,document.getElementById('stk').value,acao,c,",
  "      document.getElementById('ni').value,document.getElementById('nf').value);",
  "}",
  "</script></body></html>"
].join("\n");
