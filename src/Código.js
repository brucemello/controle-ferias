// =========================================================================================
// CONFIGURAÇÕES GLOBAIS
// =========================================================================================
// Gestor, agenda, nome do time e destinatários ficam em ConfigTime.js
// (Propriedades do Script) — configure pelo dashboard em "Configuração".

// =========================================================================================
// FUNÇÃO DO MENU (SEM NECESSIDADE DE DIGITAR O E-MAIL)
// =========================================================================================
function enviarEmailConflitosLiderDireto() {
  var ui = SpreadsheetApp.getUi();
  
  if (!emailLider_()) {
    ui.alert("Configuração Pendente", "Configure o time no Dashboard de Férias (menu lateral > Configuração).", ui.ButtonSet.OK);
    return;
  }
  
  // Atualiza os conflitos antes de disparar o e-mail
  atualizarAbaControleConflitos();
  
  var enviado = enviarEmailConflitosLider(emailLider_());
  
  if (enviado) {
    ui.alert("E-mail Enviado!", "O relatório de conflitos foi enviado com sucesso para:\n" + emailLider_(), ui.ButtonSet.OK);
  } else {
    ui.alert("Sem Conflitos", "Não foram encontrados conflitos pendentes para envio no momento.", ui.ButtonSet.OK);
  }
}

// =========================================================================================
// ATUALIZAÇÃO DO MENU
// =========================================================================================
function onOpen() {
  vincularPlanilha_();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Apoio Gestão')
    .addItem('📊 Abrir Dashboard de Férias', 'abrirDashboard')
    .addItem('Abrir Painel de Controle', 'abrirPainel')
    .addItem('Abrir Manual do Usuário', 'abManual')
    .addSeparator()
    .addItem('⚡ Atualizar Controle de Conflitos', 'processarEAtualizarConflitos')
    .addItem('📧 Enviar E-mail de Conflitos ao Líder', 'enviarEmailConflitosLiderDireto') // <-- USA O E-MAIL SALVO
    .addSeparator()
    .addItem('🧪 TESTAR Sincronização (Apenas Linha 32)', 'testarSincronizacaoUnica') 
    .addItem('🔗 Sincronizar Histórico Antigo com Agenda (Lote)', 'publicarHistoricoAntigoNoCalendario') 
    .addSeparator()
    .addItem('Verificar e Criar Estrutura de Abas (Seguro)', 'inicializarTodasAbas')
    .addItem('✅ Configurar Aprovação de Conflitos', 'configurarAprovacaoFerias')
    .addToUi();
}

// Prompt interativo para solicitar o e-mail do líder e enviar o relatório
function enviarEmailConflitosLiderPrompt() {
  var ui = SpreadsheetApp.getUi();
  
  // Garante que os conflitos estejam atualizados antes do envio
  atualizarAbaControleConflitos();
  
  var resposta = ui.prompt(
    'Enviar Relatório de Conflitos',
    'Insira o endereço de e-mail do Líder / Gestor:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resposta.getSelectedButton() === ui.Button.OK) {
    var emailDestinatario = resposta.getResponseText().trim();
    if (emailDestinatario === "" || !emailDestinatario.includes("@")) {
      ui.alert("E-mail Inválido", "Por favor, insira um e-mail válido.", ui.ButtonSet.OK);
      return;
    }
    
    var enviado = enviarEmailConflitosLider(emailDestinatario);
    if (enviado) {
      ui.alert("Sucesso!", "E-mail com o relatório de conflitos enviado com sucesso para: " + emailDestinatario, ui.ButtonSet.OK);
    } else {
      ui.alert("Aviso", "Não há conflitos pendentes identificados para envio no momento.", ui.ButtonSet.OK);
    }
  }
}

// =========================================================================================
// DISPARO DE E-MAIL FORMATADO PARA O LÍDER
// =========================================================================================
function enviarEmailConflitosLider(emailDestinatario) {
  var ss = obterPlanilhaAtiva();
  var sheetConflito = ss.getSheetByName("Controle de Conflitos");
  
  if (!sheetConflito || sheetConflito.getLastRow() < 6) return false;
  
  var dados = sheetConflito.getRange(6, 1, sheetConflito.getLastRow() - 5, 7).getValues();
  var listaConflitos = [];
  
  // Filtra apenas as linhas que possuem status "Conflito"
  for (var i = 0; i < dados.length; i++) {
    var colab = dados[i][0];
    var dtIni = dados[i][1];
    var dtFim = dados[i][2];
    var situacao = dados[i][3];
    var statusConflito = dados[i][4];
    var conflitoCom = dados[i][5];
    var diasConflito = dados[i][6];
    
    if (statusConflito === "Conflito") {
      listaConflitos.push({
        colaborador: colab,
        dtIni: dtIni,
        dtFim: dtFim,
        situacao: situacao,
        conflitoCom: conflitoCom,
        dias: diasConflito
      });
    }
  }
  
  if (listaConflitos.length === 0) return false;
  
  // Monta o Corpo do E-mail em HTML
  var assunto = "⚠️ [Aviso de Gestão] Relatório de Conflitos de Férias/Ausências - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  var htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 700px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; background-color: #ffffff;">
      <h2 style="color: #1f4e78; border-bottom: 2px solid #1f4e78; padding-bottom: 10px; margin-top: 0;">
        Relatório de Conflitos de Ausências
      </h2>
      <p>Olá, Líder,</p>
      <p>Identificamos <strong>${listaConflitos.length} sobreposição(ões) de período(s)</strong> de férias/folgas na análise atual do time. Confira os detalhes abaixo:</p>
      
      <table style="width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px;">
        <thead>
          <tr style="background-color: #1f4e78; color: #ffffff; text-align: left;">
            <th style="padding: 10px; border: 1px solid #ddd;">Colaborador</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Período</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Situação</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Conflito Com</th>
            <th style="padding: 10px; border: 1px solid #ddd; text-align: center;">Dias</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  for (var k = 0; k < listaConflitos.length; k++) {
    var item = listaConflitos[k];
    var bgRow = (k % 2 === 0) ? "#f9f9f9" : "#ffffff";
    
    htmlBody += `
      <tr style="background-color: ${bgRow};">
        <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">${item.colaborador}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">${item.dtIni} a ${item.dtFim}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">${item.situacao}</td>
        <td style="padding: 10px; border: 1px solid #ddd; color: #b85450; font-weight: bold;">${item.conflitoCom}</td>
        <td style="padding: 10px; border: 1px solid #ddd; text-align: center; font-weight: bold;">${item.dias}</td>
      </tr>
    `;
  }
  
  htmlBody += `
        </tbody>
      </table>
      
      <div style="margin-top: 25px; padding: 12px; background-color: #fff3cd; border-left: 4px solid #ffc107; font-size: 13px;">
        <strong>Orientação:</strong> Recomendamos alinhar com os colaboradores listados para ajustar as datas das sugestões/agendamentos.
      </div>
      
      <p style="font-size: 11px; color: #777; margin-top: 20px;">
        Mensagem enviada automaticamente pelo sistema de Controle de Férias.
      </p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: emailDestinatario,
    subject: assunto,
    htmlBody: htmlBody
  });
  
  return true;
}

// Função auxiliar para chamar via menu com alerta visual de sucesso
function processarEAtualizarConflitos() {
  var ui = SpreadsheetApp.getUi();
  try {
    atualizarAbaControleConflitos();
    ui.alert("Atualização Concluída", "A aba 'Controle de Conflitos' foi atualizada com sucesso com os nomes e dias de sobreposição!", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Erro", "Falha ao atualizar o controle de conflitos: " + e.message, ui.ButtonSet.OK);
  }
}

// Gatilho instalável "Ao abrir": abre o dashboard de férias
function gatilhoAberturaAutomatica() {
  try { abrirDashboard(); } catch (e) {}
}

function abrirPainel() {
  var html = HtmlService.createHtmlOutputFromFile('contole_ferias')
      .setWidth(450)
      .setHeight(700); 
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function abManual() {
  var html = HtmlService.createHtmlOutputFromFile('contole_ferias')
      .setWidth(500)
      .setHeight(650);
  html.append("<script>window.onload = function() { alternarAba('manual'); }</script>");
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// Pega dinamicamente a planilha aberta no navegador
// No App da Web de aprovação não há planilha aberta: planilha_() usa o ID salvo
function obterPlanilhaAtiva() {
  return planilha_();
}

// =========================================================================================
// FUNÇÕES DE BUSCA E LEITURA DE DADOS
// =========================================================================================
function buscarNomesTime() {
  try {
    var ss = obterPlanilhaAtiva();
    
    // Busca a aba "TIME" ignorando diferenças de maiúsculas/minúsculas e espaços
    var sheets = ss.getSheets();
    var sheetTime = null;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim().toUpperCase() === "TIME") {
        sheetTime = sheets[i];
        break;
      }
    }
    
    if (!sheetTime || sheetTime.getLastRow() < 2) return [];
    
    var dados = sheetTime.getRange(2, 1, sheetTime.getLastRow() - 1, 1).getValues();
    var nomes = [];
    for (var j = 0; j < dados.length; j++) {
      var val = dados[j][0];
      if (val !== "" && val !== undefined && val !== null) {
        nomes.push(val.toString().trim());
      }
    }
    return nomes.sort();
  } catch (e) { 
    throw new Error("Falha ao ler aba TIME: " + e.message); 
  }
}

function buscarAfastamentos() {
  try {
    var ss = obterPlanilhaAtiva();
    var sheetAfast = ss.getSheetByName("Afastamento");
    if (!sheetAfast) return [];
    var dados = sheetAfast.getRange("A2:A" + sheetAfast.getLastRow()).getValues();
    var lista = [];
    for (var i = 0; i < dados.length; i++) {
      if (dados[i][0] !== "" && dados[i][0] !== undefined && dados[i][0] !== null) {
        lista.push(dados[i][0].toString().trim());
      }
    }
    return lista;
  } catch (e) { return []; }
}

function buscarFeriasAgendadasColaborador(nomePesquisa) {
  try {
    if (!nomePesquisa) return [];
    var nomeColab = nomePesquisa.toUpperCase().trim();
    var ss = obterPlanilhaAtiva();
    var sheet = ss.getSheetByName("Controle Consolidado - 2026");
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    var registros = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
    var feriasEncontradas = [];
    var periodosVerificados = []; 
    
    for (var i = 0; i < registros.length; i++) {
      if (!registros[i][1]) continue;
      
      var n = registros[i][1].toString().toUpperCase().trim();
      var statusLinha = registros[i][8] ? registros[i][8].toString().toUpperCase().trim() : ""; 
      var situacao = registros[i][4] ? registros[i][4].toString().toUpperCase().trim() : "";   
      var dIni = registros[i][9];                                        
      var dFim = registros[i][10];                                       
      
      if (n === nomeColab && statusLinha === "AGENDADOS" && dIni && dFim && !isNaN(new Date(dIni).getTime()) && !isNaN(new Date(dFim).getTime())) {
        if (situacao.includes("+") || situacao.includes("SOLICITADO") || situacao.includes("SUGESTÃO") || 
            situacao.includes("GIFTDAY") || situacao.includes("RESET") || situacao.includes("MVT")) {
          continue; 
        }
        
        var fIni = Utilities.formatDate(new Date(dIni), Session.getScriptTimeZone(), "dd/MM/yyyy");
        var fFim = Utilities.formatDate(new Date(dFim), Session.getScriptTimeZone(), "dd/MM/yyyy");
        var chaveUnicaPeriodo = fIni + "_" + fFim;
        
        if (periodosVerificados.indexOf(chaveUnicaPeriodo) !== -1) continue;
        
        var isoIni = Utilities.formatDate(new Date(dIni), Session.getScriptTimeZone(), "yyyy-MM-dd");
        var isoFim = Utilities.formatDate(new Date(dFim), Session.getScriptTimeZone(), "yyyy-MM-dd");
        
        feriasEncontradas.push({
          linhaIndex: i + 2, 
          textoExibicao: "Férias: de " + fIni + " a " + fFim,
          dataInicioIso: isoIni,
          dataFimIso: isoFim,
          retornoPadraoOriginal: registros[i][12] ? Utilities.formatDate(new Date(registros[i][12]), Session.getScriptTimeZone(), "yyyy-MM-dd") : ""
        });
        
        periodosVerificados.push(chaveUnicaPeriodo);
      }
    }
    return feriasEncontradas;
  } catch(e) { return []; }
}

function consultingSaldosColaborador(nomePesquisa) {
  return consultarSaldosColaborador(nomePesquisa);
}

// Saldos de benefícios do ano corrente:
//  - GiftDay: 2/3/4/5 dias no ano em que completa 5/10/15/20 anos de casa (aba Admisssão)
//  - ResetDay: 1 por ano para todos
//  - MVT: conforme a aba MVT (Quantidade = 1 dá direito; Utilizou = Sim consome)
function consultarSaldosColaborador(nomePesquisa) {
  try {
    var resultado = { giftTotal: 0, giftUsado: 0, giftSaldo: 0, resetTotal: 1, resetUsado: 0, resetSaldo: 1, resetUsadoEm: "",
                      mvtTotal: 0, mvtUsado: 0, mvtSaldo: 0, anosEmpresa: 0, admissao: "" };
    if (!nomePesquisa) return resultado;
    var nomeColab = normalizar_(nomePesquisa);
    var ss = obterPlanilhaAtiva();
    var anoAtual = new Date().getFullYear();
    var doAno = function(v) { var d = parseData_(v); return !!d && d.getFullYear() === anoAtual; };

    var sheetAdmissao = ss.getSheetByName("Admisssão");
    if (sheetAdmissao && sheetAdmissao.getLastRow() > 1) {
      var dadosAdmissao = sheetAdmissao.getRange(2, 1, sheetAdmissao.getLastRow() - 1, 2).getValues();
      for (var j = 0; j < dadosAdmissao.length; j++) {
        if (dadosAdmissao[j][0] && normalizar_(dadosAdmissao[j][0]) === nomeColab) {
          var dataAdmissao = parseData_(dadosAdmissao[j][1]);
          if (dataAdmissao) {
            var anos = anoAtual - dataAdmissao.getFullYear();
            resultado.anosEmpresa = anos;
            resultado.admissao = iso_(dataAdmissao);
            if (anos === 5) resultado.giftTotal = 2;
            else if (anos === 10) resultado.giftTotal = 3;
            else if (anos === 15) resultado.giftTotal = 4;
            else if (anos === 20) resultado.giftTotal = 5;
          }
          break;
        }
      }
    }

    var sheetMVT = ss.getSheetByName("MVT");
    if (sheetMVT && sheetMVT.getLastRow() > 1) {
      var dadosMVT = sheetMVT.getRange(2, 1, sheetMVT.getLastRow() - 1, 5).getValues();
      for (var m = 0; m < dadosMVT.length; m++) {
        if (dadosMVT[m][0] && normalizar_(dadosMVT[m][0]) === nomeColab) {
          var qtdDireito = parseInt(dadosMVT[m][3], 10);
          var utilizouStr = normalizar_(dadosMVT[m][4] || "NAO");
          if (!isNaN(qtdDireito) && qtdDireito >= 1) {
            resultado.mvtTotal += qtdDireito;
            if (utilizouStr === "SIM") resultado.mvtUsado += qtdDireito;
          }
        }
      }
    }
    resultado.mvtSaldo = Math.max(0, resultado.mvtTotal - resultado.mvtUsado);

    var sheetFerias = ss.getSheetByName("Controle Consolidado - 2026");
    if (sheetFerias && sheetFerias.getLastRow() > 1) {
      var regFerias = sheetFerias.getRange(2, 1, sheetFerias.getLastRow() - 1, 12).getValues();
      for (var i = 0; i < regFerias.length; i++) {
        if (regFerias[i][1] && normalizar_(regFerias[i][1]) === nomeColab) {
          var sit = normalizar_(regFerias[i][4]);
          if (sit.indexOf("REPROVADO") === 0 || sit.indexOf("AJUSTE") === 0) continue;
          if (!doAno(regFerias[i][9])) continue;
          if (sit.indexOf("RESET") !== -1) { resultado.resetUsado = 1; resultado.resetUsadoEm = iso_(parseData_(regFerias[i][9])); }
          if (sit.indexOf("GIFT") !== -1) resultado.giftUsado += 1;
        }
      }
    }

    var sheetFolgas = ss.getSheetByName("Folgas e Emendas - 2026");
    if (sheetFolgas && sheetFolgas.getLastRow() > 1) {
      var regFolgas = sheetFolgas.getRange(2, 1, sheetFolgas.getLastRow() - 1, 5).getValues();
      for (var k = 0; k < regFolgas.length; k++) {
        if (regFolgas[k][0] && normalizar_(regFolgas[k][0]) === nomeColab) {
          if (!doAno(regFolgas[k][2])) continue;
          var tipoF = normalizar_(regFolgas[k][1]);
          var qtdF = parseInt(regFolgas[k][3], 10);
          if (isNaN(qtdF)) qtdF = 0;
          if (tipoF.indexOf("RESET") !== -1) { resultado.resetUsado = 1; resultado.resetUsadoEm = iso_(parseData_(regFolgas[k][2])); }
          if (tipoF.indexOf("GIFT") !== -1) resultado.giftUsado += qtdF;
        }
      }
    }

    resultado.giftSaldo = Math.max(0, resultado.giftTotal - resultado.giftUsado);
    resultado.resetSaldo = Math.max(0, resultado.resetTotal - resultado.resetUsado);
    try {
      var bh = saldoBancoHorasDe_(nomePesquisa);
      resultado.bhHoras = bh.disponivel || 0; resultado.bhFmt = bh.disponivelFmt || "0h00"; resultado.bhDias = bh.dias || 0;
      resultado.bhJornada = jornadaHoras_();
    } catch (e) { resultado.bhHoras = 0; resultado.bhFmt = "0h00"; resultado.bhDias = 0; }
    return resultado;
  } catch (err) {
    throw new Error("Erro no cálculo analítico dos saldos: " + err.message);
  }
}

function atualizarAbaMVTComoUtilizado(nomeColaborador) {
  var ss = obterPlanilhaAtiva();
  var sheetMVT = ss.getSheetByName("MVT");
  if (!sheetMVT || sheetMVT.getLastRow() <= 1) return;
  
  var nomeProcura = nomeColaborador.toUpperCase().trim();
  var dados = sheetMVT.getRange(2, 1, sheetMVT.getLastRow() - 1, 5).getValues();
  
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][0] && dados[i][0].toString().toUpperCase().trim() === nomeProcura) {
      var statusUtilizou = dados[i][4] ? dados[i][4].toString().toUpperCase().trim() : "NÃO";
      if (statusUtilizou === "NÃO" || statusUtilizou === "") {
        var inlineReal = i + 2;
        sheetMVT.getRange(inlineReal, 5).setValue("Sim"); 
        break;
      }
    }
  }
}

// =========================================================================================
// INTERAÇÃO NATIVA DE GRAVAÇÃO COM O GOOGLE CALENDAR
// =========================================================================================
function injetarEventoNoCalendarioGoogle(tituloEvento, dataInicioStr, dataFimStr) {
  try {
    if (!idAgenda_()) { Logger.log("Agenda do time não configurada."); return false; }
    var agenda = CalendarApp.getCalendarById(idAgenda_());
    if (!agenda) {
      Logger.log("Não foi possível encontrar a agenda pelo ID fornecido.");
      return false;
    }
    
    var processarDataSemDistorcao = function(input) {
      if (!input) return new Date();
      var d = new Date(input);
      if (input.toString().includes("-") && !input.toString().includes("T")) {
        var partes = input.toString().split("-");
        return new Date(parseInt(partes[0],10), parseInt(partes[1],10) - 1, parseInt(partes[2],10), 0, 0, 0);
      }
      return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0);
    };
    
    var d1 = processarDataSemDistorcao(dataInicioStr);
    var d2 = processarDataSemDistorcao(dataFimStr);
    
    if (d1.getTime() > d2.getTime()) {
      var temp = d1;
      d1 = d2;
      d2 = temp;
    }
    
    d2.setDate(d2.getDate() + 1);
    agenda.createAllDayEvent(tituloEvento, d1, d2);
    return true;
    
  } catch (e) {
    Logger.log("Erro ao tentar gravar na agenda via método nativo: " + e.message);
    return false;
  }
}

// =========================================================================================
// FUNÇÕES DE DISPARO E SINCRO
// =========================================================================================
function testarSincronizacaoUnica() {
  var ui = SpreadsheetApp.getUi();
  var ss = obterPlanilhaAtiva();
  var LINHA_PARA_TESTE = 32; 
  
  var sheetFerias = ss.getSheetByName("Controle Consolidado - 2026");
  if (!sheetFerias) {
    ui.alert("Erro", "Aba 'Controle Consolidado - 2026' não encontrada.", ui.ButtonSet.OK);
    return;
  }
  
  var r = sheetFerias.getRange(LINHA_PARA_TESTE, 1, 1, 15).getValues()[0];
  
  var nome = r[1] ? r[1].toString().toUpperCase().trim() : "";
  var situacao = r[4] ? r[4].toString().trim() : "";
  var dIni = r[9];
  var dFim = r[10];
  var dRetornoAfast = r[13];
  
  if (nome === "" || !dIni || !dFim) {
    ui.alert("Aviso de Teste", "A linha " + LINHA_PARA_TESTE + " não possui datas válidas preenchidas para realizar o teste.", ui.ButtonSet.OK);
    return;
  }
  
  var ultimaData = dRetornoAfast ? dRetornoAfast : dFim;
  var titulo = "Ausência: " + nome + " [" + (situacao ? situacao : "Férias") + "]";
  
  var sucesso = injetarEventoNoCalendarioGoogle(titulo, dIni, ultimaData);
  
  if (sucesso) {
    ui.alert("Sucesso Absoluto!", "Conseguimos Bruce!\n\nAs datas foram normalizadas no fuso local e o evento '" + titulo + "' foi inserido com sucesso na sua agenda.", ui.ButtonSet.OK);
  } else {
    ui.alert("Erro de Gravação", "O Google recusou a gravação do evento.", ui.ButtonSet.OK);
  }
}

function publicarHistoricoAntigoNoCalendario() {
  var ui = SpreadsheetApp.getUi();
  var totalEnviados = publicarLoteCalendario_();
  ui.alert("Sincronização Concluída", "Processo em lote finalizado com sucesso!\nForam publicados " + totalEnviados + " novos eventos.", ui.ButtonSet.OK);
}

// Núcleo do envio em lote (sem janela): usado pelo menu, pelo dashboard e pela sincronização automática
function publicarLoteCalendario_() {
  var ss = obterPlanilhaAtiva();
  var totalEnviados = 0;
  
  var sheetFerias = ss.getSheetByName("Controle Consolidado - 2026");
  if (sheetFerias && sheetFerias.getLastRow() > 1) {
    var registrosFerias = sheetFerias.getRange(2, 1, sheetFerias.getLastRow() - 1, 15).getValues(); 
    for (var i = 0; i < registrosFerias.length; i++) {
      var nome = registrosFerias[i][1] ? registrosFerias[i][1].toString().toUpperCase().trim() : "";
      var situacao = registrosFerias[i][4] ? registrosFerias[i][4].toString().trim() : "";
      var statusLinha = registrosFerias[i][8] ? registrosFerias[i][8].toString().toUpperCase().trim() : "";
      var dIni = registrosFerias[i][9];
      var dFim = registrosFerias[i][10];
      var dRetornoAfast = registrosFerias[i][13];
      var statusSistema = registrosFerias[i][14] ? registrosFerias[i][14].toString() : ""; 
      
      // EXCLUSIVO PARA IMPORTADOS: Só envia o lote se o status for estritamente AGENDADOS
      if (nome !== "" && statusLinha === "AGENDADOS" && dIni && dFim && !statusSistema.includes("Espelhado no Calendar")) {
        var ultimaData = dRetornoAfast ? dRetornoAfast : dFim;
        var titulo = "Ausência: " + nome + " [" + (situacao ? situacao : "Férias") + "]";
        
        var sucesso = injetarEventoNoCalendarioGoogle(titulo, dIni, ultimaData);
        if (sucesso) {
          sheetFerias.getRange(i + 2, 15).setValue("Espelhado no Calendar"); 
          totalEnviados++;
        }
      }
    }
  }
  
  var sheetFolgas = ss.getSheetByName("Folgas e Emendas - 2026");
  if (sheetFolgas && sheetFolgas.getLastRow() > 1) {
    var registrosFolgas = sheetFolgas.getRange(2, 1, sheetFolgas.getLastRow() - 1, 7).getValues(); 
    for (var k = 0; k < registrosFolgas.length; k++) {
      var nomeF = registrosFolgas[k][0] ? registrosFolgas[k][0].toString().toUpperCase().trim() : "";
      var beneficio = registrosFolgas[k][1] ? registrosFolgas[k][1].toString().trim() : "";
      var dataFolga = registrosFolgas[k][2];
      var dataRetorno = registrosFolgas[k][4];
      var origem = registrosFolgas[k][6] ? registrosFolgas[k][6].toString() : ""; 
      
      if (nomeF !== "" && dataFolga && !origem.includes("Espelhado no Calendar")) {
        var dataFimFolga = dataRetorno ? dataRetorno : dataFolga;
        var tituloF = "Ausência: " + nomeF + " [" + beneficio + "]";
        
        var sucessoF = injetarEventoNoCalendarioGoogle(tituloF, dataFolga, dataFimFolga);
        if (sucessoF) {
          sheetFolgas.getRange(k + 2, 7).setValue("Painel Avulso Time | Espelhado no Calendar"); 
          totalEnviados++;
        }
      }
    }
  }
  
  return totalEnviados;
}

// =========================================================================================
// FUNÇÃO DE GRAVAÇÃO DO PAINEL (CONTROLE DE REGRAS DE STATUS)
// =========================================================================================
function salvarSugestaoDev(dados) {
  try {
    var saldos = consultarSaldosColaborador(dados.nome);
    
    // Normaliza a opção do benefício. Se for vazio, nulo ou o marcador padrão, vira string limpa ""
    var opcao = dados.opcaoAfastamento ? dados.opcaoAfastamento.toString().toUpperCase().trim() : "";
    if (opcao.includes("SELECIONE") || opcao === "-" || opcao === "") {
      opcao = "";
    }
    
    var dataAtual = new Date();
    
    if (opcao !== "") {
      if (opcao.includes("MVT") && saldos.mvtSaldo <= 0) {
        throw new Error("Bloqueio de Elegibilidade: Colaborador não possui saldo ou direito a MVT ativo.");
      }
      if (opcao.includes("GIFT") && saldos.giftSaldo <= 0) {
        throw new Error("Bloqueio de Elegibilidade: Saldo de GiftDay insuficiente.");
      }
      if (opcao.includes("BANCO")) {
        var diasBH = (parseInt(dados.quantidadeDias, 10) || 0) === 0
          ? Math.max(1, parseInt(dados.quantidadeDiasBeneficio, 10) || 1)
          : (parseInt(dados.diasEmendaBH, 10) || 0);
        if (diasBH > saldos.bhDias) {
          throw new Error("Bloqueio de Elegibilidade: banco de horas disponível " + (saldos.bhFmt || "0h00") + " (" + saldos.bhDias + " dia(s)); pedido de " + diasBH + " dia(s).");
        }
      }
      if (opcao.includes("RESET") && saldos.resetSaldo <= 0) {
        throw new Error("Bloqueio de Elegibilidade: o ResetDay deste ano já foi usado" + (saldos.resetUsadoEm ? " (" + Utilities.formatDate(parseData_(saldos.resetUsadoEm), Session.getScriptTimeZone(), "dd/MM/yyyy") + ")" : "") + ".");
      }
    }

    var ss = obterPlanilhaAtiva();
    var nomeExibicaoCalendario = dados.nome.toUpperCase().trim();
    
    // REGRA DE FOLGAS AVULSAS: Criação imediata e espelhamento direto (exige benefício selecionado)
    if (parseInt(dados.quantidadeDias, 10) === 0 || !dados.quantidadeDias) {
      if (opcao === "") {
        throw new Error("Para lançar uma folga de 0 dias, você precisa selecionar o benefício correspondente (Ex: GiftDay, ResetDay, etc).");
      }
      var sheetFolgas = ss.getSheetByName("Folgas e Emendas - 2026") || inicializarTodasAbas();
      var novaLinhaFolga = [
        dados.nome.toUpperCase().trim(),
        opcao,
        dados.dataInicio,
        parseInt(dados.quantidadeDiasBeneficio, 10),
        dados.dataRetornoAfastamento,
        dataAtual,
        "Painel Avulso Time | Espelhado no Calendar" 
      ];
      sheetFolgas.appendRow(novaLinhaFolga);
      sheetFolgas.getRange(sheetFolgas.getLastRow(), 1, 1, 7).setBackground("#e8f5e9");
      
      if (opcao.includes("MVT")) atualizarAbaMVTComoUtilizado(dados.nome);
      
      var tituloFolga = "Ausência: " + nomeExibicaoCalendario + " [" + opcao + "]";
      // Último dia de folga = início + quantidade - 1 (o dia de retorno não entra na agenda)
      var qtdFolga = Math.max(1, parseInt(dados.quantidadeDiasBeneficio, 10) || 1);
      var iniFolga = parseData_(dados.dataInicio);
      var dataFimAvulsa = iniFolga ? iso_(new Date(iniFolga.getFullYear(), iniFolga.getMonth(), iniFolga.getDate() + qtdFolga - 1)) : dados.dataInicio;
      injetarEventoNoCalendarioGoogle(tituloFolga, dados.dataInicio, dataFimAvulsa);
      
      return "Sucesso! Lançamento salvo e sincronizado na agenda.";
    }
    
    var sheetFerias = ss.getSheetByName("Controle Consolidado - 2026");
    
    // REGRA DE ACOPLAMENTO: Se estiver emendando em uma linha já agendada do arquivo importado
    if (dados.linhaVinculo && parseInt(dados.linhaVinculo, 10) > 0) {
      var inlineAlvo = parseInt(dados.linhaVinculo, 10);
      var valorAtualSituacao = sheetFerias.getRange(inlineAlvo, 5).getValue().toString().trim();
      
      var textoAcoplamento = opcao ? valorAtualSituacao + " + (" + opcao + " Solicitado)" : valorAtualSituacao;
      sheetFerias.getRange(inlineAlvo, 5).setValue(textoAcoplamento);
      
      if(dados.dataRetornoAfastamento) sheetFerias.getRange(inlineAlvo, 14).setValue(new Date(dados.dataRetornoAfastamento + "T00:00:00")); 
      sheetFerias.getRange(inlineAlvo, 15).setValue("Espelhado no Calendar"); 
      sheetFerias.getRange(inlineAlvo, 16).setValue(dataAtual);
      
      if (opcao !== "" && opcao.includes("MVT")) atualizarAbaMVTComoUtilizado(dados.nome);
      
      var tituloEmenda = "Ausência: " + nomeExibicaoCalendario + " [Férias" + (opcao ? " + " + opcao : "") + "]";
      injetarEventoNoCalendarioGoogle(tituloEmenda, dados.dataInicio, dados.dataRetornoAfastamento ? dados.dataRetornoAfastamento : dados.dataFim);
      
      return "Sucesso! Benefício acoplado e agenda updated.";
    } 
    
    // CENÁRIO B: NOVA INCLUSÃO MANUAL PELO FORMULÁRIO (Configura coluna como "Sugestão")
    var situacaoLabel = opcao ? "Sugestão do Dev (" + opcao + ")" : "Sugestão do Dev (Padrão)";
    var novaLinhaFerias = [
      "jun. de 2026", 
      dados.nome.toUpperCase().trim(), 
      "-", 
      "-", 
      situacaoLabel, 
      "", 
      "", 
      "", 
      "Sugestão", 
      new Date(dados.dataInicio + "T00:00:00"), 
      new Date(dados.dataFim + "T00:00:00"), 
      parseInt(dados.quantidadeDias, 10),
      dados.dataRetornoPadrao ? new Date(dados.dataRetornoPadrao + "T00:00:00") : "", 
      dados.dataRetornoAfastamento ? new Date(dados.dataRetornoAfastamento + "T00:00:00") : "",
      "Espelhado no Calendar", 
      dataAtual, 
      "Painel Formulário Dev" 
    ];
    sheetFerias.appendRow(novaLinhaFerias);
    var linhaNova = sheetFerias.getLastRow();
    sheetFerias.getRange(linhaNova, 1, 1, 17).setBackground("#fffde7");
    
    if (opcao !== "" && opcao.includes("MVT")) atualizarAbaMVTComoUtilizado(dados.nome);
    
    // CHECAGEM DE CONFLITO: se houver sobreposição, vai para aprovação do gestor
    // e só é espelhado no Calendar depois de aprovado
    var resultadoConflito = processarLancamentoFerias(linhaNova);
    var diasEmenda = opcao.indexOf("BANCO") !== -1 ? (parseInt(dados.diasEmendaBH, 10) || 0) : 0;
    if (resultadoConflito.status === "pendente") {
      return resultadoConflito.mensagem + (diasEmenda ? " A emenda de " + diasEmenda + " dia(s) com banco de horas não foi lançada: lance como folga avulsa depois da aprovação." : "");
    }
    if (diasEmenda > 0) registrarEmendaBancoHoras_(dados.nome, dados.dataFim, diasEmenda);
    
    var tituloFeriasNovas = "Ausência: " + nomeExibicaoCalendario + " [Férias" + (opcao ? " + " + opcao : "") + "]";
    var ultimaDataAusente = dados.dataRetornoAfastamento ? dados.dataRetornoAfastamento : dados.dataFim;
    injetarEventoNoCalendarioGoogle(tituloFeriasNovas, dados.dataInicio, ultimaDataAusente);
    
    return "Sucesso! Férias gravadas como Sugestão e espelhadas na agenda.";
    
  } catch (err) { throw new Error(err.message); }
}

// Emenda de férias com banco de horas: folga de N dias a partir do dia seguinte ao fim das férias
function registrarEmendaBancoHoras_(nome, fimFeriasIso, dias) {
  var fim = parseData_(fimFeriasIso);
  if (!fim || !dias) return;
  var ini = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate() + 1);
  var ult = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate() + dias - 1);
  var ret = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate() + dias);
  var aba = obterPlanilhaAtiva().getSheetByName("Folgas e Emendas - 2026") || inicializarTodasAbas();
  aba.appendRow([String(nome).toUpperCase().trim(), "BANCO DE HORAS (emenda)", iso_(ini), dias, iso_(ret), new Date(), "Emenda de férias | Espelhado no Calendar"]);
  injetarEventoNoCalendarioGoogle("Ausência: " + String(nome).toUpperCase().trim() + " [BANCO DE HORAS - emenda]", iso_(ini), iso_(ult));
}

function inicializarTodasAbas() {
  var ss = obterPlanilhaAtiva();
  var sheetFerias = ss.getSheetByName("Controle Consolidado - 2026") || ss.insertSheet("Controle Consolidado - 2026");
  var sheetFolgas = ss.getSheetByName("Folgas e Emendas - 2026") || ss.insertSheet("Folgas e Emendas - 2026");
  return sheetFolgas;
}

// =========================================================================================
// MOTOR DE IMPORTAÇÃO E CONCILIAÇÃO INTELIGENTE DE SUGESTÕES (VERSÃO CORRIGIDA)
// =========================================================================================
// Importa a tabela de férias do RH (exportação do Looker Studio "Férias_Férias_Tabela").
// Para cada férias AGENDADA no RH:
//  1. já existe no controle (mesmo nome e início): confirma como AGENDADOS e corrige o fim, se mudou;
//  2. existe uma Sugestão do dev do mesmo colaborador que se sobrepõe: troca pelas datas do RH
//     (e remove da agenda o evento com as datas antigas);
//  3. não existe: inclui uma linha nova "Importado do RH", que o envio em lote publica na agenda.
function importarDadosPorIdPlanilha(idOrigem) {
  try {
    var idPlanilhaRH = (idOrigem && String(idOrigem).trim() !== "") ? String(idOrigem).trim() : cfg_("ID_PLANILHA_RH");
    if (!idPlanilhaRH) throw new Error("Informe o ID da planilha do RH.");
    var ssOrigem = SpreadsheetApp.openById(idPlanilhaRH);
    var sheetRH = ssOrigem.getSheetByName("Sheet1") || ssOrigem.getSheets()[0];
    var sheetDestino = obterPlanilhaAtiva().getSheetByName("Controle Consolidado - 2026");
    if (!sheetRH || !sheetDestino) throw new Error("Não foi possível localizar as abas necessárias para a importação.");
    if (sheetRH.getLastRow() < 2) return "A planilha do RH está vazia.";

    // Localiza as colunas pelo cabeçalho (com posição padrão como reserva)
    var cab = sheetRH.getRange(1, 1, 1, sheetRH.getLastColumn()).getValues()[0].map(function(c) { return normalizar_(c); });
    function col(nome, padrao) { var i = cab.indexOf(normalizar_(nome)); return i >= 0 ? i : padrao; }
    var C = {
      periodo: col("Período Selecinado", 0), nome: col("Nome", 1), tempo: col("Tempo TOTVS", 3), funcao: col("Funcao", 4),
      situacao: col("Situação Férias", 13), iniAq: col("Início Período Aquisitivo", 14), fimAq: col("Fim Período Aquisitivo", 15),
      limite: col("Data Limite", 16), agendamento: col("Agendamento", 17), ini: col("Data Inicio Programação", 18),
      fim: col("Data Fim Programação", 19), saldo: col("Saldo Dias", 20)
    };

    var dadosRH = sheetRH.getRange(2, 1, sheetRH.getLastRow() - 1, sheetRH.getLastColumn()).getValues();
    var nomeOrigem = ssOrigem.getName();
    var agora = new Date();
    var cont = { confirmadas: 0, ajustadas: 0, trocadas: 0, novas: 0, iguais: 0, historico: 0 };
    var alterados = []; // férias que entraram ou mudaram nesta importação (para o alerta de conflito)
    var novasLinhas = []; // gravadas de uma vez no final (mais confiável que appendRow linha a linha)
    var vistosRH = {};
    var doTime = nomesDoTime_();
    var foraDoTime = {};

    for (var i = 0; i < dadosRH.length; i++) {
      var r = dadosRH[i];
      var nome = String(r[C.nome] || "").trim().toUpperCase();
      var agend = normalizar_(r[C.agendamento]);
      var dIni = parseData_(r[C.ini]), dFim = parseData_(r[C.fim]);
      if (!nome || !dIni || !dFim) continue;
      // Só pessoas cadastradas no time (aba TIME). Protege contra exportação de outro time.
      if (doTime && !doTime[normalizar_(nome)]) { foraDoTime[nome] = true; continue; }

      // Férias já gozadas (Finalizadas): entram só como histórico, se ainda não existirem.
      // Não contam como conflito nem vão para a agenda.
      if (agend !== "AGENDADOS") {
        if (normalizar_(r[C.situacao]).indexOf("FINALIZAD") !== 0) continue;
        var chaveH = nome + "|" + dIni.getTime();
        if (vistosRH[chaveH]) continue;
        vistosRH[chaveH] = true;
        var ultH = sheetDestino.getLastRow();
        var existe = ultH > 1 && sheetDestino.getRange(2, 1, ultH - 1, 10).getValues().some(function(x) {
          var di = parseData_(x[9]);
          return String(x[1]).trim().toUpperCase() === nome && di && di.getTime() === dIni.getTime();
        });
        if (!existe) {
          novasLinhas.push([
            r[C.periodo] || "", nome, r[C.tempo] || "-", r[C.funcao] || "-", "Finalizadas",
            parseData_(r[C.iniAq]) || "", parseData_(r[C.fimAq]) || "", parseData_(r[C.limite]) || "",
            r[C.agendamento] || "Não Agendados", dIni, dFim, diasEntre_(dIni, dFim),
            new Date(dFim.getFullYear(), dFim.getMonth(), dFim.getDate() + 1), "",
            "Histórico do RH", agora, nomeOrigem
          ]);
          cont.historico++;
        }
        continue;
      }
      var chaveRH = nome + "|" + dIni.getTime();
      if (vistosRH[chaveRH]) continue;
      vistosRH[chaveRH] = true;

      // Relê o destino a cada item (linhas podem ter sido incluídas/alteradas)
      var ult = sheetDestino.getLastRow();
      var dest = ult > 1 ? sheetDestino.getRange(2, 1, ult - 1, 17).getValues() : [];

      // 1. Mesmo nome e mesmo início
      var achou = -1;
      for (var j = 0; j < dest.length; j++) {
        var dj = parseData_(dest[j][9]);
        if (String(dest[j][1]).trim().toUpperCase() === nome && dj && dj.getTime() === dIni.getTime()) { achou = j; break; }
      }
      if (achou >= 0) {
        var linha = achou + 2, rowD = dest[achou];
        var fimAtual = parseData_(rowD[10]);
        var mudou = false;
        if (normalizar_(rowD[8]) !== "AGENDADOS") {
          sheetDestino.getRange(linha, 9).setValue("AGENDADOS");
          if (/SUGEST/i.test(String(rowD[4]))) sheetDestino.getRange(linha, 5).setValue("Marcadas");
          sheetDestino.getRange(linha, 1, 1, 17).setBackground(null);
          cont.confirmadas++; mudou = true;
        }
        if (!fimAtual || fimAtual.getTime() !== dFim.getTime()) {
          if (String(rowD[14]).indexOf("Espelhado") !== -1) removerEventoAgenda_(nome, dIni, fimAtual || dFim);
          sheetDestino.getRange(linha, 11).setValue(dFim);
          sheetDestino.getRange(linha, 12).setValue(diasEntre_(dIni, dFim));
          sheetDestino.getRange(linha, 13).setValue(new Date(dFim.getFullYear(), dFim.getMonth(), dFim.getDate() + 1));
          sheetDestino.getRange(linha, 14).setValue("");
          sheetDestino.getRange(linha, 15).setValue("Datas atualizadas pelo RH");
          cont.ajustadas++; mudou = true;
          alterados.push({ nome: nome, inicio: dIni, fim: dFim, tipo: "datas corrigidas pelo RH" });
        }
        if (!mudou) cont.iguais++;
        continue;
      }

      // 2. Sugestão do mesmo colaborador que se sobrepõe ao período do RH
      var sug = -1;
      for (var k = 0; k < dest.length; k++) {
        var rk = dest[k];
        if (String(rk[1]).trim().toUpperCase() !== nome) continue;
        if (normalizar_(rk[8]).indexOf("SUGEST") !== 0 && normalizar_(rk[4]).indexOf("SUGEST") === -1) continue;
        var ki = parseData_(rk[9]), kf = parseData_(rk[10]);
        if (ki && kf && sobreposicao_(ki, kf, dIni, dFim)) { sug = k; break; }
      }
      if (sug >= 0) {
        var ls = sug + 2, rs = dest[sug];
        if (String(rs[14]).indexOf("Espelhado") !== -1) removerEventoAgenda_(nome, parseData_(rs[9]), parseData_(rs[10]));
        sheetDestino.getRange(ls, 5).setValue("Marcadas");
        sheetDestino.getRange(ls, 9).setValue("AGENDADOS");
        sheetDestino.getRange(ls, 10, 1, 5).setValues([[dIni, dFim, diasEntre_(dIni, dFim),
          new Date(dFim.getFullYear(), dFim.getMonth(), dFim.getDate() + 1), ""]]);
        sheetDestino.getRange(ls, 15).setValue("Datas atualizadas pelo RH");
        sheetDestino.getRange(ls, 16).setValue(agora);
        sheetDestino.getRange(ls, 1, 1, 17).setBackground(null);
        cont.trocadas++;
        alterados.push({ nome: nome, inicio: dIni, fim: dFim, tipo: "sugestão trocada pelas datas do RH" });
        continue;
      }

      // 3. Férias nova vinda do RH
      novasLinhas.push([
        r[C.periodo] || "", nome, r[C.tempo] || "-", r[C.funcao] || "-", r[C.situacao] || "Marcadas",
        parseData_(r[C.iniAq]) || "", parseData_(r[C.fimAq]) || "", parseData_(r[C.limite]) || "",
        "AGENDADOS", dIni, dFim, diasEntre_(dIni, dFim),
        new Date(dFim.getFullYear(), dFim.getMonth(), dFim.getDate() + 1), "",
        "Importado do RH", agora, nomeOrigem
      ]);
      cont.novas++;
      alterados.push({ nome: nome, inicio: dIni, fim: dFim, tipo: "nova, vinda do RH" });
    }

    // Grava todas as linhas novas de uma vez e confere se ficaram na planilha
    var faixa = "";
    if (novasLinhas.length) {
      var primeira = Math.max(sheetDestino.getLastRow(), 1) + 1;
      var falta = primeira + novasLinhas.length - 1 - sheetDestino.getMaxRows();
      if (falta > 0) sheetDestino.insertRowsAfter(sheetDestino.getMaxRows(), falta);
      sheetDestino.getRange(primeira, 1, novasLinhas.length, 17).setValues(novasLinhas);
      novasLinhas.forEach(function(l, k) {
        if (l[14] === "Importado do RH") sheetDestino.getRange(primeira + k, 1, 1, 17).setBackground("#e8f0fe");
      });
      SpreadsheetApp.flush();
      var conferidas = sheetDestino.getRange(primeira, 2, novasLinhas.length, 1).getValues()
        .filter(function(x) { return String(x[0]).trim() !== ""; }).length;
      if (conferidas !== novasLinhas.length) {
        throw new Error("Tentei gravar " + novasLinhas.length + " linhas a partir da linha " + primeira +
          " da aba Controle Consolidado, mas só " + conferidas + " ficaram gravadas.");
      }
      faixa = " Gravadas nas linhas " + primeira + " a " + (primeira + novasLinhas.length - 1) + " da aba Controle Consolidado.";
    }

    var avisoFiltro = "";
    try {
      var filtro = sheetDestino.getFilter();
      if (filtro) {
        var cols = filtro.getRange().getNumColumns(), ini = filtro.getRange().getColumn();
        for (var cf = ini; cf < ini + cols; cf++) {
          if (filtro.getColumnFilterCriteria(cf)) { avisoFiltro = " 🔎 Atenção: a aba Controle Consolidado tem filtro ativo e algumas linhas podem estar ocultas."; break; }
        }
      }
    } catch (e) {}

    var qtdAlertas = 0;
    try { qtdAlertas = alertarConflitosImportacao_(sheetDestino, alterados); } catch (e) { Logger.log("Alerta de conflitos da importação: " + e.message); }

    return (qtdAlertas ? "⚠️ " + qtdAlertas + " férias do RH com conflito (gestor avisado por e-mail). " : "") +
      "Importação concluída: " + cont.novas + " férias novas incluídas, " +
      cont.trocadas + " sugestões substituídas pelas datas do RH, " +
      cont.confirmadas + " confirmadas como AGENDADOS, " +
      cont.ajustadas + " com datas corrigidas, " + cont.iguais + " já estavam em dia" +
      (cont.historico ? ", " + cont.historico + " férias finalizadas guardadas como histórico" : "") + "." + faixa + avisoFiltro +
      (Object.keys(foraDoTime).length ? " 🚫 Ignoradas " + Object.keys(foraDoTime).length + " pessoa(s) que não estão no time (aba TIME): " +
        Object.keys(foraDoTime).slice(0, 5).map(function(n) { return primeiroNome_(n); }).join(", ") + (Object.keys(foraDoTime).length > 5 ? "…" : "") + "." : "");
  } catch (e) {
    throw new Error("Erro durante a importação/atualização: " + e.message);
  }
}

// Férias que vieram do RH já estão agendadas oficialmente, então não passam por aprovação.
// Mas, se conflitarem com outras, o gestor recebe um e-mail com a lista.
function alertarConflitosImportacao_(sheet, alterados) {
  if (!alterados.length || !emailLider_()) return 0;
  var periodos = listarPeriodosAtivos_(sheet, -1);
  var hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  var itens = [];
  alterados.forEach(function(a) {
    if (a.fim < hoje) return;
    var conf = encontrarConflitos_(a.nome, a.inicio, a.fim, periodos);
    if (conf.length) itens.push({ a: a, conf: conf });
  });
  if (!itens.length) return 0;

  var html = "<p>A importação do RH trouxe férias <b>já agendadas no RH</b> que coincidem com férias de outras pessoas do time. " +
    "Como já estão agendadas oficialmente, elas não passam pela aprovação do painel; este e-mail é para você avaliar se precisa alinhar a cobertura.</p>";
  itens.forEach(function(it) {
    html += "<p style='margin:16px 0 4px'><b>" + esc_(it.a.nome) + "</b>: " + fmt_(it.a.inicio) + " a " + fmt_(it.a.fim) +
      " <span style='color:#5f6368'>(" + esc_(it.a.tipo) + ")</span></p>" + tabelaConflitosHtml_(it.conf);
  });
  enviarEmail_(emailLider_(), "⚠️ Férias importadas do RH com conflito (" + itens.length + ")", "Conflitos na importação do RH", html);
  return itens.length;
}

// Remove da agenda o evento "Ausência: NOME [Férias...]" que cobre o período antigo
// Nomes do time (aba TIME), normalizados. null quando a aba está vazia (aceita todos).
function nomesDoTime_() {
  var nomes = [];
  try { nomes = buscarNomesTime(); } catch (e) {}
  if (!nomes.length) return null;
  var m = {};
  nomes.forEach(function(n) { m[normalizar_(n)] = true; });
  return m;
}

function removerEventoAgenda_(nome, ini, fim, qualquerTitulo) {
  try {
    if (!ini) return;
    if (!idAgenda_()) return;
    var agenda = CalendarApp.getCalendarById(idAgenda_());
    if (!agenda) return;
    var ate = new Date((fim || ini).getFullYear(), (fim || ini).getMonth(), (fim || ini).getDate() + 2);
    agenda.getEvents(ini, ate).forEach(function(ev) {
      var t = ev.getTitle().toUpperCase();
      if (t.indexOf("AUSÊNCIA: " + nome) === 0 && (qualquerTitulo || t.indexOf("FÉRIAS") !== -1)) ev.deleteEvent();
    });
  } catch (e) {
    Logger.log("Não foi possível remover o evento antigo de " + nome + ": " + e.message);
  }
}

// =========================================================================================
// ATUALIZAÇÃO DO CONTROLE DE CONFLITOS COM DISPARO AUTOMÁTICO DE E-MAIL
// =========================================================================================
function atualizarAbaControleConflitosSemEmail_() {
  atualizarAbaControleConflitos(false);
}

function atualizarAbaControleConflitos(enviarEmail) {
  var ss = obterPlanilhaAtiva();
  var sheetFerias = ss.getSheetByName("Controle Consolidado - 2026");
  var sheetConflito = ss.getSheetByName("Controle de Conflitos") || ss.insertSheet("Controle de Conflitos");
  
  if (!sheetFerias || sheetFerias.getLastRow() <= 1) return;
  
  var registros = sheetFerias.getRange(2, 1, sheetFerias.getLastRow() - 1, 11).getValues();
  var listaPeriodos = [];
  
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var hojeMs = hoje.getTime();
  
  // 1. Extrai apenas ausências válidas e futuras/vigentes
  for (var i = 0; i < registros.length; i++) {
    var nome = registros[i][1] ? registros[i][1].toString().trim() : "";
    var situacao = registros[i][4] ? registros[i][4].toString().trim() : "";
    var dtIni = registros[i][9];
    var dtFim = registros[i][10];
    
    if (nome && dtIni && dtFim && !isNaN(new Date(dtIni).getTime()) && !isNaN(new Date(dtFim).getTime())) {
      var dFimObj = new Date(dtFim);
      dFimObj.setHours(23, 59, 59, 999);
      var dFimMs = dFimObj.getTime();
      
      var situacaoUpper = situacao.toUpperCase();
      
      // Ignora datas passadas ou ausências finalizadas/pagas
      if (dFimMs < hojeMs || situacaoUpper.includes("FINALIZADA") || situacaoUpper.includes("PAGA") ||
          situacaoUpper.indexOf("REPROVADO") === 0 || situacaoUpper.indexOf("AJUSTE SOLICITADO") === 0 ||
          situacaoUpper.indexOf("INCENTIVADAS") !== -1) {
        continue;
      }
      
      listaPeriodos.push({
        linhaOriginal: i + 2,
        nome: nome,
        situacao: situacao,
        dtIniMs: new Date(dtIni).getTime(),
        dtFimMs: new Date(dtFim).getTime(),
        dtIniStr: Utilities.formatDate(new Date(dtIni), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        dtFimStr: Utilities.formatDate(new Date(dtFim), Session.getScriptTimeZone(), "dd/MM/yyyy")
      });
    }
  }
  
  // Helper para calcular a quantidade de dias sobrepostos
  var calcularDiasSobrepostos = function(start1, end1, start2, end2) {
    var maxStart = start1 > start2 ? start1 : start2;
    var minEnd = end1 < end2 ? end1 : end2;
    
    if (maxStart <= minEnd) {
      var diffTime = Math.abs(minEnd - maxStart);
      return Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    }
    return 0;
  };

  // 2. Compara cada registro contra os demais
  var resultadoMatriz = [];
  
  for (var a = 0; a < listaPeriodos.length; a++) {
    var p1 = listaPeriodos[a];
    var conflitouCom = [];
    var totalDiasConflito = 0;
    
    for (var b = 0; b < listaPeriodos.length; b++) {
      if (a === b) continue;
      var p2 = listaPeriodos[b];
      
      if (p1.nome.toUpperCase() !== p2.nome.toUpperCase()) {
        var diasSobreposicao = calcularDiasSobrepostos(p1.dtIniMs, p1.dtFimMs, p2.dtIniMs, p2.dtFimMs);
        
        if (diasSobreposicao > 0) {
          conflitouCom.push(p2.nome + " (" + diasSobreposicao + "d)");
          if (diasSobreposicao > totalDiasConflito) {
            totalDiasConflito = diasSobreposicao;
          }
        }
      }
    }
    
    var statusConflito = conflitouCom.length > 0 ? "Conflito" : "Sem Conflito";
    var detalheConflito = conflitouCom.length > 0 ? conflitouCom.join(", ") : "-";
    var exibicaoDias = totalDiasConflito > 0 ? totalDiasConflito : "-";
    
    resultadoMatriz.push([
      p1.nome,
      p1.dtIniStr,
      p1.dtFimStr,
      p1.situacao ? p1.situacao : "-",
      statusConflito,
      detalheConflito,
      exibicaoDias
    ]);
  }
  
  // 3. Atualiza a aba Controle de Conflitos
  sheetConflito.clear();
  
  var cabecalho = [["Colaborador", "Data Início", "Data Fim", "Situação", "Status de Conflito", "Conflito Com", "Dias em Conflito"]];
  sheetConflito.getRange(5, 1, 1, 7).setValues(cabecalho)
               .setFontWeight("bold")
               .setBackground("#1f4e78")
               .setFontColor("#ffffff");
               
  if (resultadoMatriz.length > 0) {
    var rangeDados = sheetConflito.getRange(6, 1, resultadoMatriz.length, 7);
    rangeDados.setValues(resultadoMatriz);
    
    for (var r = 0; r < resultadoMatriz.length; r++) {
      var linhaTarget = r + 6;
      if (resultadoMatriz[r][4] === "Conflito") {
        sheetConflito.getRange(linhaTarget, 5).setBackground("#f8cecc").setFontColor("#b85450");
      } else {
        sheetConflito.getRange(linhaTarget, 5).setBackground("#d5e8d4").setFontColor("#274e13");
      }
    }
  }

  // 4. ENVIO AUTOMÁTICO DE E-MAIL
  // Se houver gestor configurado, dispara automaticamente
  if (enviarEmail !== false && emailLider_()) {
    enviarEmailConflitosLider(emailLider_());
  }
}