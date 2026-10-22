// =========================================================================================
// CONFIGURAÇÕES GLOBAIS
// =========================================================================================
// ID fixo da planilha para estabilidade absoluta nos gatilhos e operações
const PLANILHA_ID = "1ExPIKq8Rqqw8wU9dyae8KaHCIHmWKUxgyy-1XeoA9Zg";

// ID exato e verificado da agenda "Ausências Novas Apurações"
const ID_AGENDA_ALVO = "c_309f2f359160bb1d732268dbb2e33b80dda091c626cbb2f997c8066a87c004af@group.calendar.google.com";

// E-mail fixo do Líder/Gestor que receberá os relatórios de conflitos
const EMAIL_LIDER = "bruce.mello@totvs.com.br"; // <-- INSIRA O E-MAIL DO LÍDER AQUI

// =========================================================================================
// FUNÇÃO DO MENU (SEM NECESSIDADE DE DIGITAR O E-MAIL)
// =========================================================================================
function enviarEmailConflitosLiderDireto() {
  var ui = SpreadsheetApp.getUi();
  
  if (!EMAIL_LIDER || EMAIL_LIDER.includes("seu_email_aqui")) {
    ui.alert("Configuração Pendente", "Insira o e-mail correto do líder na constante EMAIL_LIDER no topo do código do Apps Script.", ui.ButtonSet.OK);
    return;
  }
  
  // Atualiza os conflitos antes de disparar o e-mail
  atualizarAbaControleConflitos();
  
  var enviado = enviarEmailConflitosLider(EMAIL_LIDER);
  
  if (enviado) {
    ui.alert("E-mail Enviado!", "O relatório de conflitos foi enviado com sucesso para:\n" + EMAIL_LIDER, ui.ButtonSet.OK);
  } else {
    ui.alert("Sem Conflitos", "Não foram encontrados conflitos pendentes para envio no momento.", ui.ButtonSet.OK);
  }
}

// =========================================================================================
// ATUALIZAÇÃO DO MENU
// =========================================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Apoio Gestão')
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

function gatilhoAberturaAutomatica() {
  try { abrirPainel(); } catch (e) {}
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
function obterPlanilhaAtiva() {
  return SpreadsheetApp.getActiveSpreadsheet();
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

function consultarSaldosColaborador(nomePesquisa) {
  try {
    var resultado = { giftTotal: 0, giftUsado: 0, giftSaldo: 0, resetTotal: 1, resetUsado: 0, resetSaldo: 1, mvtTotal: 0, mvtUsado: 0, mvtSaldo: 0, anosEmpresa: 0 };
    if (!nomePesquisa) return resultado;
    var nomeColab = nomePesquisa.toUpperCase().trim();
    var ss = obterPlanilhaAtiva();
    
    var sheetAdmissao = ss.getSheetByName("Admisssão");
    if (sheetAdmissao && sheetAdmissao.getLastRow() > 1) {
      var dadosAdmissao = sheetAdmissao.getRange(2, 1, sheetAdmissao.getLastRow() - 1, 2).getValues();
      for (var j = 0; j < dadosAdmissao.length; j++) {
        if (dadosAdmissao[j][0] && dadosAdmissao[j][0].toString().toUpperCase().trim() === nomeColab) {
          if (dadosAdmissao[j][1]) {
            var dataAdmissao = new Date(dadosAdmissao[j][1]);
            var anosCompletosNoAno = 2026 - dataAdmissao.getFullYear();
            resultado.anosEmpresa = anosCompletosNoAno;
            if (anosCompletosNoAno === 5) resultado.giftTotal = 2;
            else if (anosCompletosNoAno === 10) resultado.giftTotal = 3;
            else if (anosCompletosNoAno === 15) resultado.giftTotal = 4;
            else if (anosCompletosNoAno === 20) resultado.giftTotal = 5;
          }
          break;
        }
      }
    }
    
    var sheetMVT = ss.getSheetByName("MVT");
    if (sheetMVT && sheetMVT.getLastRow() > 1) {
      var dadosMVT = sheetMVT.getRange(2, 1, sheetMVT.getLastRow() - 1, 5).getValues();
      for (var m = 0; m < dadosMVT.length; m++) {
        if (dadosMVT[m][0] && dadosMVT[m][0].toString().toUpperCase().trim() === nomeColab) {
          var qtdDireito = parseInt(dadosMVT[m][3], 10); 
          var utilizouStr = dadosMVT[m][4] ? dadosMVT[m][4].toString().toUpperCase().trim() : "NÃO"; 
          
          if (!isNaN(qtdDireito) && qtdDireito === 1) {
            resultado.mvtTotal += 1; 
            if (utilizouStr === "SIM") {
              resultado.mvtUsado += 1; 
            }
          }
        }
      }
    }
    resultado.mvtSaldo = Math.max(0, resultado.mvtTotal - resultado.mvtUsado);

    var sheetFerias = ss.getSheetByName("Controle Consolidado - 2026");
    if (sheetFerias && sheetFerias.getLastRow() > 1) {
      var regFerias = sheetFerias.getRange(2, 1, sheetFerias.getLastRow() - 1, 12).getValues();
      for (var i = 0; i < regFerias.length; i++) {
        if (regFerias[i][1] && regFerias[i][1].toString().toUpperCase().trim() === nomeColab) {
          var sit = regFerias[i][4] ? regFerias[i][4].toString().toUpperCase() : "";
          if (sit.includes("RESET")) resultado.resetUsado = 1;
          if (sit.includes("GIFT")) resultado.giftUsado += 1; 
        }
      }
    }
    
    var sheetFolgas = ss.getSheetByName("Folgas e Emendas - 2026");
    if (sheetFolgas && sheetFolgas.getLastRow() > 1) {
      var regFolgas = sheetFolgas.getRange(2, 1, sheetFolgas.getLastRow() - 1, 5).getValues();
      for (var k = 0; k < regFolgas.length; k++) {
        if (regFolgas[k][0] && regFolgas[k][0].toString().toUpperCase().trim() === nomeColab) {
          var tipoF = regFolgas[k][1] ? regFolgas[k][1].toString().toUpperCase() : "";
          var qtdF = parseInt(regFolgas[k][3], 10);
          if (isNaN(qtdF)) qtdF = 0;
          
          if (tipoF.includes("RESET")) resultado.resetUsado = 1;
          if (tipoF.includes("GIFT")) resultado.giftUsado += qtdF;
        }
      }
    }
    
    resultado.giftSaldo = Math.max(0, resultado.giftTotal - resultado.giftUsado);
    resultado.resetSaldo = Math.max(0, resultado.resetTotal - resultado.resetUsado);
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
    var agenda = CalendarApp.getCalendarById(ID_AGENDA_ALVO.trim());
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
  
  ui.alert("Sincronização Concluída", "Processo em lote finalizado com sucesso!\nForam publicados " + totalEnviados + " novos eventos.", ui.ButtonSet.OK);
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
      var dataFimAvulsa = dados.dataRetornoAfastamento ? dados.dataRetornoAfastamento : dados.dataInicio;
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
    sheetFerias.getRange(sheetFerias.getLastRow(), 1, 1, 17).setBackground("#fffde7");
    
    if (opcao !== "" && opcao.includes("MVT")) atualizarAbaMVTComoUtilizado(dados.nome);
    
    var tituloFeriasNovas = "Ausência: " + nomeExibicaoCalendario + " [Férias" + (opcao ? " + " + opcao : "") + "]";
    var ultimaDataAusente = dados.dataRetornoAfastamento ? dados.dataRetornoAfastamento : dados.dataFim;
    injetarEventoNoCalendarioGoogle(tituloFeriasNovas, dados.dataInicio, ultimaDataAusente);
    
    return "Sucesso! Férias gravadas como Sugestão e espelhadas na agenda.";
    
  } catch (err) { throw new Error(err.message); }
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
function importarDadosPorIdPlanilha(idOrigem) {
  try {
    var idPlanilhaRH = (idOrigem && idOrigem.trim() !== "") ? idOrigem.trim() : "1xhmo4t63liWnr-GDQyNPbEIOHjfF-va_GxQ2uy4e-aE";
    
    var ssOrigem = SpreadsheetApp.openById(idPlanilhaRH);
    var ssDestino = obterPlanilhaAtiva();
    
    var sheetRH = ssOrigem.getSheetByName("Sheet1") || ssOrigem.getSheets()[0];
    var sheetDestino = ssDestino.getSheetByName("Controle Consolidado - 2026");
    
    if (!sheetRH || !sheetDestino) {
      throw new Error("Não foi possível localizar as abas necessárias para a importação.");
    }
    
    var dadosRH = sheetRH.getRange(2, 1, sheetRH.getLastRow() - 1, sheetRH.getLastColumn()).getValues();
    var ultimaLinhaDestino = sheetDestino.getLastRow();
    
    if (ultimaLinhaDestino <= 1) {
      return "Planilha de destino vazia. Preencha os cabeçalhos antes de importar.";
    }
    
    var rangeDestino = sheetDestino.getRange(2, 1, ultimaLinhaDestino - 1, 17);
    var dadosDestino = rangeDestino.getValues();
    
    var atualizadosCount = 0;
    
    // Função para normalizar datas no fuso horário do Script
    var formatarDataPadrao = function(valor) {
      if (!valor) return "";
      if (valor instanceof Date) {
        return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      var str = valor.toString().trim();
      if (str.includes("/")) return str;
      var d = new Date(str);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return str;
    };

    // 1. Mapeia os agendamentos confirmados vindos do RH (Planilha Origem)
    var agendadosRH = [];
    for (var i = 0; i < dadosRH.length; i++) {
      var row = dadosRH[i];
      var nomeRH = row[1] ? row[1].toString().toUpperCase().trim() : "";
      var statusRH = row[17] ? row[17].toString().trim() : ""; // Coluna R (Agendamento)
      var dtIniRH = formatarDataPadrao(row[18]);               // Coluna S (Data Início Programação)
      var dtFimRH = formatarDataPadrao(row[19]);               // Coluna T (Data Fim Programação)
      
      if (nomeRH && statusRH.toUpperCase() === "AGENDADOS" && dtIniRH) {
        agendadosRH.push({
          nome: nomeRH,
          dataInicio: dtIniRH,
          dataFim: dtFimRH
        });
      }
    }
    
    // 2. Compara com a Planilha de Destino
    for (var j = 0; j < dadosDestino.length; j++) {
      var nomeDest = dadosDestino[j][1] ? dadosDestino[j][1].toString().toUpperCase().trim() : "";
      var statusDest = dadosDestino[j][8] ? dadosDestino[j][8].toString().trim() : ""; // Coluna I (Agendamento)
      var dtIniDest = formatarDataPadrao(dadosDestino[j][9]);                          // Coluna J (Data Início)
      var dtFimDest = formatarDataPadrao(dadosDestino[j][10]);                         // Coluna K (Data Fim)
      
      // Se a linha for uma Sugestão
      if (statusDest.toUpperCase().includes("SUGESTÃO")) {
        for (var k = 0; k < agendadosRH.length; k++) {
          var itemRH = agendadosRH[k];
          
          // Regra de Match: Se o nome for igual E (as datas baterem OU se a data de início da sugestão estiver vazia na planilha de destino)
          var matchNome = (nomeDest === itemRH.nome);
          var matchData = (dtIniDest === itemRH.dataInicio) || (dtIniDest === "");
          
          if (matchNome && matchData) {
            var linhaRealSheet = j + 2;
            
            // Atualiza o Status para AGENDADOS
            sheetDestino.getRange(linhaRealSheet, 9).setValue("AGENDADOS");
            
            // Preenche as datas se estiverem ausentes na planilha de destino
            if (dtIniDest === "") {
              sheetDestino.getRange(linhaRealSheet, 10).setValue(itemRH.dataInicio);
              sheetDestino.getRange(linhaRealSheet, 11).setValue(itemRH.dataFim);
            }
            
            // Ajusta rótulo de situação
            var situacaoAtual = sheetDestino.getRange(linhaRealSheet, 5).getValue().toString();
            if (situacaoAtual.includes("Sugestão do Dev")) {
              sheetDestino.getRange(linhaRealSheet, 5).setValue("-");
            }
            
            // Limpa o fundo amarelo/destaque
            sheetDestino.getRange(linhaRealSheet, 1, 1, 17).setBackground(null);
            
            // Remove do array para não reusar o mesmo item do RH em outra linha
            agendadosRH.splice(k, 1);
            atualizadosCount++;
            break;
          }
        }
      }
    }
    
    return "Processo Concluído!\n- Sugestões atualizadas para AGENDADOS: " + atualizadosCount;
    
  } catch (e) {
    throw new Error("Erro durante a importação/atualização: " + e.message);
  }
}

// =========================================================================================
// ATUALIZAÇÃO DO CONTROLE DE CONFLITOS COM DISPARO AUTOMÁTICO DE E-MAIL
// =========================================================================================
function atualizarAbaControleConflitos() {
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
      if (dFimMs < hojeMs || situacaoUpper.includes("FINALIZADA") || situacaoUpper.includes("PAGA")) {
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
  // Se houver e-mail válido configurado na constante EMAIL_LIDER, dispara automaticamente
  if (typeof EMAIL_LIDER !== "undefined" && EMAIL_LIDER && !EMAIL_LIDER.includes("seu_email_aqui")) {
    enviarEmailConflitosLider(EMAIL_LIDER);
  }
}