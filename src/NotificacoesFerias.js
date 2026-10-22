// =========================================================================
// CONFIGURAÇÕES GERAIS
// =========================================================================
const CONFIG_NOTIFICAO = {
  emailDestino: "bruce.mello@totvs.com.br,squad.tributos@totvs.com.br",
  // Opcional: Cole a URL do Webhook do Google Chat caso queira receber lá também
  webhookGoogleChat: "https://chat.googleapis.com/v1/spaces/AAQASQqZuUU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=m7pf4GBuC-Iz0bW_rFXt8m271klkAXq4WHlfILZrN2c" 
  //webhookGoogleChat: "https://chat.googleapis.com/v1/spaces/AAQAxo7NLbM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Wrx_kPU0xvSmNeMPaV22egxXocbZJ9eF2-tmplhzS7g" 
};
// URL de exibição direta da imagem do seu Google Drive
const IMAGE_URL_FERIAS = "https://lh3.googleusercontent.com/d/1SLRN8cud5B3WcFlBktv_xMKi1CUymEcJ";

// =========================================================================
// FUNÇÃO PRINCIPAL DE NOTIFICAÇÃO DE FÉRIAS
// =========================================================================
function notificarProximasFerias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Controle Consolidado - 2026");
  
  if (!sheet) {
    Logger.log("Aba 'Controle Consolidado - 2026' não foi encontrada.");
    return;
  }

  const dados = sheet.getDataRange().getValues();
  const hoje = new Date();
  let feriasEncontradas = [];

  for (let i = 1; i < dados.length; i++) {
    const colaborador = dados[i][1];          // Coluna B: Nome / Colaborador
    const situacao = dados[i][4];             // Coluna E: Situação Férias
    const dataInicio = new Date(dados[i][9]);  // Coluna J: Data Início Programação
    const dataFim = new Date(dados[i][10]);    // Coluna K: Data Fim Programação
    const saldo = dados[i][11];               // Coluna L: Saldo

    if (colaborador && !isNaN(dataInicio.getTime()) && dataInicio >= hoje) {
      const diferencaDias = Math.ceil((dataInicio - hoje) / (1000 * 60 * 60 * 24));

      // Busca férias que iniciam nos próximos 30 dias
      if (diferencaDias <= 30 && diferencaDias >= 0) {
        feriasEncontradas.push({
          nome: colaborador,
          inicio: dataInicio.toLocaleDateString("pt-BR"),
          fim: dataFim.toLocaleDateString("pt-BR"),
          diasFaltantes: diferencaDias,
          situacao: situacao || "Marcadas",
          saldo: saldo
        });
      }
    }
  }

  if (feriasEncontradas.length === 0) {
    Logger.log("Nenhuma férias próxima encontrada nos próximos 30 dias.");
    return;
  }

  // Envia por E-mail
  enviarEmailFerias(feriasEncontradas);

  // Envia para o Google Chat (se a URL do webhook estiver preenchida)
  if (CONFIG_NOTIFICAO.webhookGoogleChat) {
    enviarGoogleChatFerias(feriasEncontradas);
  }
}

// ---------------------------------------------------------------------------
// 1. Envio de E-mail Formatado (HTML)
// ---------------------------------------------------------------------------
function enviarEmailFerias(listaFerias) {
  const assunto = `🌴 Avisos de Férias da Squad Novas Apurações`;

  let itensListaHtml = "";
  let itensListaTexto = "";

  listaFerias.forEach(item => {
    itensListaHtml += `
      <li style="margin-bottom: 8px;">
        <strong>${item.nome}</strong> entra de férias em <strong>${item.diasFaltantes} dias</strong> 🧳 
        <span style="color: #555;">(${item.inicio} até ${item.fim})</span>
      </li>
    `;
    
    itensListaTexto += `• ${item.nome} entra de férias em ${item.diasFaltantes} dias (${item.inicio} até ${item.fim})\n`;
  });

  const htmlBody = `
    <div style="max-width: 500px; margin: 0 auto; font-family: Arial, sans-serif; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden; background-color: #ffffff;">
      <!-- Cabeçalho -->
      <div style="padding: 16px 20px; border-bottom: 1px solid #f0f0f0;">
        <h2 style="margin: 0; font-size: 18px; color: #202124;">Avisos de Férias</h2>
        <p style="margin: 4px 0 0 0; font-size: 13px; color: #5f6368;">Squad Novas Apurações</p>
      </div>

      <!-- Corpo -->
      <div style="padding: 20px;">
        <p style="text-align: center; font-size: 16px; margin-top: 0;">
          🏖️ 😎 <strong>Avisos de Férias da Squad Novas Apurações</strong> 😎 🏖️
        </p>

        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; margin: 15px 0;">
          <p style="margin-top: 0; font-size: 14px; font-weight: bold; color: #3c4043;">🗓️ Próximas férias:</p>
          <ul style="margin: 0; padding-left: 20px; font-size: 14px; color: #3c4043;">
            ${itensListaHtml}
          </ul>
        </div>

        <!-- Imagem -->
        <div style="text-align: center; margin-top: 15px;">
          <img src="${IMAGE_URL_FERIAS}" alt="Férias" style="width: 100%; max-width: 450px; border-radius: 12px; display: block; margin: 0 auto;" />
        </div>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: CONFIG_NOTIFICAO.emailDestino,
    subject: assunto,
    body: `Avisos de Férias da Squad Novas Apurações:\n\n${itensListaTexto}`,
    htmlBody: htmlBody
  });
}
// ---------------------------------------------------------------------------
// Envio do Card para o Google Chat (Com imagem pública de praia ao pôr do sol)
// ---------------------------------------------------------------------------
function enviarGoogleChatFerias(listaFerias) {
  let textoLista = "";
  listaFerias.forEach(item => {
    textoLista += `• *${item.nome}* entra de férias em *${item.diasFaltantes} dias* 🧳 (${item.inicio} até ${item.fim})\n`;
  });

  // URL pública direta de imagem de praia ao pôr do sol
  const imageUrlChat = "https://plus.unsplash.com/premium_vector-1714618814924-d14c9760390b?q=80&w=2638&auto=format&fit=crop&ixlib=rb-4.1.0&ixid=M3wxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8fA%3D%3D";

  const payload = {
    cardsV2: [
      {
        cardId: "notificacaoFerias",
        card: {
          header: {
            title: "Avisos de Férias",
            subtitle: "Squad Novas Apurações",
            imageUrl: "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/beach_access/default/24px.svg",
            imageType: "CIRCLE"
          },
          sections: [
            {
              widgets: [
                {
                  textParagraph: {
                    text: "🏖️ 😎 *Avisos de Férias da Squad Novas Apurações* 😎 🏖️"
                  }
                },
                {
                  textParagraph: {
                    text: `🗓️ *Próximas férias:*\n${textoLista}`
                  }
                },
                {
                  image: {
                    imageUrl: imageUrlChat,
                    altText: "Praia de Férias"
                  }
                }
              ]
            }
          ]
        }
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(CONFIG_NOTIFICAO.webhookGoogleChat, options);
}