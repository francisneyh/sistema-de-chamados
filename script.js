// Função principal chamada ao receber novas respostas
function onFormSubmit(e) {
  var data = e.namedValues;
  var sheet = e.source.getActiveSheet();
  var linhaResposta = e.range.getRow();

  var dataHora = data["Carimbo de data/hora"]?.[0];
  var emailSolicitante = data["Endereço de e-mail"]?.[0];
  var nomeSolicitante = data["Nome do Solicitante"]?.[0];
  var local = data["Local"]?.[0];
  var tipo = data["Tipo"]?.[0];
  var categoria = data["Categoria"]?.[0];
  var urgencia = data["Urgência"]?.[0];
  var tituloChamado = data["Título do Chamado"]?.[0];
  var descricaoChamado = data["Descrição do Chamado"]?.[0];
  var linksAnexo = data["Anexo"];

  // Atribuir o próximo ID sequencial
  var idChamado = linhaResposta - 1;
  sheet.getRange(linhaResposta, 1).setValue(idChamado);

  // Define o status inicial como "Aberto"
  var colunaStatus = sheet.getLastColumn();
  sheet.getRange(linhaResposta, colunaStatus).setValue("Aberto");

  // E-mail do responsável (substitua pelo e-mail de quem deve atender os chamados)
  var emailResponsavel = "francisney.ti@guaxupe.mg.gov.br";

  // Captura os arquivos enviados via Google Forms (campo de upload de arquivos)
  // Processa anexos
  var arquivosAnexados = processarAnexos(linksAnexo);

  // Enviar notificação ao responsável (suporte técnico ou administrador)
  // Conteúdo para o responsável
  var htmlResponsavel = `
    <p>Um novo chamado foi registrado com os seguintes dados:</p>
    <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
      <tr><td><strong>ID do Chamado:</strong></td><td>${idChamado}</td></tr>
      <tr><td><strong>Data/Hora:</strong></td><td>${dataHora}</td></tr>
      <tr><td><strong>Solicitante:</strong></td><td>${nomeSolicitante}</td></tr>
      <tr><td><strong>E-mail:</strong></td><td>${emailSolicitante}</td></tr>
      <tr><td><strong>Local:</strong></td><td>${local}</td></tr>
      <tr><td><strong>Tipo:</strong></td><td>${tipo}</td></tr>
      <tr><td><strong>Categoria:</strong></td><td>${categoria}</td></tr>
      <tr><td><strong>Urgência:</strong></td><td>${urgencia}</td></tr>
      <tr><td><strong>Título:</strong></td><td>${tituloChamado}</td></tr>
      <tr><td><strong>Descrição:</strong></td><td>${descricaoChamado}</td></tr>
    </table>
    ${linksAnexo ? `<p><strong>Anexos:</strong><br>${linksAnexo.join('<br>')}</p>` : ""}
  `;

  // Enviar confirmação para o solicitante
  // Corpo do e-mail para o solicitante
  var htmlSolicitante = `
    <p>Olá <strong>${nomeSolicitante}</strong>,</p>
    <p>Seu chamado foi registrado com sucesso. Aqui estão os dados enviados:</p>
    <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
      <tr><td><strong>ID do Chamado:</strong></td><td>${idChamado}</td></tr>
      <tr><td><strong>Local:</strong></td><td>${local}</td></tr>
      <tr><td><strong>Tipo:</strong></td><td>${tipo}</td></tr>
      <tr><td><strong>Categoria:</strong></td><td>${categoria}</td></tr>
      <tr><td><strong>Urgência:</strong></td><td>${urgencia}</td></tr>
      <tr><td><strong>Título:</strong></td><td>${tituloChamado}</td></tr>
      <tr><td><strong>Descrição:</strong></td><td>${descricaoChamado}</td></tr>
    </table>
    <p>Nossa equipe de TI irá analisar a sua solicitação e responderá em breve.</p>
  `;

  // Enviar e-mail de chamado para o responsável técnico
  try {
    MailApp.sendEmail({
      to: emailResponsavel,
      replyTo: emailSolicitante,
      subject: "🚨 Novo Chamado Registrado - #" + idChamado + " - " + tituloChamado,
      htmlBody: gerarEmailHTML("Novo Chamado de TI", htmlResponsavel, "#d93025"),
      // attachments: arquivosAnexados
    });
  } catch (err) {
    Logger.log("Erro ao enviar e-mail ao responsável: " + err.message);
    registrarErro(err, "Envio de E-mail para Responsável", {
      idChamado, emailResponsavel, emailSolicitante, tituloChamado
    });
  }

  // Enviar e-mail de confirmação para o solicitante
  try {
    MailApp.sendEmail({
      to: emailSolicitante,
      subject: "✅ Confirmação de Chamado Registrado",
      htmlBody: gerarEmailHTML("Chamado Registrado com Sucesso", htmlSolicitante, "#1a73e8")
    });
  } catch (err) {
    Logger.log("Erro ao enviar e-mail ao solicitante: " + err.message);
    registrarErro(err, "Envio de E-mail para Solicitante", {
      idChamado, emailSolicitante, tituloChamado
    });
  }

  aplicarFormatacaoUrgencia(sheet);
  aplicarFormatacaoStatus(sheet);
  centralizarColunas(sheet); // <- Centralizar colunas ao final

}

// 🔧 Gerar HTML do e-mail no estilo Google
function gerarEmailHTML(titulo, corpoHTML, corHeader) {
  return `
    <div style="font-family: 'Roboto', sans-serif; background-color: #f1f3f4; padding: 20px;">
      <div style="max-width: 600px; margin: auto; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        <div style="background-color: ${corHeader}; color: white; padding: 16px 24px; text-align: center;">
          <h2 style="margin: 0; font-size: 22px;">${titulo}</h2>
        </div>
        <div style="padding: 24px;">${corpoHTML}</div>
        <div style="background-color: #f8f9fa; text-align: center; padding: 16px; font-size: 12px; color: #5f6368;">
          Diretoria de Tecnologia da Informação – Prefeitura de Guaxupé<br>
          <a href="https://www.guaxupe.mg.gov.br" style="color: #1a73e8; text-decoration: none;">www.guaxupe.mg.gov.br</a>
        </div>
      </div>
    </div>`;
}

// 🛠️ Processamento dos arquivos anexados
// Captura os arquivos enviados via Google Forms (campo de upload de arquivos)
function processarAnexos(linksAnexo) {
  var arquivos = [];
  if (linksAnexo && linksAnexo.length > 0) {
    for (var i = 0; i < linksAnexo.length; i++) {
      try {
        var match = linksAnexo[i].match(/[-\w]{25,}/);
        if (match) {
          var fileId = match[0];
          var file = DriveApp.getFileById(fileId);
          var mimeType = file.getMimeType();
          if (
            mimeType === MimeType.GOOGLE_DOCS ||
            mimeType === MimeType.GOOGLE_SHEETS ||
            mimeType === MimeType.GOOGLE_SLIDES
          ) {
            arquivos.push(file.getAs(MimeType.PDF));
          } else {
            arquivos.push(file.getBlob());
          }
        }
      } catch (error) {
        Logger.log("Erro ao processar anexo: " + error.message);
        registrarErro(error, "Processamento de Anexo", { link: linksAnexo[i] });
      }
    }
  }
  return arquivos;
}

// 🔧 Formatação da coluna urgência
function aplicarFormatacaoUrgencia(sheet) {
  const colunaUrgencia = 8; // Coluna H
  const ultimaLinha = sheet.getLastRow();

   // Limpa a formatação anterior
  sheet.getRange(2, colunaUrgencia, ultimaLinha - 1).setBackground(null);

  for (let i = 2; i <= ultimaLinha; i++) {
    const urgencia = sheet.getRange(i, colunaUrgencia).getValue().toString().toLowerCase();
    let cor = null;

    if (urgencia === "alta") cor = "#e57373";       // Vermelho Claro
    else if (urgencia === "média") cor = "#fff176 "; // Amarelo Claro
    else if (urgencia === "baixa") cor = "#81c784 "; // Verde Claro

    if (cor) {
      sheet.getRange(i, colunaUrgencia).setBackground(cor);
    }
  }
}

// 🔧 Formatação da coluna status
function aplicarFormatacaoStatus(sheet) {
  const colunaStatus = 12; // Substitua pelo número real da coluna "Status", Coluna 10
  const ultimaLinha = sheet.getLastRow();

   // Limpa a formatação anterior
  sheet.getRange(2, colunaStatus, ultimaLinha - 1).setBackground(null);

  for (let i = 2; i <= ultimaLinha; i++) {
    const status = sheet.getRange(i, colunaStatus).getValue().toString().trim().toLowerCase();
    
    let cor = null;

    switch (status) {
      case "aberto":
        cor = "#42a5f5"; // Azul Claro
        break;
      case "em andamento":
        cor = "#fff176"; // Amarelo Claro
        break;
      case "solucionado":
        cor = "#81c784"; // Verde Claro
        break;
      case "fechado":
        cor = "#bdbdbd"; // Cinza Claro
        break;
    }

    if (cor) {
      sheet.getRange(i, colunaStatus).setBackground(cor);
    }
  }
}

// 🔧 Centralizar colunas ID, Tipo, Urgência, Status
function centralizarColunas(sheet) {
  const ultimaLinha = sheet.getLastRow();
  const colunaID = 1; // Coluna A
  const colunaTipo = 6; // Coluna F
  const colunaUrgencia = 8; // Coluna H
  const colunaStatus = sheet.getLastColumn(); // Última coluna

  sheet.getRange(2, colunaID, ultimaLinha - 1).setHorizontalAlignment("center");
  sheet.getRange(2, colunaTipo, ultimaLinha - 1).setHorizontalAlignment("center");
  sheet.getRange(2, colunaUrgencia, ultimaLinha - 1).setHorizontalAlignment("center");
  sheet.getRange(2, colunaStatus, ultimaLinha - 1).setHorizontalAlignment("center");
}

// 🔧 Registrar erros e notificar administrador
function registrarErro(erro, contexto, dadosExtras) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var abaLog = planilha.getSheetByName("Log de Erros");
    var dataHora = new Date();

    if (!abaLog) {
      abaLog = planilha.insertSheet("Log de Erros");
      abaLog.appendRow(["Data/Hora", "Contexto", "Erro", "Dados Extras"]);
    }

    abaLog.appendRow([
      dataHora,
      contexto,
      erro.message || erro.toString(),
      JSON.stringify(dadosExtras || {})
    ]);

    var emailAdmin = "francisney.ti@guaxupe.mg.gov.br";
    var assunto = "⚠️ Erro no Script - " + contexto;
    var corpoHTML = `
      <p><strong>Um erro foi registrado no script do formulário de chamados:</strong></p>
      <table style="font-size: 14px; border-collapse: collapse;">
        <tr><td><strong>Data/Hora:</strong></td><td>${dataHora}</td></tr>
        <tr><td><strong>Contexto:</strong></td><td>${contexto}</td></tr>
        <tr><td><strong>Erro:</strong></td><td><pre>${erro.message || erro.toString()}</pre></td></tr>
        <tr><td><strong>Dados Extras:</strong></td><td><pre>${JSON.stringify(dadosExtras || {}, null, 2)}</pre></td></tr>
      </table>
      <p>Verifique a aba <strong>Log de Erros</strong> na planilha para mais detalhes.</p>
    `;

    MailApp.sendEmail({
      to: emailAdmin,
      subject: assunto,
      htmlBody: corpoHTML
    });

  } catch (logErro) {
    Logger.log("Falha ao registrar ou notificar erro: " + logErro.message);
  }
}