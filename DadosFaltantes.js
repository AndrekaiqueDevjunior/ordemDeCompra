function verificarDadosFaltantes(linhaEditada) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaControlePedidos = ss.getSheetByName('Controle Pedidos');

  // Mapeamento das colunas críticas com descrições amigáveis
  const colunasCriticas = {
    'CNPJ (DP)': 120, // Coluna DP
    'Fornecedor (DQ)': 121, // Coluna DQ
    'Email fornecedor (DR)': 122, // Coluna DR
    'Contato SKYPE / TEAMS fornecedor (DT)': 123, // Coluna DT (Skype/Teams)
    'Telefone vendedor (DU)': 124  // Coluna DU
  };

  const dadosLinha = abaControlePedidos.getRange(linhaEditada, 1, 1, 150).getValues()[0];
  const dadosFaltantes = [];

  // Verifica cada campo crítico
  for (const [descricao, coluna] of Object.entries(colunasCriticas)) {
    if (!dadosLinha[coluna - 1] || dadosLinha[coluna - 1].toString().trim() === '') {
      dadosFaltantes.push(descricao);
    }
  }

  return dadosFaltantes;
}

function enviarEmailParaFornecedor(numeroPedido) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaOC = ss.getSheetByName("ORDEM DE COMPRA");
    
    // Captura do e-mail na célula D15:H15
    const emailRange = abaOC.getRange("D15:H15").getValues(); // array 2D
    const emailFornecedor = emailRange.flat().filter(String)[0]; // Primeiro e-mail não vazio

    Logger.log("Email do fornecedor: " + emailFornecedor);

    if (!emailFornecedor) {
      Logger.log("Nenhum e-mail do fornecedor encontrado na célula D15:H15.");
      return;
    }

    // Validação do e-mail
    if (!/@/.test(emailFornecedor)) {
      Logger.log("O valor obtido em D15:H15 não parece ser um e-mail válido: " + emailFornecedor);
      return;
    }

    // Corpo do e-mail (exemplo, personalize se quiser)
    const assunto = `Ordem de Compra ${numeroPedido}`;
    const corpoEmail = `
      <p>Prezado fornecedor,</p>
      <p>Segue em anexo a Ordem de Compra número <strong>${numeroPedido}</strong>.</p>
      <p>Atenciosamente,<br>Equipe de Compras</p>
    `;

    // Arquivo PDF gerado previamente
    const pastaPDF = DriveApp.getFolderById("1ldG-RtOThn00-5JUJu_ubHkWWxMg28lH"); // substitua pelo ID da pasta se necessário
    const arquivos = pastaPDF.getFilesByName(`${numeroPedido}.pdf`);

    if (!arquivos.hasNext()) {
      Logger.log(`PDF da OC ${numeroPedido} não encontrado na pasta.`);
      return;
    }

    const pdf = arquivos.next();

    // Enviar e-mail
    MailApp.sendEmail({
      to: emailFornecedor,
      subject: assunto,
      htmlBody: corpoEmail,
      attachments: [pdf.getAs(MimeType.PDF)]
    });

    Logger.log("E-mail enviado com sucesso para o fornecedor: " + emailFornecedor);

  } catch (erro) {
    Logger.log("Erro ao enviar e-mail para o fornecedor: " + erro);
  }
}
