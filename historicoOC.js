function enviarDadosOrdemCompraParaHistorico() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = spreadsheet.getSheetByName('ORDEM DE COMPRA');
  var abaHistorico = spreadsheet.getSheetByName('Histórico');
  var abaControlePedidos = spreadsheet.getSheetByName('Controle Pedidos');

  if (!abaProposta || !abaHistorico) {
    SpreadsheetApp.getUi().alert('Uma das abas especificadas não foi encontrada.');
    return;
  }

  var fusoHorario = "America/Sao_Paulo";
  var dataHoraAtual = new Date();
  var dataHoraBrasilia = Utilities.formatDate(dataHoraAtual, fusoHorario, "dd/MM/yyyy HH:mm:ss");

  // Gerar o número da proposta

  // Coletar dados da ordem de compra
  var dados = {
    numeroProposta: abaProposta.getRange('Q1:S1').getValue(),
    linkPDF: abaProposta.getRange('AO22').getValue(),
    nomeFantasia: abaProposta.getRange('D11:H11').getValue(),
    cnpjCpf: abaProposta.getRange('D12:H12').getValue(),
    bairro: abaProposta.getRange('D13:H13').getValue() + ' ' + abaProposta.getRange('F5').getValue(),
    cidade: abaProposta.getRange('D14:H14').getValue(),
    emailComprador: abaProposta.getRange('D15:H15').getValue() + abaProposta.getRange('H13').getValue(),
    emailVendedor: abaProposta.getRange('D16:H16').getValue() + abaProposta.getRange('H16').getValue(),
    razaoSocial: abaProposta.getRange('J11:S11').getValue() + ' ' + abaProposta.getRange('N9').getValue(),
    endereco: abaProposta.getRange('J12:S12').getValue(),
    cep: abaProposta.getRange('J13:S13').getValue(),
    estado: abaProposta.getRange('J14:S14').getValue(),
    telefone: abaProposta.getRange('J15:S15').getValue(),
    skypeTeamsWhatsapp: abaProposta.getRange('J16:S16').getValue(),
    quantidade: abaProposta.getRange('B23').getValue(),
    codigoFornecedor: abaProposta.getRange('C23').getValue(),
    unidade: abaProposta.getRange('D23:E23').getValue(),
    descricao: abaProposta.getRange('F23:I23').getValue(),
    valorUnidade: abaProposta.getRange('J23').getValue(),
    desconto: abaProposta.getRange('K23').getValue(),
    ipi: abaProposta.getRange('L23').getValue(),
    icmsST: abaProposta.getRange('M23:Q23').getValue(),
    previsaoEntrega: abaProposta.getRange('R23').getValue(),
    valorTotal: abaProposta.getRange('S23').getValue(),
    total: abaProposta.getRange('S29').getValue(),
    fretePorConta: abaProposta.getRange('B35:G35').getValue(),
    valorFrete: abaProposta.getRange('H35:J35').getValue(),
    valorTotal: abaProposta.getRange('M35:S35').getValue(),
    observacoes: abaProposta.getRange('B46:S51').getValue(),
    comprador: abaProposta.getRange('C56:E56').getValue(),
    emissao: abaProposta.getRange('C7:H7').getValue(),
    previsaoEntrega: abaProposta.getRange('J7').getValue()
  };

  // Obter o link do PDF gerado na aba "Controle Pedidos", coluna ES (índice 149)
  var linkPDF = abaHistorico.getRange('AO22').getValue();
//  abaHistorico.getRange('J' + linhaHistorico).setValue(dados.nomeFantasia);

  // Obter a próxima linha disponível na aba "Histórico"
  var linhaHistorico = abaHistorico.getLastRow() + 1;

  // Preencher a aba "Histórico" com os dados
  abaHistorico.getRange('A' + linhaHistorico).setValue(new Date()); // Data e hora da execução
  abaHistorico.getRange('B' + linhaHistorico).setValue(dataHoraBrasilia);
  abaHistorico.getRange('H' + linhaHistorico).setValue(linkPDF); // Link do PDF vindo de "Controle Pedidos"
  abaHistorico.getRange('I' + linhaHistorico).setValue(dados.numeroProposta); // Número da proposta gerada
  abaHistorico.getRange('J' + linhaHistorico).setValue(dados.nomeFantasia);
  abaHistorico.getRange('K' + linhaHistorico).setValue(dados.cnpjCpf);
  abaHistorico.getRange('L' + linhaHistorico).setValue(dados.bairro);
  abaHistorico.getRange('M' + linhaHistorico).setValue(dados.cidade);
  abaHistorico.getRange('N' + linhaHistorico).setValue(dados.emailComprador);
  abaHistorico.getRange('O' + linhaHistorico).setValue(dados.emailVendedor);
  abaHistorico.getRange('P' + linhaHistorico).setValue(dados.razaoSocial);
  abaHistorico.getRange('Q' + linhaHistorico).setValue(dados.endereco);
  abaHistorico.getRange('R' + linhaHistorico).setValue(dados.cep);
  abaHistorico.getRange('S' + linhaHistorico).setValue(dados.estado);
  abaHistorico.getRange('T' + linhaHistorico).setValue(dados.telefone);
  abaHistorico.getRange('U' + linhaHistorico).setValue(dados.skypeTeamsWhatsapp);
  abaHistorico.getRange('V' + linhaHistorico).setValue(dados.quantidade);
  abaHistorico.getRange('W' + linhaHistorico).setValue(dados.codigoFornecedor);
  abaHistorico.getRange('X' + linhaHistorico).setValue(dados.unidade);
  abaHistorico.getRange('Y' + linhaHistorico).setValue(dados.descricao);
  abaHistorico.getRange('Z' + linhaHistorico).setValue(dados.valorUnidade);
  abaHistorico.getRange('AA' + linhaHistorico).setValue(dados.desconto);
  abaHistorico.getRange('AB' + linhaHistorico).setValue(dados.ipi);
  abaHistorico.getRange('AC' + linhaHistorico).setValue(dados.icmsST);
  abaHistorico.getRange('AD' + linhaHistorico).setValue(dados.previsaoEntrega);
  abaHistorico.getRange('AE' + linhaHistorico).setValue(dados.valorTotal);
  abaHistorico.getRange('AF' + linhaHistorico).setValue(dados.total);
  abaHistorico.getRange('AG' + linhaHistorico).setValue(dados.fretePorConta);
  abaHistorico.getRange('AH' + linhaHistorico).setValue(dados.valorFrete);
  abaHistorico.getRange('AI' + linhaHistorico).setValue(dados.valorTotal);
  abaHistorico.getRange('AJ' + linhaHistorico).setValue(dados.observacoes);
  abaHistorico.getRange('AK' + linhaHistorico).setValue(dados.comprador);
  abaHistorico.getRange('AL' + linhaHistorico).setValue(dados.emissao);
  abaHistorico.getRange('AM' + linhaHistorico).setValue(dados.previsaoEntrega);
}



function gerarPDFTemporarioLimpo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const origem = ss.getSheetByName("ORDEM DE COMPRA");
  const tempName = "TEMP_PDF_PREVIEW";

  // Apagar temporária antiga
  const oldTemp = ss.getSheetByName(tempName);
  if (oldTemp) ss.deleteSheet(oldTemp);

  // Criar nova temporária
  const temp = ss.insertSheet(tempName);

  // Copiar conteúdo de A1:S58
  const rangeOrigem = origem.getRange("A1:S58");
  rangeOrigem.copyTo(temp.getRange("A1"), { contentsOnly: true });

  // Ajustar largura das colunas (igual à aba original)
  for (let col = 1; col <= 19; col++) {
    const largura = origem.getColumnWidth(col);
    temp.setColumnWidth(col, largura);
  }

  // Apagar colunas além de S
  const maxCol = temp.getMaxColumns();
  if (maxCol > 19) temp.deleteColumns(20, maxCol - 19);

  // Apagar linhas além da 58
  const maxRow = temp.getMaxRows();
  if (maxRow > 58) temp.deleteRows(59, maxRow - 58);

  // Exportar para visualização
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/export?format=pdf' +
    '&gid=' + temp.getSheetId() +
    '&portrait=true' +
    '&size=A4' +
    '&fitw=true' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenumbers=false' +
    '&top_margin=0.1' +
    '&bottom_margin=0.1' +
    '&left_margin=0.1' +
    '&right_margin=0.1' +
    '&scale=3';

  const previewUrl = url + '&access_token=' + ScriptApp.getOAuthToken();
  Logger.log("🔍 PDF Preview: " + previewUrl);
  SpreadsheetApp.getUi().alert('✅ PDF temporário gerado. Veja o link no LOG.\nAbra o Logger com Ctrl + Enter.');
}
