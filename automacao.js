function copiarDados() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaOrigem = planilha.getSheetByName("Leitor - Controle Pedidos");
  var abaDestino = planilha.getSheetByName("Controle Pedidos");

  // Verifica se a aba de origem está vazia
  if (abaOrigem.getLastRow() < 2) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Aviso', 'A aba de origem está vazia. Não há dados para copiar.', ui.ButtonSet.OK);
    return;
  }

  Logger.log("Iniciando a operação de cópia de dados...");

  var linhaAtual = 2;
  var ultimaColunaOrigem = abaOrigem.getLastColumn();

  // Obtém todos os dados da aba de origem
  var dadosOrigem = abaOrigem.getRange("A" + linhaAtual + ":DI" + abaOrigem.getLastRow()).getValues();

  // Verifica a última linha preenchida no intervalo da aba de destino (A até última coluna da origem)
  var intervaloDestino = abaDestino.getRange(1, 1, abaDestino.getLastRow(), ultimaColunaOrigem).getValues();
  var ultimaLinhaDestino = intervaloDestino.length - intervaloDestino.reverse().findIndex(function (row) {
    return row.some(function (cell) { return cell !== ""; });
  });

  // Copia os dados para a aba de destino
  var rangeDestino = abaDestino.getRange(ultimaLinhaDestino + 1, 1, dadosOrigem.length, ultimaColunaOrigem);
  rangeDestino.setValues(dadosOrigem);

  // Define a cor de fundo das células na aba de destino
  var corVerdeEscuro = "#ADD8E6"; // Cor azul claro
  rangeDestino.setBackground(corVerdeEscuro);

  // Limpa a aba de origem
  var rangeOrigem = abaOrigem.getRange("A" + linhaAtual + ":DI" + abaOrigem.getLastRow());
  rangeOrigem.clearContent();

  Logger.log("Operação de cópia de dados concluída com sucesso.");
}