function pintarLinhasComStatus() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leitor - Controle Pedidos');
  var dados = planilha.getDataRange().getValues();

  // Percorrer cada linha
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][4] != "") { // Verifica se a célula não está vazia
      if (dados[i][4] == "Cancelado") {
        // Pintar linha inteira de vermelho
        planilha.getRange(i + 1, 1, 1, planilha.getLastColumn()).setBackground("#FF0000");
      } else if (dados[i][4] == "Novo") {
        // Pintar linha inteira de verde
        planilha.getRange(i + 1, 1, 1, planilha.getLastColumn()).setBackground("#00FF00");
      } else if (dados[i][4] == "Alterado") {
        // Pintar linha inteira de amarelo
        planilha.getRange(i + 1, 1, 1, planilha.getLastColumn()).setBackground("#FFFF00");
      } else if (dados[i][4] == "Falha") {
        // Pintar linha inteira de azul claro para "Falha"
        planilha.getRange(i + 1, 1, 1, planilha.getLastColumn()).setBackground("#ADD8E6");
      }
    }
  }
}
