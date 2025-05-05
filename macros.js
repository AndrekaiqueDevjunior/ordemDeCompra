

function recolherGRUPO() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('P10').activate();
  spreadsheet.getActiveSheet().getColumnGroup(14, 1).collapse();
};

function recolherGrupoS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B10:S10').activate();
  spreadsheet.getActiveSheet().getColumnGroup(19, 1).collapse();
};



function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B10:S10').activate();
  spreadsheet.getActiveSheet().getColumnGroup(19, 1).collapse();
};

function recolherGrupoS() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName('ORDEM DE COMPRA'); // Substitua pelo nome correto da aba
  var grupoColunas = aba.getColumnGroup(19); // Verifica se o grupo da coluna 19 existe

  if (grupoColunas) {
    Logger.log("Grupo de colunas encontrado.");
  } else {
    Logger.log("Grupo de colunas não encontrado.");
  }
}

 

function liberarGrupo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S25').activate();
  spreadsheet.getActiveSheet().getColumnGroup(19, 1).expand();
};

function trocandoPagina() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Controle Pedidos'), true);
  spreadsheet.getCurrentCell().setValue('OC Gerada');
  spreadsheet.getRange('EP10').activate();
};






function recolhercolunaS() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colIndex = 19; // Coluna S
  var depth = 1;

  try {
    var group = sheet.getColumnGroup(colIndex, depth);
    if (group) {
      group.collapse();
      Logger.log("Grupo de colunas colapsado com sucesso.");
    } else {
      Logger.log("Nenhum grupo encontrado na coluna S com profundidade 1.");
    }
  } catch (e) {
    Logger.log("Erro ao tentar colapsar o grupo: " + e.message);
  }
}


function liberarGrupoS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S:S').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('Q1'));
  spreadsheet.getActiveSheet().getColumnGroup(19, 1).expand();
};