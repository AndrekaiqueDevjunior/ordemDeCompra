

function pesquisarPedido(pedidoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetControle = ss.getSheetByName('Controle de Pedidos');
  const sheetHistorico = ss.getSheetByName('Histórico de Pesquisa');
  
  // Verifica se o pedido já está no histórico
  const dadosHistorico = sheetHistorico.getDataRange().getValues();
  for (let i = 1; i < dadosHistorico.length; i++) {
    if (dadosHistorico[i][0] == pedidoId) {
      Logger.log('Pedido encontrado no histórico!');
      return;  // Pedido já está registrado no histórico, não faz nada
    }
  }

  // Buscar dados do pedido no Controle de Pedidos
  const dadosControle = sheetControle.getDataRange().getValues();
  let pedidoEncontrado = null;
  
  for (let i = 1; i < dadosControle.length; i++) {
    if (dadosControle[i][0] == pedidoId) {
      pedidoEncontrado = dadosControle[i];
      break;
    }
  }

  if (!pedidoEncontrado) {
    Logger.log('Pedido não encontrado!');
    return;  // Se o pedido não for encontrado
  }

  // Adiciona os dados do pedido no histórico
  const dadosPedido = [
    pedidoEncontrado[0], // Nome Fantasia
    pedidoEncontrado[1], // CNPJ/CPF
    pedidoEncontrado[2], // Bairro
    pedidoEncontrado[3], // Cidade
    pedidoEncontrado[4], // E-mail
    pedidoEncontrado[5], // E-mail Vendedor
    pedidoEncontrado[6], // Razão Social
    pedidoEncontrado[7], // Endereço
    pedidoEncontrado[8], // CEP
    pedidoEncontrado[9], // Estado
    pedidoEncontrado[10], // Telefone
    pedidoEncontrado[11], // Skype / Teams / Whatsapp
    pedidoEncontrado[12], // Qtd
    pedidoEncontrado[13], // Cód. Forn.
    pedidoEncontrado[14], // Unidade
    pedidoEncontrado[15], // Descrição
    pedidoEncontrado[16], // Valor Un.
    pedidoEncontrado[17], // Desconto (%)
    pedidoEncontrado[18], // IPI (%)
    pedidoEncontrado[19], // ICMS ST (R$)
    pedidoEncontrado[20], // Previsão de Entrega
    pedidoEncontrado[21], // Valor Total
    pedidoEncontrado[22], // TOTAL
    pedidoEncontrado[23], // Frete por Conta
    pedidoEncontrado[24], // Valor Frete
    pedidoEncontrado[25], // Valor Total
    pedidoEncontrado[26], // Observações
    pedidoEncontrado[27], // Comprador
    pedidoEncontrado[28], // Emissão
    pedidoEncontrado[29], // Previsão de Entrega
  ];

  // Adiciona no Histórico de Pesquisa
  sheetHistorico.appendRow(dadosPedido);
  Logger.log('Pedido registrado no histórico de pesquisa!');
}
