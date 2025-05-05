function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gera O.C e Consultar CNPJ')
    .addSeparator()
    .addItem('Buscar dados de CNPJ', 'abrirFormCNPJ')
    .addSeparator()
    .addItem('Gerar ORDEM DE COMPRA', 'enviarPdfParaDriveiexportarParaHistorico')
    .addSeparator()
    .addItem('Buscar Pedidos', 'preencherDadosProdutos')
    .addSeparator()
    .addToUi();
}

function enviarPdfParaDriveiexportarParaHistorico() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaProposta = spreadsheet.getSheetByName('ORDEM DE COMPRA');
  var abaHistorico = spreadsheet.getSheetByName('Histórico');
  var abaControlePedidos = spreadsheet.getSheetByName('Controle Pedidos');

  if (!abaProposta || !abaHistorico) {
    SpreadsheetApp.getUi().alert('Uma das abas especificadas não foi encontrada.');
    return;
  }


  var fusoHorario = "America/Sao_Paulo";

  // Obter a data e hora atuais com o fuso horário correto
  var dataHoraAtual = new Date();
  var dataHoraBrasilia = Utilities.formatDate(dataHoraAtual, fusoHorario, "dd/MM/yyyy HH:mm:ss");

  // Definir a célula onde a data e hora serão exibidas (exemplo: célula A1)

  //abaHistorico.getRange('B' + linhaHistorico).setValue(dataHoraBrasilia);

  //ocultarColunaM()
  recolherGrupoS();

  // Obter o número da proposta
  var hoje = new Date();
  var ano = hoje.getFullYear().toString().slice(-2); // Últimos 2 dígitos do ano
  var mes = (hoje.getMonth() + 1).toString().padStart(2, '0'); // Mês
  var dia = hoje.getDate().toString().padStart(2, '0'); // Dia
  var sequencia = (abaHistorico.getLastRow() + 0).toString().padStart(4, '0'); // Sequência
  var revisao = '0'; // Revisão sempre 0

  var numeroProposta = `${ano}${mes}${dia}${sequencia}-${revisao}`;
  abaProposta.getRange('Q1:S1').setValue(numeroProposta); // ORDEM DE COMPRA // antigo era I1

  abaProposta.getRange('C7:H7').setValue(dataHoraBrasilia);
  // Coletar dados da aba "Proposta Novus" // coletar dados da ordem de compra
  var dados = {
    nomeFantasia: abaProposta.getRange('D11:H11').getValue(), //NOME FANTASIA
    cnpjCpf: abaProposta.getRange('D12:H12').getValue(), // CNPJ/CPF:	
    bairro: abaProposta.getRange('D13:H13').getValue() + ' ' + abaProposta.getRange('F5').getValue(), // Bairro:	
    cidade: abaProposta.getRange('D14:H14').getValue(), //Cidade	

    emailComprador: abaProposta.getRange('D15:H15').getValue() + abaProposta.getRange('H13').getValue(), //E-mail:	
    emailVendedor: abaProposta.getRange('D16:H16').getValue() + abaProposta.getRange('H16').getValue(), //E-mail E-mail Vendedor:

    razaoSocial: abaProposta.getRange('J11:S11').getValue() + ' ' + abaProposta.getRange('N9').getValue(), //Razão Social:

    endereco: abaProposta.getRange('J12:S12').getValue(), //Endereço:

    cep: abaProposta.getRange('J13:S13').getValue(), //CEP:

    estado: abaProposta.getRange('J14:S14').getValue(), //Estado:



    telefone: abaProposta.getRange('J15:S15').getValue(), //Telefone:

    skypeTeamsWhatsapp: abaProposta.getRange('J16:S16').getValue(),  // Skype / Teams / Whatsapp:

    quantidade: abaProposta.getRange('B23').getValue(), //Qtd.

    codigoFornecedor: abaProposta.getRange('C23').getValue(), //Cód. Forn.
    unidade: abaProposta.getRange('D23:E23').getValue(), //Unidade
    descricao: abaProposta.getRange('F23:I23').getValue(), //Descrição


    valorUnidade: abaProposta.getRange('J23').getValue(), //Valor Un.
    desconto: abaProposta.getRange('K23').getValue(), //Desconto (%):

    ipi: abaProposta.getRange('L23').getValue(), //IPI (%):

    icmsST: abaProposta.getRange('M23:Q23').getValue(), //ICMS ST (R$):

    // descontoUn: abaProposta.getRange('K17:L17').getValue(), //Desconto Un.
    previsaoEntrega: abaProposta.getRange('R23').getValue(), //Previsão Entrega

    valorTotal: abaProposta.getRange('S23').getValue(), //Valor Total

    total: abaProposta.getRange('S29').getValue(),  //TOTAL

    fretePorConta: abaProposta.getRange('B35:G35').getValue(), //Frete por Conta			

    valorFrete: abaProposta.getRange('H35:J35').getValue(), //Valor Frete		

    valorTotal: abaProposta.getRange('M35:S35').getValue(), //Valor Total da OC

    observacoes: abaProposta.getRange('B46:S51').getValue(), //OBSERVAÇÕES

    comprador: abaProposta.getRange('C56:E56').getValue(), //Comprador: 


    emissao: abaProposta.getRange('C7:H7').getValue(),

    previsaoEntrega: abaProposta.getRange('J7').getValue()
    /*
    ad11: abaProposta.getRange('L9').getValue(),

   dataProposta: formatarData(abaProposta.getRange('C6').getValue()) + ' ' + formatarData(abaProposta.getRange('C6').getValue()) + ' ' + formatarData(abaProposta.getRange('C12').getValue()) + ' ' + formatarData(abaProposta.getRange('D12').getValue()),
    validadeProposta: formatarData(abaProposta.getRange('J6:N6').getValue()), // Corrigido aqui // ANTIGO F12:L12
    condicaoPagamento: abaProposta.getRange('A15:E15').getValue(),
    frete: abaProposta.getRange('F15:G15').getValue(),
    transportadora: abaProposta.getRange('I15:L15').getValue(),
    valorTotalProposta: abaProposta.getRange('G46:L46').getValue(),
    impostos: abaProposta.getRange('G47:L47').getValues(),
    frete2: abaProposta.getRange('G48').getValue(),
    notasobrePrazoEntrega: abaProposta.getRange('G49:L49').getValue(),


    vendedor: abaProposta.getRange('G51:L51').getValue() + ' ' + abaProposta.getRange('H49').getValue() + ' ' + abaProposta.getRange('I49').getValue() + ' ' + abaProposta.getRange('J49').getValue() + ' ' + abaProposta.getRange('K49').getValue() + ' ' + abaProposta.getRange('L49').getValue(),

    emailVendedor: abaProposta.getRange('G52:L52').getValue() + ' ' + abaProposta.getRange('H50').getValue() + ' ' + abaProposta.getRange('I50').getValue() + ' ' + abaProposta.getRange('J50').getValue() + ' ' + abaProposta.getRange('K50').getValue() + ' ' + abaProposta.getRange('L50').getValue(),

    telefoneVendedor: abaProposta.getRange('G53:L53').getValue() + ' ' + abaProposta.getRange('H51').getValue() + ' ' + abaProposta.getRange('I51').getValue() + ' ' + abaProposta.getRange('J51').getValue() + ' ' + abaProposta.getRange('K51').getValue() + ' ' + abaProposta.getRange('L51').getValue(),



    whatsappVendedor: abaProposta.getRange('G54:L54').getValue() + ' ' + abaProposta.getRange('H52').getValue() + ' ' + abaProposta.getRange('I52').getValue() + ' ' + abaProposta.getRange('J52').getValue() + ' ' + abaProposta.getRange('K52').getValue() + ' ' + abaProposta.getRange('L52').getValue(),

    item: abaProposta.getRange('A19').getValue(),
    fabricante: abaProposta.getRange('B19').getValue(),
    codigo: abaProposta.getRange('C19').getValue(),
    descricao: abaProposta.getRange('D19').getValue(),
    ncm: abaProposta.getRange('E19').getValue(),
    quantidade: abaProposta.getRange('F19').getValue(),
    prazoEntrega: abaProposta.getRange('G19').getValue(),
    quantidadeEstoque: abaProposta.getRange('H19').getValue(),
    valorUnitario: abaProposta.getRange('I19').getValue(),
    desconto: abaProposta.getRange('J19').getValue(),
    valorUnitarioDesconto: abaProposta.getRange('K19').getValue(),
    valorTotalItem: abaProposta.getRange('L19').getValue(),
    numeroProposta: numeroProposta
    */
  };


  var linkPDF = gerarPDF(abaProposta, numeroProposta);

  // Obter dados dos itens
  var intervaloItens = abaProposta.getRange('A19:L' + abaProposta.getLastRow());
  var itens = intervaloItens.getValues();

  // Preencher a aba Histórico com os dados
  var linhaHistorico = abaHistorico.getLastRow() + 1; // Próxima linha disponível

  abaHistorico.getRange('A' + linhaHistorico).setValue(new Date()); // Data e hora da execução
  abaHistorico.getRange('B' + linhaHistorico).setValue(dataHoraBrasilia);
  //ColunaReserva1 C + linhaHistorico).setValue();
  //Coluna2 D + linhaHistorico).setValue();
  //Coluna3 E + linhaHistorico).setValue();
  //Coluna4 F + linhaHistorico).setValue();
  //Coluna5 G + linhaHistorico).setValue();
  abaHistorico.getRange('H' + linhaHistorico).setValue(linkPDF); // Link do PDF
  abaHistorico.getRange('I' + linhaHistorico).setValue(numeroProposta); // Número da proposta gerada 
  abaControlePedidos.getRange('A' + linhaHistorico).setValue(numeroProposta); // Número da proposta gerar na aba Controle de Pedidos
  abaHistorico.getRange('J' + linhaHistorico).setValue(dados.nomeFantasia); //  NOME FANTASIA  //D9:H9
  abaHistorico.getRange('K' + linhaHistorico).setValue(dados.cnpjCpf); //  CNPJ/CPF:  //D10:H10
  abaHistorico.getRange('L' + linhaHistorico).setValue(dados.bairro); //  Bairro: //D11:H11
  abaHistorico.getRange('M' + linhaHistorico).setValue(dados.cidade); //  Cidade //D12:H12
  abaHistorico.getRange('N' + linhaHistorico).setValue(dados.email); // E-mail: //D13:H13
  abaHistorico.getRange('O' + linhaHistorico).setValue(dados.emailVendedor); // E-mail Vendedor:
  abaHistorico.getRange('P' + linhaHistorico).setValue(dados.razaoSocial); // Razão Social:
  abaHistorico.getRange('Q' + linhaHistorico).setValue(dados.endereco); //  Endereço:
  abaHistorico.getRange('R' + linhaHistorico).setValue(dados.cep); //CEP:
  abaHistorico.getRange('S' + linhaHistorico).setValue(dados.estado); //  Estado:
  abaHistorico.getRange('T' + linhaHistorico).setValue(dados.telefone); //  Telefone:
  abaHistorico.getRange('U' + linhaHistorico).setValue(dados.skypeTeamsWhatsapp); //  Skype / Teams / Whatsapp:
  abaHistorico.getRange('V' + linhaHistorico).setValue(dados.quantidade); //Qtd.
  abaHistorico.getRange('W' + linhaHistorico).setValue(dados.codigoFornecedor); //Cód. Forn.
  abaHistorico.getRange('X' + linhaHistorico).setValue(dados.unidade); //Unidade

  abaHistorico.getRange('Y' + linhaHistorico).setValue(dados.descricao); //  Descrição
  abaHistorico.getRange('Z' + linhaHistorico).setValue(dados.valorUnidade); //  Valor Un.
  abaHistorico.getRange('AA' + linhaHistorico).setValue(dados.desconto); //  Desconto (%):
  abaHistorico.getRange('AB' + linhaHistorico).setValue(dados.ipi); // IPI (%):
  abaHistorico.getRange('AC' + linhaHistorico).setValue(dados.icmsST); //  ICMS ST (R$):
  abaHistorico.getRange('AD' + linhaHistorico).setValue(dados.previsaoEntrega); // Previsão  de Entrega
  abaHistorico.getRange('AE' + linhaHistorico).setValue(dados.valorTotal); //  Valor Total 
  abaHistorico.getRange('AF' + linhaHistorico).setValue(dados.total); //  TOTAL
  abaHistorico.getRange('AG' + linhaHistorico).setValue(dados.fretePorConta);  //Frete por Conta
  abaHistorico.getRange('AH' + linhaHistorico).setValue(dados.valorFrete); // Valor Frete
  abaHistorico.getRange('AI' + linhaHistorico).setValue(dados.valorTotal); // Valor Total

  abaHistorico.getRange('AJ' + linhaHistorico).setValue(dados.observacoes); // OBSERVAÇÕES
  abaHistorico.getRange('AK' + linhaHistorico).setValue(dados.comprador); // Comprador: 
  abaHistorico.getRange('AL' + linhaHistorico).setValue(dados.emissao); // EMISSÃO
  abaHistorico.getRange('AM' + linhaHistorico).setValue(dados.previsaoEntrega); // PREVISÃO ENTREGA
  abaHistorico.getRange('AN' + linhaHistorico).setValue(dados.numeroPedidoCompra); //  Nº Pedido de Compra (nomeFantasia)  OU CHAVE
  /*
  abaHistorico.getRange('AO' + linhaHistorico).setValue(dados.valorTotalProposta); // Valor total da proposta
  abaHistorico.getRange('AP' + linhaHistorico).setValue(dados.impostos); // Impostos
  abaHistorico.getRange('AQ' + linhaHistorico).setValue(dados.frete2); // Frete 2
  abaHistorico.getRange('AR' + linhaHistorico).setValue(dados.notasobrePrazoEntrega); // Nota sobre prazo de entrega
  abaHistorico.getRange('AS' + linhaHistorico).setValue(dados.vendedor); // Vendedor
  abaHistorico.getRange('AT' + linhaHistorico).setValue(dados.emailVendedor); // Email do vendedor
  abaHistorico.getRange('AU' + linhaHistorico).setValue(dados.telefoneVendedor); // Telefone do vendedor
  abaHistorico.getRange('AV' + linhaHistorico).setValue(dados.whatsappVendedor); // WhatsApp do vendedor
  */
  var intervaloProposta = abaProposta.getRange('A19:L44');
  var dadosProposta = intervaloProposta.getValues(); // Array 2D com os dados

  // Concatenar todos os valores em uma única linha
  var linhaConcatenada = [];

  for (var i = 0; i < dadosProposta.length; i++) {
    for (var j = 0; j < dadosProposta[i].length; j++) {
      linhaConcatenada.push(dadosProposta[i][j]);
    }
  }

  // Obter a última linha preenchida na aba "Histórico"
  var ultimaLinhaHistorico = abaHistorico.getLastRow();

  // Determinar a próxima linha disponível na aba "Histórico"
  var linhaInicioHistorico = ultimaLinhaHistorico + 0;

  // Determinar a coluna inicial e número de colunas usadas
  var colunaInicial = 49; // Coluna AQ corresponde à coluna 43
  var numeroDeColunas = linhaConcatenada.length;

  // Inserir a linha concatenada na aba "Histórico" a partir da próxima linha disponível e coluna AQ
  var intervaloDestino = abaHistorico.getRange(linhaInicioHistorico, colunaInicial, 1, numeroDeColunas);
  intervaloDestino.setValues([linhaConcatenada]);


  mostrarMensagemHTML(numeroProposta, linkPDF);
  //mostrarColunaM();

}

//////////////////////////////////////////////////////////

//abaHistorico.getRange('A' + linhaHistorico).setValue(new Date()); // Data e hora da execução

function inserirDataHoraAtual() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var abaHistorico = spreadsheet.getSheetByName('Histórico');
  var linhaHistorico = abaHistorico.getLastRow() + 0; // Próxima linha disponível

  // Definir o fuso horário para Brasília
  var fusoHorario = "America/Sao_Paulo";

  // Obter a data e hora atuais com o fuso horário correto
  var dataHoraAtual = new Date();
  var dataHoraBrasilia = Utilities.formatDate(dataHoraAtual, fusoHorario, "dd/MM/yyyy HH:mm:ss");

  // Definir a célula onde a data e hora serão exibidas (exemplo: célula A1)
  sheet.getRange('C6:H6').setValue(dataHoraBrasilia); //antigo A12:B12
  abaHistorico.getRange('B' + linhaHistorico).setValue(dataHoraBrasilia);

  Logger.log('Data e Hora em Brasília: ' + dataHoraBrasilia); // Log para depuração
  return dataHoraBrasilia; // Retorna a data e hora para ser usada
}

function mostrarMensagemHTML(numeroProposta, linkPDF) {
  var dataHoraBrasilia = inserirDataHoraAtual(); // Obtém a data e hora atual

  Logger.log('Valor de dataHoraBrasilia: ' + dataHoraBrasilia); // Log para depuração

  var html = `
    <style>
      body {
        font-family: Arial, sans-serif;
        margin: 20px auto;
        padding: 20px;
        max-width: 400px;
        text-align: center;
        border: 1px solid #ccc;
        border-radius: 10px;
        background-color: #f9f9f9;
        box-shadow: 0px 4px 8px rgba(0, 0, 0, 0.1);
      }
      h2 {
        color: #4CAF50;
      }
      p {
        font-size: 16px;
        margin-bottom: 15px;
      }
      .order-number {
        font-size: 18px;
        font-weight: bold;
        color: #333;
      }
      a {
        display: inline-block;
        margin: 10px 0;
        color: #1E90FF;
        text-decoration: none;
        font-weight: bold;
        transition: color 0.3s;
      }
      a:hover {
        text-decoration: underline;
        color: #0056b3;
      }
      .button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 15px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        font-size: 16px;
        transition: background-color 0.3s ease-in-out;
      }
      .button:hover {
        background-color: #45a049;
      }
      .copy-button {
        background-color: #008CBA;
        margin-left: 10px;
      }
      .copy-button:hover {
        background-color: #0077A2;
      }
      .close-button {
        background-color: #f44336;
      }
      .close-button:hover {
        background-color: #d32f2f;
      }
      @media (max-width: 480px) {
        body {
          margin: 10px;
          padding: 15px;
        }
      }
    </style>
    <body>
      <h2>Ordem de Compra Gerada com Sucesso!</h2>
      <p>Número da Ordem de Compra: 
        <span class="order-number" id="orderNumber">${numeroProposta}</span>
        <button class="button copy-button" onclick="copyOrderNumber()">📋 Copiar</button>
      </p>
      <p>
        <a href="${linkPDF}" target="_blank">📄 Clique aqui para visualizar o PDF</a>
      </p>
      <p>Data & Hora: <strong>${dataHoraBrasilia}</strong></p>
      <p>
        <button class="button close-button" onclick="google.script.host.close()">❌ Fechar</button>
      </p>

      <script>
        function copyOrderNumber() {
          var orderNumber = document.getElementById("orderNumber").innerText;
          navigator.clipboard.writeText(orderNumber).then(function() {
            alert("Número da Ordem de Compra copiado!");
          }).catch(function(err) {
            console.error("Erro ao copiar: ", err);
          });
        }
      </script>
    </body>
  `;


  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), 'Ordem De Compra Gerada');
}



/*/ function gerarPDF(abaProposta, numeroProposta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ORDEM DE COMPRA");

  if (!sheet) {
    throw new Error('Aba não encontrada: ' + abaProposta);
  }

  // Captura o nome do nomeFantasia e o email e envia VIA EMAIL COM PDF ANEXADO NO EMAIL
  var nomenomeFantasia = sheet.getRange("D11:H11").getValue();
  var destinatario = sheet.getRange("D15:H15").getValue(); // Email do nomeFantasia

  // Configura as opções de exportação para o PDF
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&portrait=true' +  // Orientação retrato
    '&size=A4' +        // Tamanho do papel A4
    '&fitw=true' +      // Ajustar largura da página
    '&gridlines=false' + // Ocultar linhas de grade
    '&printtitle=false' + // Ocultar títulos
    '&sheetnames=false' + // Ocultar nomes das abas
    '&pagenumbers=false' + // Ocultar números das páginas
    '&top_margin=0.12' + // Margem superior
    '&bottom_margin=0.2' + // Margem inferior
    '&left_margin=0.2' + // Margem esquerda
    '&right_margin=0.2' + // Margem direita
    '&scale=4';         // Escala de ajuste

  try {
    // Faz a requisição para exportar o PDF
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    // Verifica se o código de resposta não é 200 (sucesso)
    if (response.getResponseCode() !== 200) {
      throw new Error('Erro ao gerar PDF: ' + response.getContentText());
    }

    // Obtém o conteúdo do PDF
    var pdfBlob = response.getBlob().setName(`Ordem de Compra_${numeroProposta}.pdf`);

    // Cria o arquivo PDF na pasta desejada
    var folder = DriveApp.getFolderById("1ldG-RtOThn00-5JUJu_ubHkWWxMg28lH"); // ID DA PASTA DRIVE ---> 3 - Ordens de Compra - OC
    var arquivoPDF = folder.createFile(pdfBlob);

    // HTML personalizado para o corpo do email
    var htmlBody = `
 <html>
  <body style="
    font-family: Arial, sans-serif; 
    color: #333; 
    background-color: #f4f4f4;
    display: flex;
    justify-content: center;
    align-items: center;
    min-height: 100vh;
    margin: 0;
  ">
    <div style="
      padding: 20px; 
      background-color: #fff; 
      border: 1px solid #ddd; 
      border-radius: 5px;
      width: 80%;
      max-width: 600px;
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
    ">
      <div style="
        display: flex; 
        align-items: center; 
        justify-content: space-between;
        margin-bottom: 20px;
      ">
        <h2 style="color: #009688; margin: 0;">Ordem de Compra ${numeroProposta}</h2>
        <img src="https://i.ibb.co/ZgQ6xTM/Equipa-BR-LOGO-TRANSLUCIDO.png" 
             alt="Logo" 
             style="
               width: 150px; 
               height: auto; 
               border-radius: 5px; 
               box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
               object-fit: contain;
             ">
      </div>
      <p>Olá, <strong>${nomenomeFantasia}</strong></p></p>
      <p>Segue em anexo o PDF da Ordem de Compra Nº <strong>${numeroProposta}</strong>.</p>
      <p>Atenciosamente,<br><strong>EQUIPABR</strong></p>
    </div>
  </body>
</html>

    `;

    // Envia o PDF por e-mail com HTML
    MailApp.sendEmail({
      to: destinatario,
      subject: 'Ordem de Compra PDF ' + numeroProposta,
      htmlBody: htmlBody,
      attachments: [pdfBlob]
    });

    return arquivoPDF.getUrl(); // Retorna o link do arquivo PDF
  } catch (e) {
    Logger.log('Erro: ' + e.message); // Loga a mensagem de erro no console
    throw e;
  }
}


*/


function formatarData(data) {
  if (data instanceof Date) {
    return Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return '';
}

function gerarPDFEQRCode(abaProposta, numeroProposta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(abaProposta);

  if (!sheet) {
    throw new Error('Aba não encontrada: ' + abaProposta);
  }

  // Configura as opções de exportação para o PDF
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&portrait=true' +
    '&size=A4' +
    '&fitw=true' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenumbers=false' +
    '&horizontal_alignment=CENTER' +
    '&vertical_alignment=TOP';

  // Faz a requisição para exportar o PDF
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  // Obtém o conteúdo do PDF
  var pdfBlob = response.getBlob().setName(`Proposta_${numeroProposta}.pdf`);

  // Cria o arquivo PDF na pasta desejada
  var folder = DriveApp.getFolderById("10ZbsX--0llWDx4grgBANet1xzqAEICNF");
  var arquivoPDF = folder.createFile(pdfBlob);

  // Gera o link do PDF
  var linkPDF = arquivoPDF.getUrl();

  // Verifica se o linkPDF está correto
  if (!linkPDF) {
    throw new Error('Erro ao gerar link do PDF.');
  }

  // Gera o QR Code
  var qrCodeUrl = gerarQRCode(linkPDF);

  // Gera o QR Code usando a API
  var qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob().setName(`QRCode_Proposta_${numeroProposta}.png`);
  folder.createFile(qrCodeBlob);

  return { pdfUrl: linkPDF, qrCodeUrl: qrCodeUrl };
}

function gerarQRCode(link) {
  // Encoda o link para garantir que não há caracteres especiais
  var encodedLink = encodeURIComponent(link);
  var qrCodeApiUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodedLink}`;
  return qrCodeApiUrl;
}