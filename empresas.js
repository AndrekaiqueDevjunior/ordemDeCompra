function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAtiva = ss.getActiveSheet();
  const linhaEditada = e.range.getRow();
  const colunaEditada = e.range.getColumn();

  // Verifica se a edição foi na aba 'Controle Pedidos'
  if (abaAtiva.getName() === 'Controle Pedidos') {
    Logger.log("Edição ocorreu na aba 'Controle Pedidos'.");

    // Verifica se a edição foi na coluna EO (145ª coluna)
    if (colunaEditada === 145) {
      Logger.log("Edição ocorreu na coluna EO.");

      // Verifica se o valor na célula é 'Gerar OC'
      if (e.value === 'Gerar OC') {
        Logger.log("Valor 'Gerar OC' selecionado na linha: " + linhaEditada);
        const dadosFaltantes = verificarDadosFaltantes(linhaEditada);
        if (dadosFaltantes.length > 0) {
          // Monta mensagem detalhada para a coluna ER
          const mensagemER = '❌ Dados faltantes para gerar OC:\n' +
            dadosFaltantes.map(item => `• ${item}`).join('\n') +
            '\n\nPreencha os dados e tente novamente.';

          // Preenche a coluna ER com os dados faltantes
          const warningCell = abaAtiva.getRange(linhaEditada, 148);
          warningCell.setValue(mensagemER);

          // Formatação para melhor visualização
          warningCell.setWrap(true);
          warningCell.setBackground('#FFF2CC'); // Amarelo claro
          warningCell.setFontColor('#D50000'); // Vermelho

          // Limpa a seleção de "Gerar OC"
          e.range.setValue('');

          Logger.log("Dados faltantes detectados: " + dadosFaltantes.join(', '));
          return; // Interrompe a execução
        }

        // Se não houver dados faltantes, continua o processo normal
        Logger.log("Todos os dados necessários estão presentes. Prosseguindo com a geração da OC.");


        // Preenche o aviso na coluna ER (148ª coluna)
        const warningCell = abaAtiva.getRange(linhaEditada, 148);
        warningCell.setValue('Aviso: Revisão necessária antes de gerar OC');
        const dataHoraBrasilia = inserirDataHoraAtual();
        abaAtiva.getRange(linhaEditada, 146).setValue(dataHoraBrasilia);
        // Gera o número da OC
        const numeroOC = gerarNumeroOC();
        Logger.log("Número da OC gerado: " + numeroOC);

        // Preenche o número da OC na coluna EQ (147ª coluna)
        abaAtiva.getRange(linhaEditada, 147).setValue(numeroOC);
        Logger.log("Número da OC registrado na coluna EQ: " + numeroOC);

        // Aguarda a atualização para garantir a leitura correta
        Utilities.sleep(500);

        // Recupera o número da OC gerado (confirmação)
        const numeroPedido = abaAtiva.getRange(linhaEditada, 147).getValue();
        const numeroPedidoColunaC = abaAtiva.getRange(linhaEditada, 3).getValue();
        const nomeEmpresa = abaAtiva.getRange(linhaEditada, 7).getValue();
        if (!numeroPedido) {
          Logger.log("ERRO: Não foi possível recuperar o número do pedido.");
          return;
        }
        Logger.log("Número do pedido identificado: " + numeroPedido);

        // Acessa a aba 'ORDEM DE COMPRA'
        const abaOrdemCompra = ss.getSheetByName('ORDEM DE COMPRA');
        if (!abaOrdemCompra) {
          Logger.log("ERRO: Aba 'ORDEM DE COMPRA' não encontrada.");
          return;
        }

        // Preenche a célula AO1 com o número do pedido da coluna C
        abaOrdemCompra.getRange('AO1').setValue(numeroPedidoColunaC);
        Logger.log("Célula AO1 preenchida com o número do pedido: " + numeroPedidoColunaC);

        // Preenche as células Q1:S1 com o número da OC
        abaOrdemCompra.getRange('Q1:S1').setValue(numeroPedido);
        Logger.log("Célula Q1:S1 preenchida com o número do pedido: " + numeroPedido);

        abaOrdemCompra.getRange('AO3').setValue(nomeEmpresa);
        Logger.log("Célula AO3 preenchida com O NOME DA EMPRESA " + nomeEmpresa);

        // Executa as funções de preenchimento
        preencherDadosEmpresa();
        preencherDadosProdutos(numeroPedido, linhaEditada);

        // Gera o PDF e salva o link na coluna ES (149ª coluna)
        const linkPDF = gerarPDF(numeroPedido, linhaEditada);

        if (linkPDF) {
          abaAtiva.getRange(linhaEditada, 149).setValue(linkPDF);
        } else {
          Logger.log("ERRO: Link do PDF está vazio ou undefined.");
        }
        Logger.log("Link do PDF: " + linkPDF);
        enviarDadosOrdemCompraParaHistorico();

        // Adicione uma coluna para registrar o envio (por exemplo, coluna 150)
        // NOVO BLOCO - Dispara envio do e-mail com base na coluna ET (150)
        if (abaAtiva.getName() === 'Controle Pedidos' && colunaEditada === 150) {
          const valorET = e.value;
          Logger.log("Valor na coluna ET: " + valorET);

          if (valorET === "SIM") { // DESEJA ENVIAR EMAIL ???
            Logger.log("Disparando função de envio de e-mail para o fornecedor");

            const linha = linhaEditada;

            const numeroOC = abaAtiva.getRange(linha, 147).getValue(); // EQ
            const email = abaAtiva.getRange(linha, 148).getValue();     // ER
            const responsavel = abaAtiva.getRange(linha, 6).getValue(); // F (por exemplo)
            const empresa = abaAtiva.getRange(linha, 7).getValue();     // G (por exemplo)
            const nomeFantasia = abaAtiva.getRange(linha, 8).getValue(); // H (por exemplo)
            const linkPDF = abaAtiva.getRange(linha, 149).getValue();   // ES

            if (!email || !linkPDF) {
              Logger.log("E-mail ou PDF ausente. Envio cancelado.");
              abaAtiva.getRange(linha, 150).setValue("Erro: Dados ausentes");
              return;
            }

            // Chamada da função de envio de e-mail
            enviarEmailParaVendedor(email, numeroOC, linkPDF, responsavel, empresa, nomeFantasia);

            abaAtiva.getRange(linha, 150).setValue("Enviado com sucesso ✔️");
          }
        }

        abaAtiva.getRange(linhaEditada, 145).setValue("OC Gerada");
        //enviarPdfParaDriveiexportarParaHistorico(numeroPedido, linkPDF);

        // Preenche a data/hora de geração do OC na coluna EP (146ª coluna)
      } else {
        Logger.log("Valor selecionado na coluna EO não é 'Gerar OC'.");
      }
    }
  }
      if (abaAtiva.getName() === "Controle Pedidos" && colunaEditada === 120) {
  const cnpj = e.range.getValue();
  if (!cnpj || cnpj.toString().length < 11) return;

  const linhaEditada = e.range.getRow();
  const pedido = abaAtiva.getRange(linhaEditada, 3).getValue(); // Coluna C (nº do pedido)

  Logger.log("Consultando CNPJ na API diretamente: " + cnpj);
  const dados = buscarCNPJ(cnpj);

  if (dados[0].startsWith("Erro") || dados[0] === "CNPJ inválido") {
    Logger.log("Erro ao buscar CNPJ: " + dados[0]);
    return;
  }

  const razaoSocial = dados[0] || dados[1];
  const nomeFantasia = dados[1] || dados[0];
  const email = dados[4];
  const telefone = dados[3];

  const todasLinhas = abaAtiva.getDataRange().getValues();
  for (let i = 1; i < todasLinhas.length; i++) { // pula o cabeçalho
    const linhaCNPJ = String(todasLinhas[i][119]).replace(/[^\d]/g, ''); // Coluna DP (120) - índice 119
    const linhaPedido = todasLinhas[i][2]; // Coluna C - índice 2

    if (linhaCNPJ === String(cnpj).replace(/[^\d]/g, '') && linhaPedido === pedido) {
      abaAtiva.getRange(i + 1, 121).setValue(nomeFantasia); // Coluna DQ
      abaAtiva.getRange(i + 1, 122).setValue(email);        // Coluna DR
      abaAtiva.getRange(i + 1, 124).setValue(telefone);     // Coluna DT
    }
  }
}




  // Verifica se a edição foi na aba 'ORDEM DE COMPRA' na célula AO3
  if (abaAtiva.getName() === 'ORDEM DE COMPRA' && colunaEditada === 41 && linhaEditada === 3) {
    Logger.log("Edição ocorreu na célula AO3 na aba 'ORDEM DE COMPRA'.");

    // Chama a função para preencher os dados da empresa
    preencherDadosEmpresa();
  }
}

function gerarPDF(numeroPedido, linhaControlePedidos) {
  Logger.log("🔄 Iniciando geração de PDF para o pedido: " + numeroPedido);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ORDEM DE COMPRA');
  var abaControle = ss.getSheetByName('Controle Pedidos');

  if (!sheet) throw new Error('❌ Aba "ORDEM DE COMPRA" não encontrada.');

  // Sua função personalizada
  recolhercolunaS();
  // Dados necessários
  var nomeFantasia = sheet.getRange("D11:H11").getValue();
  var destinatario = sheet.getRange("D15").getValue();
  var emailVendedor = sheet.getRange("D16:H16").getValue();
  var responsavel = sheet.getRange("C56:E56").getValue();
  var empresa = sheet.getRange("AO3").getValue();
  var numeroOC = sheet.getRange("Q1").getValue();

  // Exportar como PDF
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
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log("❌ Erro ao baixar PDF: " + response.getContentText());
      throw new Error('Falha ao baixar o PDF');
    }

    var pdfBlob = response.getBlob().setName(`Ordem_de_Compra_${numeroPedido}.pdf`);
    var folder = DriveApp.getFolderById("1ldG-RtOThn00-5JUJu_ubHkWWxMg28lH");
    var arquivoPDF = folder.createFile(pdfBlob);
    var linkPDF = arquivoPDF.getUrl();
    Logger.log("✅ PDF criado com sucesso: " + linkPDF);

    // Gravar o link na coluna ES (149)
    //var linhaAtiva = abaControle.getActiveRange().getRow();
    var linhaAtiva = linhaControlePedidos; // Aqui agora vai funcionar
    abaControle.getRange(linhaAtiva, 149).setValue(linkPDF);



    Logger.log("📌 Link do PDF salvo na coluna ES da linha " + linhaAtiva);

    var dataHoraBrasilia = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");

    // Exibir diálogo de envio ao fornecedor
    if (destinatario && destinatario.toString().includes('@')) {
      var ui = SpreadsheetApp.getUi();
      var resposta = ui.alert(
        'Enviar OC ao Fornecedor?',
        `Deseja enviar a Ordem de Compra ${numeroOC} para o fornecedor ${nomeFantasia}?\n\nEmail: ${destinatario}`,
        ui.ButtonSet.YES_NO
      );


      if (resposta === ui.Button.YES) {
        enviarEmailParaFornecedor(destinatario, numeroOC, linkPDF, responsavel, empresa);
      }
    } else {
      Logger.log("⚠️ E-mail do fornecedor ausente ou inválido: " + destinatario);
    }

    // Enviar ao vendedor (interno)
    if (emailVendedor && emailVendedor.toString().includes('@')) {
      enviarEmailParaVendedor(emailVendedor, numeroOC, linkPDF, responsavel, empresa, nomeFantasia);
      Logger.log("📤 E-mail enviado para o vendedor: " + emailVendedor);
    }

    // HTML de feedback
    var html = `
  <style>
    body { font-family: Arial, sans-serif; background: #f9f9f9; padding: 20px; border-radius: 10px; }
    h2 { color: #4CAF50; }
    a { color: #1E90FF; text-decoration: none; }
    .button { background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; }
  </style>
  <body>
    <h2>Ordem de Compra Gerada com Sucesso!</h2>
    <p><strong>Número:</strong> ${numeroPedido}</p>
    <p><a href="${linkPDF}" target="_blank">Clique aqui para visualizar o PDF</a></p>
    ${destinatario && destinatario.toString().includes('@') ? `<p><strong>Fornecedor:</strong> ${nomeFantasia} (${destinatario})</p>` : ''}
    ${emailVendedor && emailVendedor.toString().includes('@') ? `<p><strong>Vendedor interno:</strong> ${emailVendedor}</p>` : ''}
    <p><em>Gerado em: ${dataHoraBrasilia}</em></p>
    <button class="button" onclick="google.script.host.close()">Fechar</button>
  </body>
`;

    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), 'PDF Gerado');

    liberarGrupo(); // Sua função personalizada
    return linkPDF;

  } catch (e) {
    Logger.log("❌ Erro ao gerar PDF: " + e.message);
    SpreadsheetApp.getUi().alert("Erro ao gerar PDF: " + e.message);
    throw e;
  }
}

function preencherDadosProdutos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaControlePedidos = planilha.getSheetByName('Controle Pedidos');
  var abaOrdemCompra = planilha.getSheetByName('ORDEM DE COMPRA');

  if (!abaControlePedidos || !abaOrdemCompra) {
    Logger.log("Erro: Aba não encontrada.");
    return;
  }

  // Limpa os intervalos de uma vez
  abaOrdemCompra.getRange("D11:H16").clearContent();
  abaOrdemCompra.getRange("J11:S16").clearContent();
  abaOrdemCompra.getRange("B23:S28").clearContent(); // Limpa todas as linhas de produtos


  var numeroPedido = abaOrdemCompra.getRange("AO1").getValue();
  if (!numeroPedido) {
    SpreadsheetApp.getUi().alert('Por favor, selecione um pedido no menu suspenso.');
    return;
  }

  // Lê todos os dados da aba "Controle Pedidos" de uma vez
  var dadosPedidos = abaControlePedidos.getDataRange().getValues();
  var pedidosEncontrados = [];

  // Busca todas as linhas correspondentes ao número do pedido
  for (var i = 1; i < dadosPedidos.length; i++) {
    if (dadosPedidos[i][2].toString().trim() === numeroPedido.toString().trim()) {
      pedidosEncontrados.push(dadosPedidos[i]);
    }
  }

  if (pedidosEncontrados.length === 0) {
    SpreadsheetApp.getUi().alert('Pedido não encontrado na aba "Controle Pedidos".');
    return;
  }

  // Agrupa os produtos pelo CNPJ
  var produtosPorCNPJ = {};
  pedidosEncontrados.forEach(function (pedido) {
    var cnpj = pedido[119]; // CNPJ está na coluna 119
    if (!produtosPorCNPJ[cnpj]) {
      produtosPorCNPJ[cnpj] = [];
    }
    produtosPorCNPJ[cnpj].push(pedido);
  });

  // Itera sobre cada CNPJ e preenche os dados
  for (var cnpj in produtosPorCNPJ) {
    if (produtosPorCNPJ.hasOwnProperty(cnpj)) {
      var produtos = produtosPorCNPJ[cnpj];
      var primeiroProduto = produtos[0];

      // Preenche os dados gerais (apenas uma vez por CNPJ)
      var mapeamentoDadosGerais = [
        { range: "Q1:S1", value: primeiroProduto[146] },  // GERAR OC 
        { range: "AO3", value: primeiroProduto[6] },
        { range: "B23", value: primeiroProduto[20] },     // Quantidade
        { range: "C23", value: primeiroProduto[18] },     // Cód. Forn
        { range: "D23:E23", value: primeiroProduto[30] }, // Unidade
        { range: "F23:I23", value: primeiroProduto[17] }, // Descrição
        { range: "J23", value: primeiroProduto[126] },   // Valor Unitário
        { range: "K23", value: primeiroProduto[128] },   // Desconto
        { range: "L23", value: primeiroProduto[129] },   // IPI
        { range: "M23:Q23", value: primeiroProduto[130] }, // ICMS ST
        //{ range: "S23", value: primeiroProduto[132] },   // VALOR TOTAL 1 
        //{ range: "S29", value: primeiroProduto[132] },   // VALOR TOTAL 2
        //{ range: "M35:S35", value: primeiroProduto[132] }, // VALOR TOTAL 3 
        { range: "H35:L35", value: primeiroProduto[131] }, // Valor Frete
        { range: "B35:G35", value: primeiroProduto[136] }, // Frete por Conta					***
        { range: "D16:H16", value: primeiroProduto[121] }, // Email fornecedor
        { range: "J16:S16", value: primeiroProduto[122] }, // Contato fornecedor
        { range: "C56:E56", value: primeiroProduto[118] }, // Comprador
        { range: "C46", value: primeiroProduto[0] },       // Referência Interna
        { range: "C47", value: primeiroProduto[133] },     // Condições de Pagamento
        { range: "D12:H12", value: primeiroProduto[119] } // CNPJ

      ];

      mapeamentoDadosGerais.forEach(function (dado) {
        abaOrdemCompra.getRange(dado.range).setValue(dado.value);
      });





      // Preenche os dados dos produtos (um por linha)
      var linhaInicial = 23; // Linha inicial para preencher os dados dos produtos
      produtos.forEach(function (produto, index) {
        var linhaAtual = linhaInicial + index;

        // Concatena a descrição do produto com o ICMS Substituição Tributária
        var descricaoCompleta = produto[17] + "\nICMS Substituição Tributária: 0 % - 0.00 BRL";

        var mapeamentoDadosProduto = [
          { range: "B" + linhaAtual, value: produto[20] },     // Quantidade
          { range: "C" + linhaAtual, value: produto[18] },     // Cód. Forn
          { range: "D" + linhaAtual + ":E" + linhaAtual, value: produto[30] }, // Unidade
          { range: "F" + linhaAtual + ":I" + linhaAtual, value: descricaoCompleta }, // Descrição
          { range: "J" + linhaAtual, value: produto[126] },   // Valor Unitário
          { range: "K" + linhaAtual, value: produto[128] },   // Desconto
          { range: "L" + linhaAtual, value: produto[129] },   // IPI
          { range: "M" + linhaAtual + ":Q" + linhaAtual, value: produto[130] }, // ICMS ST
          { range: "R" + linhaAtual, value: produto[135] },   // Previsão de Entrega
          { range: "S" + linhaAtual, value: produto[132] }    // Valor Total
        ];

        mapeamentoDadosProduto.forEach(function (dado) {
          abaOrdemCompra.getRange(dado.range).setValue(dado.value);
        });
      });

      // Calcula a data de entrega
      var dataGeracaoOC = abaOrdemCompra.getRange("C7").getValue();
      var tempoEntrega = primeiroProduto[135];

      if (dataGeracaoOC && tempoEntrega) {
        var dataPrevisaoEntrega = new Date(dataGeracaoOC);
        dataPrevisaoEntrega.setDate(dataPrevisaoEntrega.getDate() + 1 + tempoEntrega);

        abaOrdemCompra.getRange("R23:R28").setValue(dataPrevisaoEntrega);
        abaOrdemCompra.getRange("J7").setValue(dataPrevisaoEntrega);
        Logger.log("Data prevista de entrega calculada e preenchida em R23: " + dataPrevisaoEntrega);
      } else {
        Logger.log("Erro: Dados insuficientes para calcular a data de entrega.");
      }

      // Busca o CNPJ do fornecedor
      if (cnpj) {
        Logger.log("Buscando CNPJ: " + cnpj);
          preencherDadosCNPJ(cnpj);
      } else {
        Logger.log("CNPJ não encontrado para o pedido.");
      }
    }
  }

  Logger.log("Dados do pedido preenchidos com sucesso!");
}

function preencherDadosEmpresa(e) {
  var ss = e ? e.source : SpreadsheetApp.getActiveSpreadsheet();
  var abaEmpresas = ss.getSheetByName('Empresas');
  var abaOrdemCompra = ss.getSheetByName('ORDEM DE COMPRA');
  var abaControlePedidos = ss.getSheetByName('Controle Pedidos');

  // Obter a empresa selecionada no menu suspenso
  var nomeEmpresa = abaOrdemCompra.getRange("AO3").getValue();

  // Verificar se a célula AO3 está preenchida
  if (!nomeEmpresa) {
    Logger.log("AO3 não está preenchido. Aguardando dados...");
    return; // Não usamos alert() em gatilhos
  }

  // Buscar todas as linhas da aba "Empresas"
  var dadosEmpresas = abaEmpresas.getDataRange().getValues();
  var empresaEncontrada = null;

  // Procurar a empresa na coluna A (índice 0)
  for (var i = 1; i < dadosEmpresas.length; i++) {
    if (dadosEmpresas[i][0].toString().trim().toLowerCase() === nomeEmpresa.toString().trim().toLowerCase()) {
      empresaEncontrada = dadosEmpresas[i];
      break;
    }
  }

  if (!empresaEncontrada) {
    Logger.log('Empresa não encontrada: ' + nomeEmpresa);
    return;
  }

  // Mapeamento das colunas da aba "Empresas" para "ORDEM DE COMPRA"
  var mapeamento = {
    "Razão Social": { coluna: 1, destino: ["AD10:AM10", "AD10:AM10"] },
    "Nome Fantasia": { coluna: 2, destino: ["X10:AB10"] },
    "CNPJ": { coluna: 3, destino: ["X11:AB11"] },
    "Rua/ENDERECO": { coluna: 6, destino: ["E3:J3"] },
    "Bairro": { coluna: 9, destino: ["R3:S3"] },
    "Cidade": { coluna: 10, destino: ["R4:S4"] },
    "Estado": { coluna: 11, destino: ["AD13:AM13"] },
    "CEP": { coluna: 12, destino: ["E4:J4"] },
    "Telefone": { coluna: 13, destino: ["AD14:AM14"] },
    "Email": { coluna: 17, destino: ["X14:AB14"] },
    "nfe XML:": { coluna: 18, destino: ["C49"] }
  };

  // Preencher os dados na aba "ORDEM DE COMPRA"
  Object.keys(mapeamento).forEach(function (chave) {
    var colunaIndex = mapeamento[chave].coluna;
    var intervalos = mapeamento[chave].destino;
    var valor = empresaEncontrada[colunaIndex];

    intervalos.forEach(function (intervalo) {
      abaOrdemCompra.getRange(intervalo).setValue(valor);
    });
  });

  // Concatenar Nome Fantasia, Razão Social e CNPJ para a célula Q1
  var nomeFantasia = empresaEncontrada[mapeamento["Nome Fantasia"].coluna] || "";
  var razaoSocial = empresaEncontrada[mapeamento["Razão Social"].coluna] || "";
  var cnpj = empresaEncontrada[mapeamento["CNPJ"].coluna] || "";
  var email = empresaEncontrada[mapeamento["Email"].coluna] || "";
  var telefone = empresaEncontrada[mapeamento["Telefone"].coluna] || "";

  var partesLinha1 = [];
  if (nomeFantasia) partesLinha1.push(nomeFantasia);
  if (razaoSocial) partesLinha1.push(razaoSocial);
  if (cnpj) partesLinha1.push("CNPJ: " + cnpj);

  var linha1 = partesLinha1.join(", ");
  var linha2 = telefone ? "Telefone:  " + telefone : "";
  var linha3 = email ? "Email:  " + email : "";

  var textoFormatado = linha1 + "\n" + linha2 + "\n" + linha3;

  abaOrdemCompra.getRange("D1:N1").setValue(textoFormatado);

  // Verificar o número do pedido na coluna G da aba "Controle de Pedidos"
  var numeroPedido = abaControlePedidos.getRange("G2").getValue(); // Aqui você pode ajustar a linha onde está o número do pedido

  var textoEmpresa = "";

  if (numeroPedido === "avsk") {
    textoEmpresa = "AVSK SOLUTIONS, AVSK Comércio, Importação e Exportação LTDA, CNPJ: 12297813/0001-74\nTelefone: (11) 2808-0202\nEmail: contato@avsk.com.br";
  } else if (numeroPedido === "ay") {
    textoEmpresa = "EQUIPABR, AY MANUTENÇÃO INDUSTRIAL, ENGENHARIA E EQUIPAMENTOS EIRELI – EPP, CNPJ: 2304543/00001-25\nTelefone: (11) 2630-1200";
  } else if (numeroPedido === "dns") {
    textoEmpresa = "ENG DNS, DNS – FUTURE & COMERCIO LTDA - ME, CNPJ: 31885811/0001-40\nTelefone: (11) 2321-1497\nEmail: contato@engdns.com.br";
  } else if (numeroPedido === "dantools") {
    textoEmpresa = "DANTOOLS, DANTOOLS FERRAMENTAS EIRELI - EPP, CNPJ: 29077923/0001-23\nTelefone: (11) 4375-8000\nEmail: contato@dantools.com.br";
  }

  // Preencher a célula D1:N1 com as informações da empresa correspondente
  if (textoEmpresa) {
    abaOrdemCompra.getRange("D1:N1").setValue(textoEmpresa);
  }
}
/*
function preencherDadosCNPJ(cnpj) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = planilha.getSheetByName("ORDEM DE COMPRA"); // Seleciona a aba específica "Proposta 2024"

  if (!sheet) {
    return 'Erro: Aba "Proposta 2024" não encontrada.'; // Verifica se a aba existe
  }

  // Limpar os dados das células que serão preenchidas
  sheet.getRange("D11:H14").clearContent();
  sheet.getRange("J11:S15").clearContent();

  // Chama a função que busca os dados do CNPJ
  var dados = buscarCNPJ(cnpj);

  // Verifica se houve erro na busca de dados
  if (dados[0].startsWith('Erro') || dados[0] === 'CNPJ inválido') {
    return dados[0];  // Retorna a mensagem de erro para o usuário
  }

  // Preenche as células específicas na aba "Proposta 2024" com validação
  if (dados[0]) {
    sheet.getRange("J11:S11").merge();
    sheet.getRange("J11").setValue(dados[0]); // Nome da empresa
  }

  // Se houver Nome Fantasia, usa ele; senão, usa a Razão Social
  var nomeFantasiaOuRazaoSocial = dados[1] ? dados[1] : dados[0];

  if (nomeFantasiaOuRazaoSocial) {
    sheet.getRange("D11").merge();
    sheet.getRange("D11:H11").setValue(nomeFantasiaOuRazaoSocial);
  }


  if (dados[2]) {
    sheet.getRange("J14:S14").setValue(dados[2]); // UF
  }

  if (dados[3]) {
    sheet.getRange("J15:S15").setValue(dados[3]); // Telefone
  }

  if (dados[4]) {
    sheet.getRange("D13:H13").merge();
    sheet.getRange("D13").setValue(dados[4]); // Email
  }

  if (dados[7]) {
    sheet.getRange("J12:S12").setValue(dados[7]); // Logradouro
  }

  //if (dados[8]) {
  //  sheet.getRange("").setValue(dados[8]); // Número
  //}

  //if (dados[9]) {
  //sheet.getRange("F7").setValue(dados[9]); // Complemento do logradouro
  //}

  if (dados[10]) {
    sheet.getRange("D13:H13").setValue(dados[10]); // Bairro
  }

  if (dados[11]) {
    sheet.getRange("D14:H14").setValue(dados[11]); // Município // cidade
  }

  if (dados[13]) {
    sheet.getRange("D12:H12").merge();
    sheet.getRange("D12").setValue(dados[13]); // CNPJ
  }
  if (dados[14]) {
    sheet.getRange("J13:S13").merge();
    sheet.getRange("J13").setValue(dados[14]); // CEP
  }



  return dados;  // Retorna os dados para exibição no HTML, se necessário
}
*/
function gerarNumeroOC() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ultimoNumeroOC';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaControlePedidos = ss.getSheetByName('Controle Pedidos');

  const valoresOC = abaControlePedidos.getRange(2, 147, abaControlePedidos.getLastRow() - 1).getValues();

  let ultimaSequencia = 0;
  valoresOC.forEach(row => {
    const match = row[0] ? String(row[0]).match(/(\d{6})(\d{4})-0$/) : null;
    if (match) {
      const sequencia = parseInt(match[2], 10);
      if (sequencia > ultimaSequencia) {
        ultimaSequencia = sequencia;
      }
    }
  });

  const novaSequencia = (ultimaSequencia + 1).toString().padStart(4, '0');
  const hoje = new Date();
  const ano = hoje.getFullYear().toString().slice(-2);
  const mes = (hoje.getMonth() + 1).toString().padStart(2, '0');
  const dia = hoje.getDate().toString().padStart(2, '0');
  const numeroOC = `${ano}${mes}${dia}${novaSequencia}-0`;

  cache.put(cacheKey, novaSequencia, 21600);

  const abaProposta = ss.getSheetByName('ORDEM DE COMPRA');
  const dataHoraBrasilia = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
  abaProposta.getRange('C7').setValue(dataHoraBrasilia);

  return numeroOC;
}

function enviarEmailParaVendedor(emailVendedor, numeroOC, linkPDF, responsavel, empresa, nomeFantasia) { //localizada na celula D16:H16 E-mail Vendedor:	
  var dataHora = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");

  // Gerar URL para o QR Code usando a API do Google Charts
  var qrCodeUrl = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" + encodeURIComponent(linkPDF);
  // HTML para o e-mail com QR Code
  var htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Nova Ordem de Compra - ${numeroOC}</title>
      <style>
        body {
          font-family: 'Arial', sans-serif;
          line-height: 1.6;
          color: #333;
          max-width: 600px;
          margin: 0 auto;
          padding: 20px;
          background-color: #f9f9f9;
        }
        .container {
          background-color: #ffffff;
          border-radius: 8px;
          padding: 25px;
          box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
          position: relative;
        }
        .header {
          text-align: center;
          padding-bottom: 20px;
          border-bottom: 1px solid #eeeeee;
          margin-bottom: 20px;
        }
        .header h1 {
          color: #2c3e50;
          margin: 0;
          font-size: 24px;
        }
        .content {
          margin-bottom: 25px;
        }
        .content p {
          margin-bottom: 15px;
        }
        .highlight {
          background-color: #f8f9fa;
          padding: 15px;
          border-radius: 5px;
          border-left: 4px solid #3498db;
        }
        .button {
          display: inline-block;
          background-color: #3498db;
          color: #ffffff !important;
          text-decoration: none;
          padding: 12px 25px;
          border-radius: 5px;
          font-weight: bold;
          margin: 15px 0;
        }
        .footer {
          text-align: center;
          font-size: 12px;
          color: #7f8c8d;
          padding-top: 20px;
          border-top: 1px solid #eeeeee;
          position: relative;
        }
        .details {
          margin: 20px 0;
        }
        .details-item {
          margin-bottom: 8px;
        }
        .details-label {
          font-weight: bold;
          color: #2c3e50;
        }
        .qr-code {
          position: absolute;
          bottom: 20px;
          right: 20px;
          text-align: center;
        }
        .qr-code img {
          width: 80px;
          height: 80px;
          border: 1px solid #eee;
          padding: 5px;
          background: white;
        }
        .qr-code p {
          font-size: 10px;
          margin-top: 5px;
          color: #7f8c8d;
        }
        @media only screen and (max-width: 480px) {
          .qr-code {
            position: static;
            text-align: center;
            margin-top: 20px;
          }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>Nova Ordem de Compra</h1>
        </div>
        
        <div class="content">
          <p>Prezado(a) Vendedor,</p>
          
          <p>Informamos que foi gerada uma nova Ordem de Compra para sua empresa.</p>
          
          <div class="highlight">
            <p><strong>Por favor, verifique os detalhes abaixo:</strong></p>
          </div>
          
          <div class="details">
            <div class="details-item">
              <span class="details-label">Número da OC:</span> ${numeroOC}
            </div>
            <div class="details-item">
              <span class="details-label">Empresa Solicitante:</span> ${empresa}
            </div>
            <div class="details-item">
              <span class="details-label">Cliente:</span> ${nomeFantasia}
            </div>
            <div class="details-item">
              <span class="details-label">Responsável:</span> ${responsavel}
            </div>
            <div class="details-item">
              <span class="details-label">Data e Hora:</span> ${dataHora}
            </div>
          </div>
          
          <center>
            <a href="${linkPDF}" class="button" target="_blank">Visualizar Ordem de Compra</a>
          </center>
          
          <p>Agradecemos pela atenção e ficamos à disposição para qualquer dúvida.</p>
          
          <p>Atenciosamente,<br>Equipe de Compras<br>${empresa}</p>
        </div>
        
        <div class="footer">
          <div class="qr-code">
            <img src="${qrCodeUrl}" alt="QR Code com link para a Ordem de Compra">
            <p>Scan para acessar</p>
          </div>
          <p>Este é um e-mail automático, por favor não responda diretamente a esta mensagem.</p>
          <p>© ${new Date().getFullYear()} ${empresa}. Todos os direitos reservados.</p>
        </div>
      </div>
    </body>
    </html>
  `;

  var assunto = `Nova Ordem de Compra ${numeroOC} - ${empresa}`;
  var pdfBlob;
  try {
    var pdfResponse = UrlFetchApp.fetch(linkPDF, {
      muteHttpExceptions: true
    });

    if (pdfResponse.getResponseCode() === 200) {
      pdfBlob = pdfResponse.getBlob().setName(`OC_${numeroOC}.pdf`);
    } else {
      throw new Error("Falha ao baixar o PDF");
    }
  } catch (e) {
    Logger.log("Erro ao baixar PDF: " + e.message);
  }

  // Opções do e-mail
  var options = {
    htmlBody: htmlBody,
    name: empresa,
    noReply: true,
    attachments: pdfBlob ? [pdfBlob] : []  // anexa o PDF se estiver disponível
  };

  // Se temos o QR Code, adicionamos como imagem inline
  if (qrCodeBlob) {
    options.inlineImages = {
      qrCode: qrCodeBlob
    };
  }

  MailApp.sendEmail(emailVendedor, assunto, "", options);

  try {
    // Primeiro tentamos obter a imagem do QR Code
    var qrCodeBlob;

    try {
      var qrCodeResponse = UrlFetchApp.fetch(qrCodeUrl, {
        muteHttpExceptions: true
      });

      if (qrCodeResponse.getResponseCode() === 200) {
        qrCodeBlob = qrCodeResponse.getBlob();
      } else {
        throw new Error("Falha ao gerar QR Code");
      }
    } catch (e) {
      Logger.log("Não foi possível gerar QR Code: " + e.message);
      // Se falhar, continuamos sem o QR Code
      htmlBody = htmlBody.replace('<div class="qr-code">.*?</div>', '');
    }

    // Opções do e-mail
    var options = {
      htmlBody: htmlBody,
      name: empresa,
      noReply: true
    };

    // Se temos o QR Code, adicionamos como imagem inline
    if (qrCodeBlob) {
      options.inlineImages = { qrCode: qrCodeBlob };
    }
    if (pdfBlob) {
      options.attachments = [pdfBlob];
    }

    MailApp.sendEmail(emailVendedor, assunto, "", options);
    Logger.log(`E-mail enviado com sucesso para: ${emailVendedor}`);
  } catch (e) {
    Logger.log(`Erro ao enviar e-mail para ${emailVendedor}: ${e.message}`);
    throw new Error(`Falha ao enviar e-mail para o vendedor: ${e.message}`);
  }
}

function testarEnvioEmail() {
  enviarEmailParaVendedor("andre.isola@tothetop.app.br", "TESTE123", "https://exemplo.com", "Fulano", "Empresa Teste", "Cliente Teste");
}


function encontrarEmpresaPorNome(abaEmpresas, nomeEmpresa) {
  var dadosEmpresas = abaEmpresas.getDataRange().getValues();
  for (var i = 1; i < dadosEmpresas.length; i++) {
    if (dadosEmpresas[i][0].toString().trim().toLowerCase() === nomeEmpresa.toString().trim().toLowerCase()) {
      return dadosEmpresas[i];
    }
  }
  return null;
}
// Cache em memória
const cacheCNPJ = {};

function buscarNoBancoDadosLocal(cnpj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBancoEmpresas = ss.getSheetByName('Banco de Dados - Empresas');
  const dados = abaBancoEmpresas.getDataRange().getValues();

  // Encontra a linha com o CNPJ (assumindo que está na coluna B)
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][1] && dados[i][1].replace(/[^\d]+/g, '') === cnpj) {
      return {
        nome: dados[i][0],
        cnpj: dados[i][1],
        razaoSocial: dados[i][2] || dados[i][0],
        email: dados[i][3],
        bairro: dados[i][4],
        cep: dados[i][5],
        logradouro: dados[i][6],
        municipio: dados[i][7],
        uf: dados[i][8],
        telefone: dados[i][9]
      };
    }
  }
  return null;
}

function buscarNaAPIReceitaWS(cnpj) {
  const url = "https://receitaws.com.br/v1/cnpj/" + cnpj;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return JSON.parse(response.getContentText());
}

function salvarNoBancoDadosLocal(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBancoEmpresas = ss.getSheetByName('Banco de Dados - Empresas');

  abaBancoEmpresas.appendRow([
    dados.nome,
    dados.cnpj,
    dados.razaoSocial || dados.nome,
    dados.email,
    dados.bairro,
    dados.cep,
    dados.logradouro,
    dados.municipio,
    dados.uf,
    dados.telefone
  ]);
}







// Exemplo para formatar números com duas casas decimais
function setValueWithFormat(range, value, format) {
  if (typeof value === "number" || !isNaN(parseFloat(value))) {
    range.setNumberFormat(format);  // Define o formato
    range.setValue(value);          // Preenche o valor
  } else {
    range.setValue(value);          // Apenas preenche o valor se não for número
  }
}

function mostrarMensagemHTML(numeroProposta, linkPDF) {
  var dataHoraBrasilia = inserirDataHoraAtual(); // Obtém a data e hora atual

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

function inserirDataHoraAtual(aba = null, celula = null) {
  const fusoHorario = "America/Sao_Paulo";
  const agora = new Date();
  const dataHoraFormatada = Utilities.formatDate(agora, fusoHorario, "dd/MM/yyyy HH:mm:ss");

  // Se aba e célula forem fornecidos, insere a data/hora no local especificado
  if (aba && celula) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(aba);
    sheet.getRange(celula).setValue(dataHoraFormatada);
  }

  Logger.log('Data e Hora em Brasília: ' + dataHoraFormatada);
  return dataHoraFormatada;
}

// Função auxiliar para atualizar uma célula somente se ela não possuir uma fórmula
function atualizarCelulaSeNaoTiverFormula(celula, valor) {
  if (!celula.getFormula()) {
    celula.setValue(valor);
  }
}



function encontrarLinhaPorPedido(sheet, numeroPedido) {
  var valores = sheet.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (valores[i][0] == numeroPedido) {
      return i + 1;
    }
  }
  return null;
}

function encontrarLinhaPorCNPJ(sheet, cnpj) {
  var valores = sheet.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (valores[i][1] == cnpj) { // CNPJ está na segunda coluna
      return i + 1;
    }
  }
  return null;
}

function preencherOrdemDeCompra(sheet, linha, dados) {
  sheet.getRange(linha, 1, 1, dados.length).setValues([dados]);
}

function salvarNoBancoDeDados(sheet, dados) {
  sheet.appendRow(dados);
}


//======================================
