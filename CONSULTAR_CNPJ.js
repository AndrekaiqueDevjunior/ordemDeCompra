
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('FormCNPJ');
}



// Função que abre o formulário HTML
function abrirFormCNPJ() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('FormCNPJ')
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Consulta CNPJ');
}

// Função para buscar os dados de um CNPJ (usada no script anterior)

// Função para buscar dados do CNPJ na célula
function ConsultarCNPJ(cnpj) {
  // Chama a função que busca os dados do CNPJ
  var dados = buscarCNPJ(cnpj);

  // Verifica se houve erro na busca de dados
  if (dados[0].startsWith('Erro') || dados[0] === 'CNPJ inválido') {
    return dados[0];  // Retorna a mensagem de erro
  }

  // Retorna os dados como uma matriz (array) para que sejam exibidos corretamente na planilha
  return [dados];  // Retorna os dados como um array de arrays
}

function buscarCNPJ(cnpj) {
  cnpj = cnpj.replace(/[^\d]+/g, '');

  if (cnpj.length !== 14) return ['CNPJ inválido'];

  var url = 'https://www.receitaws.com.br/v1/cnpj/' + cnpj;
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var content = response.getContentText();

  // Verifica se a resposta parece ser JSON antes de tentar fazer parse
  if (!content.trim().startsWith('{')) {
    Logger.log("Resposta bruta da API (não é JSON):");
    Logger.log(content);
    return ['Erro: Limite de requisições excedido ou resposta inválida'];
  }

  try {
    var result = JSON.parse(content);

    if (result.status !== 'OK') return ['Erro: ' + result.message];

    return [
      result.nome,
      result.fantasia,
      result.uf,
      result.telefone,
      result.email,
      result.atividade_principal[0].text,
      result.situacao,
      result.logradouro,
      result.numero,
      result.complemento,
      result.bairro,
      result.municipio,
      result.capital_social,
      result.cnpj,
      result.cep
    ];
  } catch (erro) {
    Logger.log("Erro no parse da resposta: " + erro);
    Logger.log("Resposta bruta da API:");
    Logger.log(content);
    return ['Erro: resposta inválida da API'];
  }
}



// Função para buscar e preencher os dados na planilha "Proposta 2024"
function preencherDadosCNPJ(cnpj) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = planilha.getSheetByName("ORDEM DE COMPRA");
  var bdEmpresas = planilha.getSheetByName("Banco de Dados - Empresas");

  if (!sheet || !bdEmpresas) return 'Erro: Abas necessárias não encontradas.';

  // Limpa os dados da área antes de preencher
  sheet.getRange("D11:H14").clearContent();
  sheet.getRange("J11:S15").clearContent();

  // Normaliza o CNPJ
  cnpj = cnpj.replace(/[^\d]+/g, '');
  if (cnpj.length !== 14) return 'CNPJ inválido';

  // Verifica se o CNPJ já está no banco de dados
  var bdDados = bdEmpresas.getDataRange().getValues();
  var index = bdDados.findIndex((linha, i) => i > 0 && String(linha[1]).replace(/[^\d]+/g, '') === cnpj);

  var dados = [];

  if (index !== -1) {
    Logger.log("CNPJ encontrado no banco local.");
    var linha = bdDados[index];

    var enderecoSplit = linha[6]?.split(',') || [];
    var complementoSplit = enderecoSplit[1]?.split(' - ') || [];

    dados = [
      linha[2],  // Razão Social
      linha[0],  // Nome Fantasia
      linha[8],  // Estado
      linha[9],  // Telefone
      linha[3],  // E-mail
      '',        // Atividade principal
      '',        // Situação
      enderecoSplit[0] || '',            // Logradouro
      complementoSplit[0]?.trim() || '', // Número
      complementoSplit[1] || '',         // Complemento
      linha[4],  // Bairro
      linha[7],  // Cidade (municipio)
      '',        // Capital Social
      linha[1],  // CNPJ
      linha[5]   // CEP
    ];




  } else {
    Logger.log("CNPJ não encontrado. Consultando API...");
    dados = buscarCNPJ(cnpj);
    if (dados[0].startsWith('Erro') || dados[0] === 'CNPJ inválido') return dados[0];


      // Garante que Nome Fantasia ou Razão Social não fiquem vazios
    if (!dados[0]) dados[0] = dados[1]; // Razão Social
    if (!dados[1]) dados[1] = dados[0]; // Nome Fantasia
    // Salva os dados na aba Banco de Dados - Empresas
    // Salva os dados na aba Banco de Dados - Empresas
    bdEmpresas.appendRow([
      dados[1],  // Nome Fantasia
      dados[13], // CNPJ
      dados[0],  // Razão Social
      dados[4],  // E-mail
      dados[10], // Bairro
      dados[14], // CEP
      `${dados[7]}${dados[8] ? ', ' + dados[8] : ''}${dados[9] ? ' - ' + dados[9] : ''}`, // Endereço
      dados[11], // Cidade
      dados[2],  // Estado
      dados[3]   // Telefone
    ]);



  }

  // Preenche os dados na planilha "ORDEM DE COMPRA"
  // Coluna D: Informações principais
  sheet.getRange("D11:H11").setValue(dados[1]);   // Nome Fantasia
  sheet.getRange("D12:H12").setValue(dados[13]);  // CNPJ
  sheet.getRange("D13:H13").setValue(dados[10]);  // Bairro
  sheet.getRange("D14:H14").setValue(dados[11]);  // Cidade
  sheet.getRange("D15:H15").setValue(dados[4]);   // E-mail

  // Coluna J: Dados adicionais
  sheet.getRange("J11:S11").setValue(dados[0]);   // Razão Social
  sheet.getRange("J12:S12").setValue(
    `${dados[7]}${dados[8] ? ', ' + dados[8] : ''}${dados[9] ? ' - ' + dados[9] : ''}`.trim()
  );                                              // Endereço completo (logradouro, número, complemento)
  sheet.getRange("J13:S13").setValue(dados[14]);  // CEP
  sheet.getRange("J14:S14").setValue(dados[2]);   // Estado
  sheet.getRange("J15:S15").setValue(dados[3]);   // Telefone



  return dados;
}

function testarBuscaCNPJDaCelula() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("ORDEM DE COMPRA");
  if (!aba) {
    Logger.log("Aba 'ORDEM DE COMPRA' não encontrada.");
    return;
  }

  var cnpj = String(aba.getRange("D12").getValue());
  if (!cnpj) {
    Logger.log("CNPJ não encontrado na célula D12.");
    return;
  }

  Logger.log("Buscando CNPJ: " + cnpj);
  var resultado = preencherDadosCNPJ(cnpj);

  Logger.log("Resultado da busca:");
  Logger.log(resultado);
}

