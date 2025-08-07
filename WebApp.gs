// WebApp.gs (Versão Final Corrigida e Melhorada)

// --- Variáveis e Funções Auxiliares ---

var PRINT_OPTIONS = {
  'size': 7, 'fzr': false, 'portrait': true, 'fitw': true,
  'top_margin': 0.8, 'bottom_margin': 0.0, 'left_margin': 0.6, 'right_margin': 0.0,
  'gridlines': false, 'printtitle': false, 'sheetnames': false,
  'pagenum': 'UNDEFINED', 'attachment': false,
};

var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function objectToQueryString(obj) {
  return Object.keys(obj).map(function(key) {
    return Utilities.formatString('&%s=%s', key, obj[key]);
  }).join('');
}

/**
 * NOVA VERSÃO - Salva o PDF diretamente no Google Drive, em vez de copiar a planilha.
 * @param {string} pdfUrl - A URL do PDF a ser salvo.
 * @param {string} nomeArquivo - O nome que o arquivo PDF terá.
 */
function salvarPdfNoDrive(pdfUrl, nomeArquivo) {
  try {
    var pastaDestino = DriveApp.getFolderById("1Wt82Htk8FzRtd4xQTGljV8OSfb_PzK-G"); // ID da sua pasta

    var response = UrlFetchApp.fetch(pdfUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });

    var blob = response.getBlob().setName(nomeArquivo + '.pdf');
    pastaDestino.createFile(blob);
    
  } catch (e) {
    throw new Error("Falha ao salvar PDF no Drive: " + e.message);
  }
}


// --- Funções Principais do Web App ---

function doGet(e) {
  var nomes = getNomesDaBaseDeDados();
  var template = HtmlService.createTemplateFromFile('Index');
  template.nomesParaDropdown = nomes;
  return template.evaluate()
    .setTitle("Sistema Gerador de Memorandos")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function getNomesDaBaseDeDados() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaBaseDados = planilha.getSheetByName("BASE DE DADOS");
  
  if (!guiaBaseDados) {
    return ["ERRO: Aba 'BASE DE DADOS' não encontrada"];
  }

  // Pega todos os valores da coluna B, começando da B3
  var range = guiaBaseDados.getRange("B3:B" + guiaBaseDados.getLastRow());
  
  var nomes = range.getDisplayValues()
                   .map(function(linha) { 
                      // Pega o primeiro (e único) item de cada linha do array
                      return linha[0]; 
                   }) 
                   .filter(function(nome) { 
                      // A MÁGICA ACONTECE AQUI:
                      // Retorna 'true' apenas se 'nome' não for nulo, indefinido, 
                      // e, após remover os espaços, não for uma string vazia.
                      return nome && nome.trim() !== ""; 
                   }); 
                   
  return nomes;
}

/**
 * VERSÃO ATUALIZADA - Corrige a célula do nome e chama a nova função de salvar.
 */
function gerarMemorandoPeloWebApp(dadosDoFormulario) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaMemorando = planilha.getSheetByName("MEMORANDO");

    if (!guiaMemorando) {
      throw new Error("Aba 'MEMORANDO' não encontrada.");
    }
    
    // --- CORREÇÃO APLICADA AQUI ---
    // O nome do interessado agora é preenchido na célula E8.
    guiaMemorando.getRange("E8").setValue(dadosDoFormulario.destinatario);
    guiaMemorando.getRange("D4").setValue("MEMORANDO Nº CPI2 - " + dadosDoFormulario.numeroOficio);
    
    SpreadsheetApp.flush();
    
    // 3. Gera a URL do PDF
    var range = guiaMemorando.getRange("B1:H45");
    var gid = guiaMemorando.getSheetId();
    var printRange = objectToQueryString({
      'c1': range.getColumn() - 1, 'r1': range.getRow() - 1,
      'c2': range.getColumn() + range.getWidth() - 1, 'r2': range.getRow() + range.getHeight() - 1
    });
    var url = planilha.getUrl().replace(/edit$/, '') + 'export?format=pdf' + PDF_OPTS + printRange + "&gid=" + gid;

    // 4. Cria o nome do arquivo e chama a nova função para salvar o PDF no Drive
    var data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH-mm");
    var nomeArquivo = "Memorando " + dadosDoFormulario.destinatario + " " + data;
    salvarPdfNoDrive(url, nomeArquivo);

    // 5. Retorna a URL para a interface do usuário abrir
    return url;

  } catch (e) {
    // Se der qualquer erro, envia uma mensagem clara para o usuário
    return "ERRO: " + e.message;
  }
}

// --- Código da Ponte (Menu na Planilha) - OPCIONAL ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sistema de Memorandos')
      .addItem('Abrir Gerador de Memorandos', 'abrirWebApp')
      .addToUi();
}

function abrirWebApp() {
  var url = "https://script.google.com/macros/s/AKfycby83a2tXAKZeneV2fi4b_BCGf-oBhBvjverkWRzqy7jFQtBZzPWFlwYs8-120yu8rs/exec"; // Lembre-se de colocar sua URL aqui
  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");</script>')
      .setWidth(100).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrindo o sistema...');
}