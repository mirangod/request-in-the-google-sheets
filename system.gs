// GLOBAIS
const SHEET = SpreadsheetApp.getActiveSpreadsheet();
//const UI = SpreadsheetApp.getUi();

// LOGIN E PASSWORD
const ABA_LOGIN = SHEET.getSheetByName("LOGIN");
const ABA_CANCELAMENTO = SHEET.getSheetByName("CANCELAMENTO");
const ABA_DEVOLUÇÃO    = SHEET.getSheetByName("DEVOLUÇÃO");
const lgn = ABA_LOGIN.getRange("A2").getValue();
const pwd = ABA_LOGIN.getRange("B2").getValue();

// VARIÁVEIS DA TABELA
const cancelamentoRow = ABA_CANCELAMENTO.getRange("I1").getValue();
const devolucaoRow    = ABA_DEVOLUÇÃO.   getRange("I1").getValue();
const ABA_DATABASE = SHEET.getSheetByName("DATABASE");

function enterUser(logg, pass){
  if (logg == lgn && pass == pwd){
    SHEET.toast("Login sucedido!");

    var html_contabilidade = HtmlService.createHtmlOutputFromFile("Contabilidade").setWidth(1200).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html_contabilidade, "Contabilidade");
  } else {
    SHEET.toast("Contate o departamento de TI para dúvidas.","Senha/Login incorretas.");
  }
}

function onLogin() {
  var html = HtmlService.createHtmlOutputFromFile("Usuário e Login")
  .setWidth(300)
  .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, "Autorização Contábil");
}

function showData(){
  const data = ABA_DATABASE.getDataRange().getValues();
  return data;
}

function showUI() {
  SpreadsheetApp.getUi().createMenu("Contabilidade").addItem("Autorização Contábil", "onLogin").addToUi();
}

function isAllowed(id, obs){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SOLICITAÇÕES");
  const RANGE = sheet.getDataRange().getValues();
  for (var i = 0; i < RANGE.length; i++){
    if (RANGE[i][1] == id){
      sheet.getRange(i + 1, 3).setValue(new Date);
      sheet.getRange(i + 1, 33).setValue(obs);
      var result = SpreadsheetApp.getUi().alert("Autorização para cancelamento/devolução de documento fiscal concluída.", SpreadsheetApp.getUi().ButtonSet.OK);
      
      var emailBody = "Nova Solicitação na <a href='https://docs.google.com/spreadsheets/d/1zjSCsjd1L87_Ok-8SAlW5Q3Pcc_RpHHUOFe340drrFg/edit#gid=2001053228'>Planilha</a>:<br><br>";
      var tipoTitle = sheet.getRange(2,6).getValue();
      var tipoValue = sheet.getRange(i+1,6).getValue();
      var solicitanteTitle = sheet.getRange(2,4).getValue();
      var solicitanteValue = sheet.getRange(i+1,4).getValue();
      var autorizadoTitle = sheet.getRange(2,5).getValue();
      var autorizadoValue = sheet.getRange(i+1,5).getValue();
      if (tipoValue == "CANCELAMENTO"){
        var notaTitle = sheet.getRange(2,26).getValue();
        var notaValue = sheet.getRange(i+1,26).getValue();
        var motivoTitle = sheet.getRange(2,27).getValue();
        var motivoValue = sheet.getRange(i+1,27).getValue();
        var anexoTitle = sheet.getRange(2,28).getValue();
        var anexoValue = sheet.getRange(i+1,28).getValue();
      }else if (tipoValue =="DEVOLUÇÃO DE MERCADORIA"){
        var notaTitle = sheet.getRange(2,29).getValue();
        var notaValue = sheet.getRange(i+1,29).getValue();
        var motivoTitle = sheet.getRange(2,30).getValue();
        var motivoValue = sheet.getRange(i+1,30).getValue();
        var anexoTitle = sheet.getRange(2,32).getValue();
        var anexoValue = sheet.getRange(i+1,32).getValue();
      }
      emailBody += "<b>"+tipoTitle+"</b><br>"+tipoValue+"<br><br>"+"<b>"+solicitanteTitle+"</b><br>"+solicitanteValue+"<br><br>"+"<b>"+autorizadoTitle+"</b><br>"+autorizadoValue+"<br><br>"+"<b>"+notaTitle+"</b><br>"+notaValue+"<br><br>"+"<b>"+motivoTitle+"</b><br>"+motivoValue+"<br><br>"+"<b>"+anexoTitle+"</b><br>"+anexoValue+"<br><br>"+"<b>OBSERVAÇÃO CONTÁBIL</b>"+"<br>"+obs;

      if (result == SpreadsheetApp.getUi().Button.OK){
        var html_contabilidade = HtmlService.createHtmlOutputFromFile("Contabilidade").setWidth(1200).setHeight(600);
        SpreadsheetApp.getUi().showModalDialog(html_contabilidade, "Contabilidade");
      }
    }
  }

  MailApp.sendEmail({
    to: "fiscal@ativanautica.com.br, faturamento@ativanautica.com.br, contasareceber@ativanautica.com.br",
    subject: "Nova Resposta no Formulário",
    name: "SOLICITAÇÂO PARA FATURAMENTO - "+tipoValue,
    htmlBody: emailBody
  });
}

function notAllowed(id, obs){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SOLICITAÇÕES");
  const RANGE = sheet.getDataRange().getValues();
  for (var i = 0; i < RANGE.length; i++){
    if (RANGE[i][1] == id){
      sheet.getRange(i + 1, 3).setValue("NÃO AUTORIZADO");
      sheet.getRange(i + 1, 33).setValue(obs);
      var result = SpreadsheetApp.getUi().alert("Não Autorização para cancelamento/devolução de documento fiscal concluída.", SpreadsheetApp.getUi().ButtonSet.OK);
      // nao envia email
      if (result == SpreadsheetApp.getUi().Button.OK){
        var html_contabilidade = HtmlService.createHtmlOutputFromFile("Contabilidade").setWidth(1200).setHeight(600);
        SpreadsheetApp.getUi().showModalDialog(html_contabilidade, "Contabilidade");
      }
    }
  }
}

function enviarEmail(){
  const FORM = FormApp.openById("1W4ARlFLwjRTtLGvsO3MWxI7MX46rlKAOShI5jk_nJC0").getResponses();
  
  var resp = FORM[FORM.length - 1].getItemResponses();

  var emailBody = "Nova Solicitação na <a href='https://docs.google.com/spreadsheets/d/1zjSCsjd1L87_Ok-8SAlW5Q3Pcc_RpHHUOFe340drrFg/edit#gid=2001053228'>Planilha</a>:<br><br>";

  resp.forEach(function(resp){
    var title = "<b>" + resp.getItem().getTitle() + "</b>";
    var response = resp.getResponse();
    emailBody += title + "<br>" + response + "<br><br>";
  });

  MailApp.sendEmail({
    to: "fiscal@ativanautica.com.br, faturamento@ativanautica.com.br",
    to: "dev.jtativa@gmail.com",
    subject: "Nova Resposta no Formulário",
    name: "SOLICITAÇÂO PARA FATURAMENTO - "+ resp[2].getResponse(),
    htmlBody: emailBody
  });
}
