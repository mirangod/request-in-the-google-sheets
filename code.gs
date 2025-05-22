function onFormSubmit(){
  var form = FormApp.openById("1W4ARlFLwjRTtLGvsO3MWxI7MX46rlKAOShI5jk_nJC0").getResponses();

  var resp = form[form.length - 1].getItemResponses();

  var emailBody = "Nova Solicitação na <a href='https://docs.google.com/spreadsheets/d/1zjSCsjd1L87_Ok-8SAlW5Q3Pcc_RpHHUOFe340drrFg/edit#gid=2001053228'>Planilha</a>:<br><br>";

  resp.forEach(function(resp){
    var title = "<b>" + resp.getItem().getTitle() + "</b>";
    var response = resp.getResponse();
    emailBody += title + "<br>" + response + "<br><br>";
  });

  switch(resp[2].getResponse()){
    case "CANCELAMENTO":
      MailApp.sendEmail({
        to: "financeiro@ativanautica.com.br, contabilidade@ativanautica.com.br, julia@ativanautica.com.br",
        subject: "Nova Resposta no Formulário",
        name: "SOLICITAÇÃO PARA FATURAMENTO - CANCELAMENTO",
        htmlBody: emailBody + "Suas credenciais para acessar a aba Contabilidade são:<br>Login: " + lgn + "<br>Senha: " + pwd
      });
      break;
    case "DEVOLUÇÃO DE MERCADORIA":
      MailApp.sendEmail({
        to: "financeiro@ativanautica.com.br, contabilidade@ativanautica.com.br, julia@ativanautica.com.br",
        subject: "Nova Resposta no Formulário",
        name: "SOLICITAÇÃO PARA FATURAMENTO - DEVOLUÇÃO DE MERCADORIA",
        htmlBody: emailBody + "Suas credenciais para acessar a aba Contabilidade são:<br>Login: " + lgn + "<br>Senha: " + pwd
      });
      break;
    default:
      MailApp.sendEmail({
        to: "fiscal@ativanautica.com.br, faturamento@ativanautica.com.br, contabilidade@ativanautica.com.br, julia@ativanautica.com.br",
        subject: "Nova Resposta no Formulário",
        name: "SOLICITAÇÃO PARA FATURAMENTO - REMESSA/RETORNO/CARTA DE CORREÇÃO",
        htmlBody: emailBody
      });
      break;
  }
}

function EMAILSENVIADOS(e) {

  var aba = e.range.getSheet();

  var planilha = SpreadsheetApp.openById("1zjSCsjd1L87_Ok-8SAlW5Q3Pcc_RpHHUOFe340drrFg")
  var link = planilha.getUrl()

  var linha = e.range.getRow()

  var valores = aba.getRange(linha + ":" + linha).getValues()[0]
  var cab = aba.getRange(2 + ":" + 2).getValues()[0]

  for (let i = 0; i < cab.length; i++) {
      if(cab[i] == "VALOR")
      {  
        valores[i]="R$"+valores[i]
      }

  }

  valores[0] = new Date().toLocaleString("pt-BR", {timeZone: "America/Sao_Paulo"})
 
  //MailApp.sendEmail("suporte02.jtativa@gmail.com","TESTES","...")

  var t = HtmlService.createTemplateFromFile("modelo")
  t.dados = valores
  t.cab = cab
  t.link = link
  var corpo = t.evaluate().getContent()

  // Verificação para caso haja cancelamento/devolução de mercadoria
  var form = FormApp.openById("1W4ARlFLwjRTtLGvsO3MWxI7MX46rlKAOShI5jk_nJC0").getResponses();
  var resp = form[form.length -1].getItemResponses();

  for (var i in resp){
    var j = resp[i].getItem().getTitle();
    var h = resp[i].getResponse();

    if (j == "TIPO"){
      if (h == "CANCELAMENTO" || h == "DEVOLUÇÂO DE MERCADORIA"){
        const mensagem = {
          "to": "fiscal@ativanautica.com.br, faturamento@ativanautica.com.br, contabilidade@ativanautica.com.br, julia@ativanautica.com.br, contasareceber@ativanautica.com.br, ativacomercial@ativanautica.com.br",
          "subject": "Nova Solicitação",
          "name": "SOLICITAÇÃO PARA FATURAMENTO",
          "htmlBody": corpo
        }

        MailApp.sendEmail(mensagem);
      } else {
        const mensagem = {
          "to": "fiscal@ativanautica.com.br, faturamento@ativanautica.com.br, contabilidade@ativanautica.com.br, julia@ativanautica.com.br, ativacomercial@ativanautica.com.br",
          "subject": "Nova Solicitação",
          "name": "SOLICITAÇÃO PARA FATURAMENTO",
          "htmlBody": corpo
        }

        MailApp.sendEmail(mensagem);
      }
    }
  }
}

function teste(){
  var form = FormApp.openById("1W4ARlFLwjRTtLGvsO3MWxI7MX46rlKAOShI5jk_nJC0").getResponses(); // pega as respostas

  var resp = form[form.length -1].getItemResponses(); // pega os itens da ultima resposta

  for (var item in resp){
    var j = resp[item].getItem().getTitle();
    var h = resp[item].getResponse();

    if (j == "TIPO"){
      if (h == "CANCELAMENTO" || h == "DEVOLUÇÃO DE MERCADORIA"){
        Logger.log("Teste");
      }else{
        Logger.log("Teste 1");
      }
    }
  }
}
