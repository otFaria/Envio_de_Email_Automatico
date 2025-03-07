function safeDecodeURI(uri) {
  try {
    // Remove caracteres "%" que não fazem parte de uma sequência válida
    var safeUri = uri.replace(/%(?![0-9A-Fa-f]{2})/g, ""); 
    return decodeURIComponent(safeUri);
  } catch (e) {
    console.error("Erro ao decodificar URI:", e.message);
    return uri; // Retorna a string original caso falhe
  }
}

function getEmailData(dados) {
  var email = dados.match(/mailto:([^?]+)/);
  var mesage = dados.match(/body=(.*)/);
  var subject = dados.match(/subject=([^&]+)/);

  if (!email || !mesage || !subject) {
    throw new Error("Dados do e-mail incompletos ou inválidos.");
  }

  var safeUri = safeDecodeURI(mesage[1]);
  var body = decodeURIComponent(safeUri).replace(/\n/g, "<br>");
  var subjectFormat = decodeURIComponent(subject[1]);

  var testando = body.match(/0(.*)+/);
  var mensagem2 = "Prezado(a) usuario(a):<br><br>" + testando[1];

  return {
    email: email[1],
    subject: subjectFormat,
    body: mensagem2
  };
}

function sendEmail(emailData) {
  try {
    MailApp.sendEmail({
      to: emailData.email,
      subject: emailData.subject,
      htmlBody: emailData.body,
      name: "Biblioteca Jofre Moreira"
    });
    return "Enviado";
  } catch (error) {
    Logger.log("❌ Erro ao enviar e-mail para " + emailData.email + ": " + error.message);
    return "Erro";
  }
}

function processSheet(sheet) {
  var emailColumn = 6; // Coluna onde estão os e-mails
  var statusColumn = 10; // Coluna para status de envio
  var numRows = sheet.getLastRow();

  if (numRows == 1) {
    numRows = numRows - 1;
  }

  if (numRows != 0) {
    for (var i = 2; i <= numRows; i++) {
      var cell = sheet.getRange(i, emailColumn);
      var richText = cell.getRichTextValue();
      var dados = richText.getLinkUrl();

      if (dados) {
        try {
          var emailData = getEmailData(dados);
          var status = sendEmail(emailData);
          sheet.getRange(i, statusColumn).setValue(status);
          Logger.log("✅ E-mail enviado para: " + emailData.email);
        } catch (error) {
          Logger.log("❌ Erro ao processar e-mail: " + error.message);
          sheet.getRange(i, statusColumn).setValue("Erro");
        }
      } else {
        Logger.log("Email vazio ou não cadastrado.");
      }
    }
  } else {
    Logger.log("Planilha vazia");
  }
}

function myFunction() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página1");
  processSheet(sheet);
}