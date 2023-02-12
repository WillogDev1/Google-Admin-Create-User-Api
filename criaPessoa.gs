function criaEmail() {

// Pega planilha
var sheet = SpreadsheetApp.getActiveSheet();

// Obtem todos os dados da planilha
var data = sheet.getDataRange().getValues();

// Pega ultima coluna que corresponde ao e-mail
var coluna = sheet.getLastColumn();

// pega ultima linha inserida
var linha = sheet.getLastRow();

// Transforma os dados em um objeto JSON
var jsonData = JSON.stringify(data);

// Palavras para ignorar na hora de criar e-mail
const ignoredWords = ["de", "do", "da", "dos", "das"];

// Pega ultimo nome completo inserido da planilha // Usado para criar E-mail
var nomeCompleto = sheet.getRange(linha,coluna).getValue();
// Pega ultimo nome completo inserido da planilha // Usado para criar Primeiro nome e Sobrenome
var nomeCompletoParaSobreNome = sheet.getRange(linha,coluna).getValue();

// Transformo em string
nomeCompleto.toString();

// Dividi o nome completo em partes e filtra retirando as palavras que coincidem com ignoredWords
let nomeParte = nomeCompleto.split(" ").filter(word => !ignoredWords.includes(word));

// Dividi o nome sem o primeiro nome, formando o sobrenome completo
let nomeParteParaSobreNome = nomeCompletoParaSobreNome.split(" ");

//Use seu Dominio no final
var email1 = nomeParte[0] + "." + nomeParte[nomeParte.length -1] + "@wgibram.com";
var email2 = nomeParte[0] + "." + nomeParte[nomeParte.length -2] + "@wgibram.com";
var email3 = nomeParte[0] + "." + nomeParte[nomeParte.length -3] + "@wgibram.com";
var email4 = nomeParte[0] + "." + nomeParte[nomeParte.length -4] + "@wgibram.com";

// Dividi em pedaços e ignora o primeiro
const restoNome = nomeParteParaSobreNome.slice(1);

// Junta todo o resto(Tirando o primeiro nome) com espaço entre eles
const restoNomeS = restoNome.join(" ").toString();

// Pega o primeiro nome
const firstName = nomeParte[0].toString();

// Variavel para ser usado no Switch(Case)
var count = 0;

    // itera pelo JSON a procura de igualdade de string
    for (var i = 0; i < data.length; i++) {
        for (var key in data[i]) {
          // Se encontrar adiciona +1 a count
          if (data[i][key] === email1.toLowerCase()){
                count++;
            }
          if (data[i][key] === email2.toLowerCase()){
                count++;
            }
          if (data[i][key] === email3.toLowerCase()){
                count++;
            }
          if (data[i][key] === email4.toLowerCase()){
                count++;
            }
        }
      }

    // Switch para enviar a criação de e-mail sem duplicatas
      switch (count){
        case 0:
          console.log("opcao0: " + email1);
          // Api que envia o email nome e sobrenome para o google
          const send1 = AdminDirectory.Users.insert(
              {
              "name": {
                "familyName": restoNomeS,
                "givenName": firstName
              },
              "password": "12345678",
              "primaryEmail": email1
            }
          );
          //Imprime na planilha
          const newValue = sheet.getRange(linha,2).setValue(email1.toLocaleLowerCase());
          const newValue2 = sheet.getRange(linha,2).getValue();
          break;
        case 1:
          console.log("opcao1: " + email2);
                    const send2 = AdminDirectory.Users.insert(
              {
              "name": {
                "familyName":  restoNomeS,
                "givenName": firstName
              },
              "password": "12345678",
              "primaryEmail": email2
            }
          );
          const newValue3 = sheet.getRange(linha,2).setValue(email2.toLocaleLowerCase());
          const newValue4 = sheet.getRange(linha,2).getValue();
          break;
        case 2:
          console.log("opcao2: " + email3);
                    const send3 = AdminDirectory.Users.insert(
              {
              "name": {
                "familyName": restoNomeS,
                "givenName": firstName
              },
              "password": "12345678",
              "primaryEmail": email3
            }
          );
          const newValue5 = sheet.getRange(linha,2).setValue(email3.toLocaleLowerCase());
          const newValue6 = sheet.getRange(linha,2).getValue();
          break;
        case 3:
          console.log("opcao3:" + email4);
                    const send4 = AdminDirectory.Users.insert(
              {
              "name": {
                "familyName": restoNomeS,
                "givenName": firstName
              },
              "password": "12345678",
              "primaryEmail": email4
            }
          );
          const newValue7 = sheet.getRange(linha,2).setValue(email4.toLocaleLowerCase());
          const newValue8 = sheet.getRange(linha,2).getValue();
          break;
    }



}
