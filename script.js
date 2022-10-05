//Recebendo arquivo anexado ao e-mail e salvando em uma pasta no Drive.
//Necessário add serviços DRIVE API

const relatorio_301 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Relatorio_301");
const now = new Date();

//____________________________________________________________________________________
//Pegar relatório 301 do e-mail e salvar na pasta "CONSULTA"
function getRel301() {
  let idPastaDestino = "1upBCebUcIPZVUA8Ch5YfPS7YXZQyJ6zc"; // Pasta "CONSULTA"
  let pastaDestino = DriveApp.getFolderById(idPastaDestino)
  let assuntoAProcurar = 'Arquivo(s) anexo(s).';
  let lista = GmailApp.search(assuntoAProcurar);
  let attachments = GmailApp.getMessageById(lista[0].getId()).getAttachments();
  let files = pastaDestino.getFiles()
  let encontrou = false;
  let l = 0;
  if (relatorio_301.getLastRow() > 1) { l = 1 }
  let dados = relatorio_301.getRange(2, 1, relatorio_301.getLastRow() - l, 3).getValues();
  let ul = ultimoLinhaColuna(relatorio_301, 3) + 1
  if (lista[0].getFirstMessageSubject().includes(assuntoAProcurar)) {

    //Verifica se o arquivo já foi registrado na planilha.
    for (let i = 0; i < dados.length; i++) {
      if (dados[i][0] == attachments[0].getHash()) {
        encontrou = true;
        Logger.log('Arquivo já foi salvo na pasta.')
        return;
      }
    }

    //Se o arquivo não foi registrado 
    if (encontrou == false) {
      //Excluir arquivo
      Logger.log('Excluindo arquivo antigo...')
      while (files.hasNext()) {
        let nomeArquivo = files.next();
        let id = nomeArquivo.getId()

        //Excluir relatório 301 da pasta "CONSULTA"
        if (nomeArquivo == 'SCSC301.XLSX') {
          Logger.log('Excluindo arquivo ' + nomeArquivo)
          DriveApp.getFolderById(id).setTrashed(true);
        }
      }

      //Salvar arquivo na pasta
      Logger.log('Salvando novo arquivo...')
      let arquivo = DriveApp.createFile(attachments[0].copyBlob());
      arquivo.setName('SCSC301.XLSX');
      arquivo.moveTo(pastaDestino);

      //Inserir informações na planilha
      Logger.log('Registrando na planilha...')
      relatorio_301.getRange(ul, 1).setValue(attachments[0].getHash());
      relatorio_301.getRange(ul, 2).setValue(now);
      relatorio_301.getRange(ul, 3).setValue('Relatório salvo na pasta');

    }

  }

}

//____________________________________________________________________________________
//Pegar última linha de uma coluna específica.
function ultimoLinhaColuna(planilha, coluna) {
  x = 1; do { ; x++; }
  while (planilha.getRange(x, coluna).getValue() != "");
  return (x - 1)
}
