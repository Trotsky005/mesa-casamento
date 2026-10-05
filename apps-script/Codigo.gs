/**
 * Mesas Henrique & Livia — script da planilha
 * Cole este código num projeto novo em script.google.com e implante como "App da Web".
 *
 * A planilha precisa das colunas: Nome | Mesa | Chegou | Quem é  (na primeira aba)
 */

// Planilha "Mesas do Casamento (site)" no Google Drive
var PLANILHA_ID = '1x9sVz9o3HLyXUlqZYYuPkZYtZKLWIh3UBlCJ03sa3EY';

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.acao === 'chegada') return json(marcarChegada(p.nome, p.mesa, p.valor === '1'));
    return json(lerLista());
  } catch (err) {
    return json({ erro: String(err) });
  }
}

function aba() {
  var sh = SpreadsheetApp.openById(PLANILHA_ID).getSheets()[0];
  if (String(sh.getRange(1, 3).getValue()).trim() === '') sh.getRange(1, 3).setValue('Chegou');
  if (String(sh.getRange(1, 4).getValue()).trim() === '') sh.getRange(1, 4).setValue('Quem é');
  return sh;
}

function lerLista() {
  var sh = aba();
  var n = sh.getLastRow();
  var lista = [];
  if (n >= 2) {
    var vals = sh.getRange(2, 1, n - 1, 4).getDisplayValues();
    vals.forEach(function (r) {
      var nome = String(r[0]).trim(), mesa = String(r[1]).trim().replace(/^mesa\s*/i, '');
      if (nome && mesa) lista.push({ n: nome, m: mesa, c: String(r[2]).trim(), q: String(r[3]).trim() });
    });
  }
  return { guests: lista, t: Date.now() };
}

function marcarChegada(nome, mesa, chegou) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = aba();
    var n = sh.getLastRow();
    if (n >= 2) {
      var vals = sh.getRange(2, 1, n - 1, 2).getDisplayValues();
      var alvoN = norm(nome), alvoM = norm(String(mesa).replace(/^mesa\s*/i, ''));
      for (var i = 0; i < vals.length; i++) {
        if (norm(vals[i][0]) === alvoN && norm(String(vals[i][1]).replace(/^mesa\s*/i, '')) === alvoM) {
          var hora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'HH:mm');
          sh.getRange(i + 2, 3).setValue(chegou ? hora : '');
          break;
        }
      }
    }
    return lerLista();
  } finally {
    lock.releaseLock();
  }
}

function norm(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
