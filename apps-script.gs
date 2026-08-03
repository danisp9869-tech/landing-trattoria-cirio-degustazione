/* ═══════════════════════════════════════════════════════════════════════════
   ANTICA TRATTORIA CIRIO — Web App che riceve le prenotazioni
   ───────────────────────────────────────────────────────────────────────────
   Un solo Sheet per tutte le landing Cirio. La colonna "Offerta" dice da quale
   proposta arriva la prenotazione:

     "Degustazione"          → menudegustazione.trattoriacirio.it
     "Prenotazione normale"  → prenotazioni.trattoriacirio.it e prenota.trattoriacirio.it

   Il valore lo manda la landing nel campo `origine` (CONFIG.origine).

   COME SI PUBBLICA (senza questo passaggio non cambia nulla):
     1. Apri lo Sheet → Estensioni → Apps Script
     2. Incolla questo file al posto del codice esistente
     3. Distribuisci → Gestisci distribuzioni → matita → Versione: "Nuova versione"
        → Distribuisci.  L'URL /exec resta lo stesso.
     4. Verifica aprendo l'URL nel browser: deve rispondere "… · v2 (colonna Offerta)".
        Se leggi una versione diversa, la distribuzione è ancora quella vecchia.

   NOTA: lo script è legato al foglio (container-bound), quindi scrive sul foglio
   che lo contiene. Non serve nessun ID.
   ═══════════════════════════════════════════════════════════════════════════ */

// ⚠️ Indirizzo che riceve la notifica di ogni nuova prenotazione.
var EMAIL_RISTORATORE = 'info@trattoriacirio.it';

var INTESTAZIONI = ['Data','Ora','Persone','Nome','Telefono','Email','Richieste','Privacy','Creata','Offerta'];

function doPost(e) {
  // Due prenotazioni nello stesso istante scriverebbero sulla stessa riga
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // La landing invia un corpo JSON (text/plain per evitare il preflight CORS)
    var d = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(INTESTAZIONI);
      sheet.setFrozenRows(1);
    } else {
      // Foglio nato prima che esistesse la colonna: la aggiunge in fondo una
      // volta sola, così le righe già presenti restano allineate
      var ultimaCol = sheet.getLastColumn();
      var testate = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0];
      if (testate.indexOf('Offerta') === -1) {
        sheet.getRange(1, ultimaCol + 1).setValue('Offerta');
      }
    }

    sheet.appendRow([
      // "aaaa-mm-gg" come nelle righe già presenti: cambiare formato qui
      // spezzerebbe l'ordinamento del foglio e la lettura dei coperti occupati.
      // La versione leggibile ("sab 08/08/2026") va solo nella mail.
      d.data || '',
      d.ora || '', d.persone || '', d.nome || '', d.telefono || '',
      d.email || '', d.richieste || '', d.privacy || '', d.creata || '',
      d.origine || ''                // ← "Degustazione" / "Prenotazione normale"
    ]);

    // Notifica al ristoratore. Se salta, la riga sul foglio resta comunque.
    try {
      var offerta = d.origine || '';
      var oggetto = '🍽️ ' + (offerta ? offerta + ' — ' : '') + 'Nuova prenotazione — ' +
                    (d.nome || '') + ' · ' + (d.dataLabel || d.data || '') + ' ' +
                    (d.ora || '') + ' · ' + (d.persone || '') + 'p';
      var html =
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;line-height:1.5">' +
          '<h2 style="margin:0 0 4px">Nuova prenotazione — Antica Trattoria Cirio</h2>' +
          '<p style="margin:0 0 14px;color:#666">Ricevuta dal sito delle prenotazioni</p>' +
          '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #eee">' +
            riga_('Offerta', offerta) +
            riga_('Data', d.dataLabel || d.data) + riga_('Orario', d.ora) +
            riga_('Persone', d.persone) + riga_('Nome', d.nome) +
            riga_('Telefono', d.telefono) + riga_('Email', d.email) +
            riga_('Richieste', d.richieste) +
          '</table>' +
        '</div>';
      var opzioni = { htmlBody: html, name: 'Prenotazioni Antica Trattoria Cirio' };
      // Rispondere alla mail scrive direttamente al cliente
      if (d.email && d.email.indexOf('@') > 0) opzioni.replyTo = d.email;
      MailApp.sendEmail(EMAIL_RISTORATORE, oggetto, 'Nuova prenotazione ricevuta.', opzioni);
    } catch (erroreMail) {}

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
                         .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Una riga della tabella nella mail; se il valore è vuoto la salta
function riga_(etichetta, valore) {
  if (!valore) return '';
  return '<tr><td style="border:1px solid #eee;color:#666;font-weight:bold">' + etichetta +
         '</td><td style="border:1px solid #eee">' + valore + '</td></tr>';
}

/* Aprendo l'URL /exec nel browser si legge la versione pubblicata: salvare in
   Apps Script non basta, bisogna ridistribuire. Se qui non compare "v2", la
   distribuzione attiva è ancora quella vecchia.
   Con ?action=occupati restituisce i coperti già prenotati, per l'indicazione
   "ultimi posti" dal vivo (serve CONFIG.integrazione.leggiOccupatiDaSheets = true,
   oggi disattivata su tutte le landing Cirio). */
function doGet(e) {
  var azione = (e && e.parameter && e.parameter.action) || '';
  if (azione !== 'occupati') {
    return ContentService.createTextOutput(
      'Antica Trattoria Cirio — endpoint prenotazioni attivo · v2 (colonna Offerta)');
  }
  var righe = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
  var out = {};
  for (var i = 1; i < righe.length; i++) {
    var data = righe[i][0], ora = righe[i][1], persone = righe[i][2];
    if (!data) continue;
    var k = chiaveData_(data);
    if (!k) continue;
    out[k] = out[k] || {};
    out[k][ora] = (out[k][ora] || 0) + Number(persone || 0);
  }
  return ContentService.createTextOutput(JSON.stringify(out))
                       .setMimeType(ContentService.MimeType.JSON);
}

/* La colonna Data può contenere "gg/mm/aaaa" (o "sab 08/08/2026") oppure una
   data vera di Sheets: la landing si aspetta sempre "aaaa-mm-gg". */
function chiaveData_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Rome', 'yyyy-MM-dd');
  var m = String(v).match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}
