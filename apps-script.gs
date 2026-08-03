/* ═══════════════════════════════════════════════════════════════════════════
   ANTICA TRATTORIA CIRIO — Web App che riceve le prenotazioni
   ───────────────────────────────────────────────────────────────────────────
   Un solo foglio per tutte le landing Cirio. La colonna "Offerta" (10ª) dice
   da quale proposta arriva la prenotazione:

     "Degustazione"          → menudegustazione.trattoriacirio.it
     "Prenotazione normale"  → prenotazioni.trattoriacirio.it e prenota.trattoriacirio.it

   Il valore lo manda la landing nel campo `origine` (CONFIG.origine).

   COME SI PUBBLICA (senza questo passaggio non cambia nulla):
     1. Foglio → Estensioni → Apps Script
     2. Incolla questo file al posto del codice esistente
     3. Distribuisci → Gestisci distribuzioni → matita → Versione: "Nuova versione"
        → Distribuisci.  L'URL /exec resta lo stesso.
     4. Verifica aprendo l'URL /exec nel browser: deve rispondere
        "… · v3 (offerta + consenso marketing)".
   ═══════════════════════════════════════════════════════════════════════════ */

const SHEET_ID = "1RUQqs4gOqCHNZB-Cnx2wuwFdD5J12BOphyNmh68l5_s";

// Indirizzo che riceve la notifica di ogni nuova prenotazione
const EMAIL_RISTORATORE = "info@trattoriacirio.it";

function doPost(e) {
  // Due prenotazioni nello stesso istante scriverebbero sulla stessa riga
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const d = JSON.parse(e.postData.contents);

    // "Degustazione" oppure "Prenotazione normale"
    const offerta = d.origine || "—";
    // Data leggibile per la mail ("sab 08/08/2026"); sul foglio resta aaaa-mm-gg
    const dataLeggibile = d.dataLabel || d.data;

    sh.appendRow([
      d.data,
      d.ora,
      d.persone,
      d.nome,
      d.telefono,
      d.email,
      d.richieste,
      d.privacy,
      d.creata,
      d.origine,      // ← colonna "Offerta"
      d.sorgente,
      d.marketing,    // ← colonna "Marketing": consenso facoltativo alle offerte
    ]);

    // Se la mail fallisce, la riga sul foglio deve restare comunque
    try {
      MailApp.sendEmail({
        to: EMAIL_RISTORATORE,
        // L'offerta è anche nell'oggetto: si distingue dalla lista senza aprire
        subject: `🍽️ ${offerta} — Nuova prenotazione: ${d.nome || ""} · ${dataLeggibile || ""} ${d.ora || ""} · ${d.persone || ""}p`,
        // Rispondendo alla mail si scrive direttamente al cliente
        replyTo: (d.email && d.email.indexOf("@") > 0) ? d.email : undefined,
        htmlBody: `
      <h2>Nuova prenotazione ricevuta</h2>

      <p style="font-size:16px"><strong>🏷️ Offerta:</strong>
        <span style="background:#f0e7d2;padding:3px 10px;border-radius:6px;font-weight:bold">${offerta}</span>
      </p>

      <hr>

      <p><strong>📅 Data:</strong> ${dataLeggibile}</p>
      <p><strong>🕒 Ora:</strong> ${d.ora}</p>
      <p><strong>👥 Persone:</strong> ${d.persone}</p>

      <hr>

      <p><strong>🙍 Nome:</strong> ${d.nome}</p>
      <p><strong>📞 Telefono:</strong> ${d.telefono}</p>
      <p><strong>✉️ Email:</strong> ${d.email}</p>

      <p><strong>📝 Richieste:</strong><br>
      ${d.richieste || "Nessuna richiesta"}
      </p>

      <p><strong>📣 Consenso marketing:</strong>
        ${d.marketing === "Sì" ? "✅ Sì, vuole ricevere offerte e novità" : "❌ No"}
      </p>
    `
      });
    } catch (erroreMail) {}

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/* Aprendo l'URL /exec nel browser si legge la versione pubblicata: salvare in
   Apps Script non basta, bisogna ridistribuire. Se non compare "v3", la
   distribuzione attiva è ancora quella vecchia. */
function doGet() {
  return ContentService.createTextOutput(
    "Antica Trattoria Cirio — endpoint prenotazioni attivo · v3 (offerta + consenso marketing)");
}

/* ─── Utilità da lanciare a mano dall'editor, non servono al funzionamento ─── */

// Scrive la riga delle intestazioni. Da eseguire UNA VOLTA se in cima al foglio
// mancano i titoli delle colonne. Non tocca le prenotazioni già presenti.
function sistemaIntestazioni() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sh.getRange(1, 1, 1, 12).setValues([[
    "Data", "Ora", "Persone", "Nome", "Telefono", "Email",
    "Richieste", "Privacy", "Creata", "Offerta", "Sorgente", "Marketing"
  ]]);
  sh.setFrozenRows(1);
}

function testMail() {
  MailApp.sendEmail(
    "dani.sp9869@gmail.com",
    "Test Apps Script",
    "Se ricevi questa email, MailApp funziona."
  );
}
