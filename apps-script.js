// ═══════════════════════════════════════════════════════════════
// MATRIMONIO — Sistema Inviti + RSVP
// Google Apps Script — incolla in Estensioni → Apps Script
// ═══════════════════════════════════════════════════════════════
//
// SETUP GOOGLE SHEET:
// 1. Crea un nuovo Google Sheet
// 2. Rinomina il primo tab → "Invitati"
//    Intestazioni riga 1:
//    A: Nome | B: Cognome | C: Email | D: Telefono | E: Ospiti Previsti | F: Stato | G: Data Invio | H: Data Risposta
//
// 3. Crea un secondo tab → "Risposte"
//    Intestazioni riga 1:
//    A: Timestamp | B: Nome | C: Cognome | D: Email | E: Telefono | F: Partecipa | G: Ospiti | H: Note
//
// 4. Compila il tab "Invitati" con i dati degli ospiti (colonne A-D),
//    lascia E-F-G vuote (verranno compilate automaticamente)
//    Nella colonna E (Stato) metti "Da invitare" per chi deve ricevere l'invito
//
// DEPLOY:
// 1. Estensioni → Apps Script → incolla tutto questo codice
// 2. Deploy → Nuova distribuzione → Tipo: App Web
//    - Esegui come: Me
//    - Chi ha accesso: Chiunque
// 3. Copia l'URL generato (tipo: https://script.google.com/macros/s/XXXX/exec)
// 4. Incollalo nel sito wedding.html nella variabile SCRIPT_URL
// 5. Ricarica il foglio: comparirà il menu "🎊 Matrimonio"
// ═══════════════════════════════════════════════════════════════

// ── CONFIGURAZIONE (✏️ PERSONALIZZA QUESTI VALORI) ──────────

const CONFIG = {
  URL_SITO: 'https://gianlucaederikasposi.it',  // URL dove hosti il sito
  NOME_SPOSI: 'Gianluca & Erika',
  DATA_MATRIMONIO: '26 Agosto 2027',
  LUOGO_CERIMONIA: 'Basilica di Capodimonte, Napoli',
  ORA_CERIMONIA: '14:30',
  FOGLIO_INVITATI: 'Invitati',
  FOGLIO_RISPOSTE: 'Risposte',
  CARTELLA_FOTO: 'Foto Matrimonio',         // Nome cartella in Google Drive
  ALBUM_GOOGLE_PHOTOS: 'Matrimonio Gianluca ed Erika',  // Nome album su Google Photos
};


// ── MENU PERSONALIZZATO ─────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎊 Matrimonio')
    .addItem('📨 Invia inviti (solo "Da invitare")', 'inviaInviti')
    .addItem('🔔 Invia solleciti (solo "Invitato")', 'inviaSolleciti')
    .addSeparator()
    .addItem('📊 Riepilogo risposte', 'mostraRiepilogo')
    .addToUi();
}


// ── RICEZIONE DATI (endpoint web app) ───────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Smista in base al campo "mode"
    if (data.mode === 'upload') {
      return gestisciUploadFoto(data);
    } else {
      return gestisciRSVP(data);
    }

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ── GESTIONE RSVP ───────────────────────────────────────────

function gestisciRSVP(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Validazione: verifica che l'email sia nella lista Invitati ──
  if (!data.email || !isInvitatoAutorizzato(ss, data.email)) {
    // Rifiuta silenziosamente — non registra la risposta
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const risposte = ss.getSheetByName(CONFIG.FOGLIO_RISPOSTE);
  risposte.appendRow([
    new Date(),
    data.nome || '',
    data.cognome || '',
    data.email || '',
    data.telefono || '',
    data.partecipa || '',
    data.ospiti || 1,
    data.note || ''
  ]);

  aggiornaStatoInvitato(
    ss,
    data.email,
    data.partecipa === 'si' ? '✅ Confermato' : '❌ Rifiutato'
  );

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── Verifica se l'email è nella lista invitati ─────────────
function isInvitatoAutorizzato(ss, email) {
  const sheet = ss.getSheetByName(CONFIG.FOGLIO_INVITATI);
  const dati = sheet.getDataRange().getValues();
  const emailLow = email.toLowerCase().trim();

  for (let i = 1; i < dati.length; i++) {
    const emailInvitato = (dati[i][2] || '').toString().toLowerCase().trim();
    if (emailInvitato === emailLow) return true;
  }
  return false;
}


// ── GESTIONE UPLOAD FOTO ────────────────────────────────────

function gestisciUploadFoto(data) {
  // data.file     = stringa base64 del file
  // data.filename = nome originale del file
  // data.mimeType = tipo MIME (image/jpeg, video/mp4, ecc.)
  // data.nome     = nome dell'ospite che carica

  if (!data.file || !data.filename) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'File mancante' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Trova o crea la cartella principale
  const cartellaPrincipale = getOrCreateFolder(CONFIG.CARTELLA_FOTO);

  // Crea sottocartella per l'ospite (raggruppa le foto per persona)
  const nomeOspite = (data.nome || 'Anonimo').trim();
  const cartellaOspite = getOrCreateFolder(nomeOspite, cartellaPrincipale);

  // Decodifica base64 e salva su Drive
  const decoded = Utilities.base64Decode(data.file);
  const blob = Utilities.newBlob(decoded, data.mimeType || 'image/jpeg', data.filename);
  const file = cartellaOspite.createFile(blob);

  // ── Copia automatica su Google Photos ──
  let photosStatus = 'skipped';
  try {
    photosStatus = uploadToGooglePhotos(blob, data.filename);
  } catch (err) {
    photosStatus = 'error: ' + err.toString();
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      fileId: file.getId(),
      fileName: file.getName(),
      googlePhotos: photosStatus
    }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── GOOGLE PHOTOS: Upload automatico nell'album ────────────

function uploadToGooglePhotos(blob, filename) {
  const token = ScriptApp.getOAuthToken();

  // Step 1: Trova l'album per nome (cerca una sola volta, poi usa cache)
  const albumId = getAlbumIdByName(CONFIG.ALBUM_GOOGLE_PHOTOS, token);
  if (!albumId) return 'album_not_found';

  // Step 2: Upload dei bytes raw
  const uploadRes = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/uploads', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/octet-stream',
      'X-Goog-Upload-Protocol': 'raw',
      'X-Goog-Upload-File-Name': filename,
    },
    payload: blob.getBytes(),
    muteHttpExceptions: true,
  });

  if (uploadRes.getResponseCode() !== 200) {
    return 'upload_failed_' + uploadRes.getResponseCode();
  }

  const uploadToken = uploadRes.getContentText();

  // Step 3: Crea il media item nell'album
  const createRes = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/mediaItems:batchCreate', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      albumId: albumId,
      newMediaItems: [{
        simpleMediaItem: {
          uploadToken: uploadToken,
          fileName: filename,
        }
      }]
    }),
    muteHttpExceptions: true,
  });

  if (createRes.getResponseCode() !== 200) {
    return 'create_failed_' + createRes.getResponseCode();
  }

  return 'ok';
}


function getAlbumIdByName(albumName, token) {
  // Cerca tra gli album esistenti
  let nextPageToken = '';
  do {
    const url = 'https://photoslibrary.googleapis.com/v1/albums?pageSize=50'
      + (nextPageToken ? '&pageToken=' + nextPageToken : '');

    const res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() !== 200) return null;

    const json = JSON.parse(res.getContentText());
    const albums = json.albums || [];

    for (const album of albums) {
      if (album.title === albumName) return album.id;
    }

    nextPageToken = json.nextPageToken || '';
  } while (nextPageToken);

  return null;
}


// ── UTILITY: Trova o crea cartella Drive ────────────────────

function getOrCreateFolder(nome, parent) {
  const root = parent || DriveApp.getRootFolder();
  const iter = root.getFoldersByName(nome);
  if (iter.hasNext()) return iter.next();
  return root.createFolder(nome);
}

// Permetti anche GET per test
function doGet(e) {
  return ContentService
    .createTextOutput('Matrimonio API attiva ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}


// ── INVIO INVITI ─────────────────────────────────────────────

function inviaInviti() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.FOGLIO_INVITATI);
  const dati = sheet.getDataRange().getValues();
  
  // Conta quanti da inviare
  let daInviare = 0;
  for (let i = 1; i < dati.length; i++) {
    const stato = (dati[i][5] || '').toString().trim();
    if (stato === 'Da invitare' || stato === '') daInviare++;
  }

  if (daInviare === 0) {
    ui.alert('Nessun invito da inviare.\n\nAssicurati che la colonna Stato contenga "Da invitare".');
    return;
  }

  const conferma = ui.alert(
    'Conferma invio',
    `Stai per inviare ${daInviare} inviti via email.\n\nProcedere?`,
    ui.ButtonSet.YES_NO
  );
  if (conferma !== ui.Button.YES) return;

  let inviati = 0;
  let errori = 0;

  for (let i = 1; i < dati.length; i++) {
    const [nome, cognome, email, telefono, ospitiPrevisti, stato] = dati[i];
    const statoStr = (stato || '').toString().trim();

    if (statoStr !== 'Da invitare' && statoStr !== '') continue;
    if (!email) continue;

    try {
      const linkPersonale = `${CONFIG.URL_SITO}?e=${encodeURIComponent(email)}&n=${encodeURIComponent(nome)}&c=${encodeURIComponent(cognome)}`;
      
      const html = creaEmailInvito(nome, cognome, linkPersonale);

      MailApp.sendEmail({
        to: email,
        subject: `${CONFIG.NOME_SPOSI} - Siete invitati al nostro matrimonio`,
        htmlBody: html,
        name: CONFIG.NOME_SPOSI,        // Il mittente appare come "Gianluca & Erika"
        replyTo: email,                  // Reply va all'invitato stesso (evita bounce)
      });

      // Aggiorna stato e data invio
      sheet.getRange(i + 1, 6).setValue('Invitato');       // Col F — Stato
      sheet.getRange(i + 1, 7).setValue(new Date());       // Col G — Data Invio
      inviati++;

    } catch (err) {
      sheet.getRange(i + 1, 6).setValue('⚠️ Errore invio');
      errori++;
    }
  }

  ui.alert(`Invio completato!\n\n✅ Inviati: ${inviati}\n❌ Errori: ${errori}`);
}


// ── INVIO SOLLECITI ──────────────────────────────────────────

function inviaSolleciti() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.FOGLIO_INVITATI);
  const dati = sheet.getDataRange().getValues();

  let daSollecitare = 0;
  for (let i = 1; i < dati.length; i++) {
    if ((dati[i][5] || '').toString().trim() === 'Invitato') daSollecitare++;
  }

  if (daSollecitare === 0) {
    ui.alert('Tutti hanno già risposto! 🎉');
    return;
  }

  const conferma = ui.alert(
    'Conferma solleciti',
    `${daSollecitare} invitati non hanno ancora risposto.\n\nInviare promemoria?`,
    ui.ButtonSet.YES_NO
  );
  if (conferma !== ui.Button.YES) return;

  let inviati = 0;

  for (let i = 1; i < dati.length; i++) {
    const [nome, cognome, email, telefono, , stato] = dati[i];
    if ((stato || '').toString().trim() !== 'Invitato') continue;
    if (!email) continue;

    try {
      const linkPersonale = `${CONFIG.URL_SITO}?e=${encodeURIComponent(email)}&n=${encodeURIComponent(nome)}&c=${encodeURIComponent(cognome)}`;
      const html = creaEmailSollecito(nome, cognome, linkPersonale);

      MailApp.sendEmail({
        to: email,
        subject: `${CONFIG.NOME_SPOSI} - Conferma la tua presenza`,
        htmlBody: html,
        name: CONFIG.NOME_SPOSI,
      });
      inviati++;
    } catch (err) { /* skip */ }
  }

  ui.alert(`Solleciti inviati: ${inviati}`);
}


// ── RIEPILOGO ────────────────────────────────────────────────

function mostraRiepilogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.FOGLIO_INVITATI);
  const dati = sheet.getDataRange().getValues();

  let totale = 0, invitati = 0, confermati = 0, rifiutati = 0, inAttesa = 0, ospiti = 0;

  for (let i = 1; i < dati.length; i++) {
    const stato = (dati[i][5] || '').toString().trim();
    totale++;
    if (stato.includes('Confermato')) { confermati++; ospiti += parseInt(dati[i][4]) || 1; }
    else if (stato.includes('Rifiutato')) rifiutati++;
    else if (stato === 'Invitato') inAttesa++;
    else invitati++;
  }

  SpreadsheetApp.getUi().alert(
    `📊 Riepilogo Matrimonio\n\n` +
    `Totale lista: ${totale}\n` +
    `Da invitare: ${invitati}\n` +
    `Invitati (in attesa): ${inAttesa}\n` +
    `✅ Confermati: ${confermati} (≈${ospiti} ospiti)\n` +
    `❌ Rifiutati: ${rifiutati}`
  );
}


// ── UTILITY ──────────────────────────────────────────────────

function aggiornaStatoInvitato(ss, email, nuovoStato) {
  const sheet = ss.getSheetByName(CONFIG.FOGLIO_INVITATI);
  const dati = sheet.getDataRange().getValues();
  const emailLow = email.toLowerCase().trim();

  for (let i = 1; i < dati.length; i++) {
    if ((dati[i][2] || '').toString().toLowerCase().trim() === emailLow) {
      sheet.getRange(i + 1, 6).setValue(nuovoStato);       // Col F — Stato
      sheet.getRange(i + 1, 8).setValue(new Date());        // Col H — Data Risposta
      break;
    }
  }
}


// ── TEMPLATE EMAIL INVITO ────────────────────────────────────

function creaEmailInvito(nome, cognome, link) {
  const nomeCompleto = [nome, cognome].filter(Boolean).join(' ');
  const saluto = nomeCompleto ? `Caro/a ${nomeCompleto}` : 'Cari amici';

  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#FAFAF7;font-family:Georgia,'Times New Roman',serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#FAFAF7;padding:40px 20px;">
    <tr><td align="center">
      <table width="560" cellpadding="0" cellspacing="0" style="background:#FFFFFF;border-radius:4px;overflow:hidden;">
        
        <!-- Header -->
        <tr><td style="background:#7D8B6A;padding:48px 40px;text-align:center;">
          <p style="margin:0;color:rgba(255,255,255,.7);font-size:13px;letter-spacing:2px;">VI INVITIAMO A CELEBRARE</p>
          <h1 style="margin:16px 0 0;color:#FFFFFF;font-size:36px;font-weight:300;line-height:1.2;">${CONFIG.NOME_SPOSI}</h1>
        </td></tr>

        <!-- Body -->
        <tr><td style="padding:40px;">
          <p style="margin:0 0 20px;color:#2E2E2C;font-size:16px;line-height:1.7;">
            ${saluto},
          </p>
          <p style="margin:0 0 20px;color:#2E2E2C;font-size:16px;line-height:1.7;">
            con immensa gioia vi annunciamo il nostro matrimonio che si terrà il 
            <strong>${CONFIG.DATA_MATRIMONIO}</strong> presso ${CONFIG.LUOGO_CERIMONIA}, 
            con inizio alle ore ${CONFIG.ORA_CERIMONIA}.
          </p>
          <p style="margin:0 0 30px;color:#2E2E2C;font-size:16px;line-height:1.7;">
            La vostra presenza renderebbe questo giorno ancora più speciale.
            Sul nostro sito troverete tutti i dettagli e potrete confermare la vostra partecipazione.
          </p>

          <!-- CTA -->
          <p style="text-align:center;margin:0 0 30px;">
            <a href="${link}" style="display:inline-block;background:#7D8B6A;color:#FFFFFF;text-decoration:none;padding:14px 36px;border-radius:3px;font-family:Arial,sans-serif;font-size:14px;letter-spacing:1px;">SCOPRI I DETTAGLI E CONFERMA</a>
          </p>

          <p style="margin:0;color:#6B6B66;font-size:14px;line-height:1.7;text-align:center;">
            Con affetto,<br>
            <em>${CONFIG.NOME_SPOSI}</em>
          </p>
        </td></tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;
}


// ── TEMPLATE EMAIL SOLLECITO ─────────────────────────────────

function creaEmailSollecito(nome, cognome, link) {
  const nomeCompleto = [nome, cognome].filter(Boolean).join(' ');
  const saluto = nomeCompleto ? `Caro/a ${nomeCompleto}` : 'Cari amici';

  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#FAFAF7;font-family:Georgia,'Times New Roman',serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#FAFAF7;padding:40px 20px;">
    <tr><td align="center">
      <table width="560" cellpadding="0" cellspacing="0" style="background:#FFFFFF;border-radius:4px;overflow:hidden;">
        
        <tr><td style="background:#C4A265;padding:36px 40px;text-align:center;">
          <h1 style="margin:0;color:#FFFFFF;font-size:28px;font-weight:300;">${CONFIG.NOME_SPOSI}</h1>
          <p style="margin:8px 0 0;color:rgba(255,255,255,.8);font-size:13px;letter-spacing:1px;">PROMEMORIA</p>
        </td></tr>

        <tr><td style="padding:40px;">
          <p style="margin:0 0 20px;color:#2E2E2C;font-size:16px;line-height:1.7;">
            ${saluto},
          </p>
          <p style="margin:0 0 20px;color:#2E2E2C;font-size:16px;line-height:1.7;">
            non abbiamo ancora ricevuto la conferma della vostra presenza al nostro matrimonio 
            del <strong>${CONFIG.DATA_MATRIMONIO}</strong>.
          </p>
          <p style="margin:0 0 30px;color:#2E2E2C;font-size:16px;line-height:1.7;">
            Vi preghiamo di farci sapere se potrete essere dei nostri, così da organizzare 
            al meglio la giornata.
          </p>

          <p style="text-align:center;margin:0 0 30px;">
            <a href="${link}" style="display:inline-block;background:#C4A265;color:#FFFFFF;text-decoration:none;padding:14px 36px;border-radius:3px;font-family:Arial,sans-serif;font-size:14px;letter-spacing:1px;">CONFERMA LA TUA PRESENZA</a>
          </p>

          <p style="margin:0;color:#6B6B66;font-size:14px;line-height:1.7;text-align:center;">
            Grazie di cuore,<br>
            <em>${CONFIG.NOME_SPOSI}</em>
          </p>
        </td></tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;
}
