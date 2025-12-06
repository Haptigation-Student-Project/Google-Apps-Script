/**
 * Auto-Responder für eingehende E-Mails
 * Versendet eine E-Mails an alle ungelesenen Mails im Posteingang die nicht das Label 'AutoResponded' haben.
 * Nutzt dafür den ersten E-Mail Draft der im Betreff "E-Mail Response Automation Draft - DO NOT DELETE" stehen hat
 */
// ==================== KONFIGURATION ====================
const CONFIG = {
  DRAFT_SUBJECT: 'E-Mail Response Automation Draft - DO NOT DELETE', // Betreff des Draft-Entwurfs
  REPLY_SUBJECT: 'Re: ',                                             // Betreff für die Antwort (z.B. 'Vielen Dank für Ihre E-Mail!')
  MAX_EMAILS_PER_RUN: 50,                                            // Max. Anzahl E-Mails pro Durchlauf
  CHECK_INTERVAL_MINUTES: 1,                                         // Prüfintervall in Minuten
  LABEL_NAME: 'AutoResponded',                                       // Label für bereits beantwortete E-Mails
  
  // Debug-Modus: Wenn true, werden E-Mails in Logs NICHT anonymisiert
  debugMode: false
};
// =======================================================

// ========== HILFSFUNKTION FÜR EMAIL-ANONYMISIERUNG ==========

/**
 * Anonymisiert eine E-Mail-Adresse für Logs
 * Beispiel: max.mustermann@gmail.com -> m...@g...com
 */
function anonymizeEmail(email) {
  // Wenn Debug-Modus aktiv ist, E-Mail nicht anonymisieren
  if (CONFIG.debugMode) {
    return email;
  }
  
  // E-Mail in Bestandteile zerlegen
  const parts = email.split('@');
  if (parts.length !== 2) {
    return email; // Ungültiges Format, nicht anonymisieren
  }
  
  const localPart = parts[0];
  const domain = parts[1];
  
  // Domain in Name und TLD zerlegen
  const domainParts = domain.split('.');
  
  // Lokalen Teil anonymisieren (nur ersten Buchstaben zeigen)
  const anonymizedLocal = localPart.charAt(0) + '...';
  
  // Domain anonymisieren (nur ersten Buchstaben des Domain-Namens zeigen)
  const anonymizedDomain = domainParts.map((part, index) => {
    if (index === domainParts.length - 1) {
      // TLD vollständig zeigen (z.B. "de", "com")
      return part;
    }
    return part.charAt(0) + '...';
  }).join('.');
  
  return `${anonymizedLocal}@${anonymizedDomain}`;
}

function autoResponseToEmails() {
  const debugInfo = CONFIG.debugMode ? ' | DEBUG-MODUS AKTIV' : '';
  Logger.log(`========== AUTO-RESPONDER PRÜFUNG ==========`);
  Logger.log(`Modus: ${debugInfo}`);
  Logger.log(`===========================================\n`);
  
  // Erstelle oder hole das Label für bereits beantwortete E-Mails
  let label = GmailApp.getUserLabelByName(CONFIG.LABEL_NAME);
  if (!label) {
    label = GmailApp.createLabel(CONFIG.LABEL_NAME);
    Logger.log(`✓ Label "${CONFIG.LABEL_NAME}" erstellt`);
  }
  
  // Suche nach ungelesenen E-Mails, die noch NICHT das Label haben
  const searchQuery = 'is:unread -label:' + CONFIG.LABEL_NAME;
  const threads = GmailApp.search(searchQuery, 0, CONFIG.MAX_EMAILS_PER_RUN);
  
  if (threads.length === 0) {
    Logger.log('Keine neuen E-Mails gefunden');
    return;
  }
  
  Logger.log(`${threads.length} neue E-Mail(s) gefunden`);
  
  // Hole den Draft mit dem Betreff "E-Mail response"
  const drafts = GmailApp.getDrafts();
  let responseDraft = null;
  
  for (let i = 0; i < drafts.length; i++) {
    if (drafts[i].getMessage().getSubject() === CONFIG.DRAFT_SUBJECT) {
      responseDraft = drafts[i];
      break;
    }
  }
  
  if (!responseDraft) {
    Logger.log('Kein Draft mit dem Betreff "' + CONFIG.DRAFT_SUBJECT + '" gefunden');
    return;
  }
  
  Logger.log(`✓ Response-Draft gefunden: "${CONFIG.DRAFT_SUBJECT}"`);
  
  let successCount = 0;
  let failCount = 0;
  
  // Verarbeite jeden Thread
  threads.forEach(thread => {
    try {
      const messages = thread.getMessages();
      const latestMessage = messages[messages.length - 1];
      
      // Prüfe ob die E-Mail bereits beantwortet wurde (sollte durch Search-Query bereits gefiltert sein)
      if (!latestMessage.isUnread()) {
        return;
      }
      
      const sender = latestMessage.getFrom();
      const subject = latestMessage.getSubject();
      const draftMessage = responseDraft.getMessage();
      
      // Erstelle den Antwort-Betreff
      const replySubject = CONFIG.REPLY_SUBJECT.includes('Re:') 
        ? CONFIG.REPLY_SUBJECT + subject 
        : CONFIG.REPLY_SUBJECT;
      
      // Hole alle Anhänge (Inline-Bilder) aus dem Draft
      const attachments = draftMessage.getAttachments();
      const inlineImages = {};
      
      // Erstelle inlineImages Objekt für eingebettete Bilder
      attachments.forEach((attachment, index) => {
        inlineImages['img' + index] = attachment;
      });
      
      // Sende die Antwort mit Inline-Bildern (ohne attachments Parameter)
      GmailApp.sendEmail(
        sender,
        replySubject,
        draftMessage.getPlainBody(),
        {
          htmlBody: draftMessage.getBody(),
          inlineImages: inlineImages,
          replyTo: Session.getActiveUser().getEmail()
        }
      );
      
      // Markiere den Thread mit Label (damit keine Duplikate gesendet werden)
      thread.addLabel(label);
      
      // Optional: Markiere die E-Mail als gelesen (kannst du auskommentieren wenn nicht gewünscht)
      // latestMessage.markRead();
      
      successCount++;
      Logger.log(`✓ Auto-Response gesendet an: ${anonymizeEmail(sender)}`);
      
    } catch (error) {
      failCount++;
      const sender = thread.getMessages()[thread.getMessages().length - 1].getFrom();
      Logger.log(`✗ Fehler bei ${anonymizeEmail(sender)}: ${error.toString()}`);
    }
  });
  
  Logger.log(`\n========== ZUSAMMENFASSUNG ==========`);
  Logger.log(`✓ Erfolgreich: ${successCount}`);
  Logger.log(`✗ Fehlgeschlagen: ${failCount}`);
  Logger.log(`=====================================\n`);
}

// Funktion zum Einrichten des Triggers
function setupTrigger() {
  // Lösche bestehende Trigger für diese Funktion
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoResponseToEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Erstelle neuen Trigger (läuft alle X Minuten)
  ScriptApp.newTrigger('autoResponseToEmails')
    .timeBased()
    .everyMinutes(CONFIG.CHECK_INTERVAL_MINUTES)
    .create();
  
  Logger.log('Trigger erfolgreich eingerichtet (alle ' + CONFIG.CHECK_INTERVAL_MINUTES + ' Minuten)');
}

// Funktion zum Entfernen des Triggers
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoResponseToEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  Logger.log('Trigger entfernt');
}
