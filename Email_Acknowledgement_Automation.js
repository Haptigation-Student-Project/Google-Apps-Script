/**
 * Auto-Responder für eingehende E-Mails
 * Versendet eine E-Mails an alle ungelesenen Mails im Posteingang die nicht das Label 'AutoResponded' haben.
 * Nutzt dafür den ersten E-Mail Draft der im Betreff "E-Mail Response Automation Draft - DO NOT DELETE" stehen hat
 */
// ==================== KONFIGURATION ====================
const CONFIG = {
  // Drafts
  DRAFT_SUBJECT: 'E-Mail Response Automation Draft - DO NOT DELETE',
  FEEDBACK_DRAFT_SUBJECT: 'Feedback E-Mail Response Automation Draft - DO NOT DELETE',
  // Read Subject (Prefixes)
  TEST_SUBJECT_PREFIX: 'Test: ',
  FEEDBACK_SUBJECT_PREFIX: 'Feedback zum App-Design',
  // Send Subject
  RESPONSE_SUBJECT: 'Re: ',
  FEEDBACK_REPLY_SUBJECT: 'Vielen Dank für Ihr Feedback',
  // Other Settings
  MAX_EMAILS_PER_RUN: 50,
  CHECK_INTERVAL_MINUTES: 1,
  LABEL_NAME: 'AutoResponded',
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

/**
 * Prüft ob eine Nachricht eine Emoji-Reaktion ist
 * Emoji-Reaktionen haben typischerweise:
 * - Sehr kurzen oder leeren Body
 * - Bestimmte Header-Eigenschaften
 * - Sehr kleine Größe
 */
function isEmojiReaction(message) {
  try {
    const plainBody = message.getPlainBody().trim();
    const subject = message.getSubject();
    
    // Emoji-Reaktionen haben meist einen leeren oder sehr kurzen Body
    // und enthalten oft spezielle Marker im Body oder Subject
    if (plainBody.length === 0) {
      return true;
    }
    
    // Prüfe auf typische Emoji-Reaktions-Muster
    // Gmail-Reaktionen enthalten oft nur ein einzelnes Emoji-Zeichen
    const emojiPattern = /^[\u{1F300}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}]+$/u;
    if (plainBody.length <= 10 && emojiPattern.test(plainBody)) {
      return true;
    }
    
    // Zusätzliche Prüfung: Sehr kurze Nachrichten mit nur Emojis oder Leerzeichen
    if (plainBody.length <= 5 && plainBody.replace(/\s/g, '').length <= 2) {
      return true;
    }
    
    return false;
  } catch (error) {
    Logger.log(`Fehler bei Emoji-Check: ${error.toString()}`);
    return false;
  }
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
  const searchQuery = `is:unread -label:${CONFIG.LABEL_NAME}`;
  const threads = GmailApp.search(searchQuery, 0, CONFIG.MAX_EMAILS_PER_RUN);

  if (threads.length === 0) {
    Logger.log('Keine neuen E-Mails gefunden');
    return;
  }

  Logger.log(`${threads.length} neue E-Mail(s) gefunden`);

  // Hole die Drafts anhand der konfigurierten Betreffzeilen
  const drafts = GmailApp.getDrafts();
  let defaultDraft = null;
  let feedbackDraft = null;
  drafts.forEach(draft => {
    const subject = draft.getMessage().getSubject();
    if (subject === CONFIG.DRAFT_SUBJECT) {
      defaultDraft = draft;
    }
    if (subject === CONFIG.FEEDBACK_DRAFT_SUBJECT) {
      feedbackDraft = draft;
    }
  });

  if (!defaultDraft) {
    Logger.log(`Kein Draft mit Betreff "${CONFIG.DRAFT_SUBJECT}" gefunden`);
    return;
  }
  
  if (!feedbackDraft) {
    Logger.log(`Kein Feedback-Draft mit Betreff "${CONFIG.FEEDBACK_DRAFT_SUBJECT}" gefunden`);
  }
  
  if (defaultDraft) {
    Logger.log(`✓ Response-Draft gefunden: "${CONFIG.DRAFT_SUBJECT}"`);
  }

  if (feedbackDraft) {
    Logger.log(`✓ Feedback-Draft gefunden: "${CONFIG.FEEDBACK_DRAFT_SUBJECT}"`);
  }

  let successCount = 0;
  let failCount = 0;
  let skippedEmojiCount = 0;

  // Verarbeite jeden Thread
  threads.forEach(thread => {
    try {
      const messages = thread.getMessages();
      const latestMessage = messages[messages.length - 1];

      // Prüfe ob die E-Mail bereits beantwortet wurde (sollte durch Search-Query bereits gefiltert sein)
      if (!latestMessage.isUnread()) {
        return;
      }

      // Prüfe ob es sich um eine Emoji-Reaktion handelt
      if (isEmojiReaction(latestMessage)) {
        skippedEmojiCount++;
        const sender = latestMessage.getFrom();
        Logger.log(`⊘ Emoji-Reaktion übersprungen von: ${anonymizeEmail(sender)}`);
        // Markiere mit Label, damit es nicht erneut geprüft wird
        thread.addLabel(label);
        return;
      }

      const sender = latestMessage.getFrom();
      const subject = latestMessage.getSubject();

      // Test-Mails komplett ignorieren
      if (subject.startsWith(CONFIG.TEST_SUBJECT_PREFIX)) {
        Logger.log(`⊘ Test-Mail ignoriert: ${subject}`);
        thread.addLabel(label);
        return;
      }

      // Draft & Betreff je nach Typ auswählen
      let activeDraft = defaultDraft;
      let replySubject = `${CONFIG.RESPONSE_SUBJECT}${subject}`;

      // Feedback-Mails mit Sonderbehandlung
      if (subject.contains(CONFIG.FEEDBACK_SUBJECT_PREFIX)) {
        if (!feedbackDraft) {
          Logger.log(`Feedback-Draft fehlt – Mail übersprungen`);
          thread.addLabel(label);
          return;
        }
        activeDraft = feedbackDraft;
        replySubject = CONFIG.FEEDBACK_REPLY_SUBJECT;
      }

      const draftMessage = activeDraft.getMessage();

      // Hole alle Anhänge (Inline-Bilder) aus dem Draft
      const attachments = draftMessage.getAttachments();
      const inlineImages = {};
      attachments.forEach((attachment, index) => {
        inlineImages[`img${index}`] = attachment;
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
  Logger.log(`⊘ Emoji-Reaktionen ignoriert: ${skippedEmojiCount}`);
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