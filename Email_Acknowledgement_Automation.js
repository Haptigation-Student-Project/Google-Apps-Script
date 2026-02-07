/**
 * Auto-Responder für eingehende E-Mails
 * Versendet eine E-Mails an alle ungelesenen Mails im Posteingang die nicht das Label 'AutoResponded' haben.
 * Nutzt dafür den ersten E-Mail Draft der im Betreff "E-Mail Response Automation Draft - DO NOT DELETE" stehen hat
 */
// ==================== KONFIGURATION ====================
const CONFIG = {
  DRAFT_SUBJECT: 'E-Mail Response Automation Draft - DO NOT DELETE', // Betreff des Drafts für allgemeine Antworten
  FEEDBACK_DRAFT_SUBJECT: 'Feedback E-Mail Response Automation Draft - DO NOT DELETE', // Betreff des Drafts für Feedback-Antworten

  TEST_SUBJECT_PREFIXES: ['[TEST]', '[Test]', '[test]', 'TEST:', 'Test:', 'test:'], // Präfixe für Test-Mails werden von Auto Responder ignoriert

  FEEDBACK_SUBJECT_PREFIX: 'Feedback zum App-Design', // Betreff der Feedback Emails die eine spezielle Antwort erhalten sollen
  FEEDBACK_REPLY_SUBJECT: 'Vielen Dank für Ihr Feedback', // Betreff der Antwort auf Feedback-Mails

  MAX_EMAILS_PER_RUN: 50, // Maximale Anzahl von E-Mails, die pro Ausführung verarbeitet werden (um Limits zu vermeiden)
  CHECK_INTERVAL_MINUTES: 1, // Intervall in Minuten, in dem nach neuen E-Mails gesucht wird (z.B. alle 5 Minuten)
  LABEL_NAME: 'AutoResponded' // Name des Labels, das auf E-Mails angewendet wird, die bereits beantwortet wurden (um Duplikate zu vermeiden)
};
// =======================================================

// ========== HAUPTFUNKTIONEN ===========
// Funktion zum Einrichten des Triggers
function setupTrigger() {
  // Lösche bestehende Trigger für diese Funktion
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoResponseToEmails') {
      Logger.log('Alter Trigger gelöscht.')
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

function autoResponseToEmails() {
  Logger.log(`========== AUTO-RESPONDER PRÜFUNG ==========\n`);
  
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

  Logger.log(`✓ Response-Draft gefunden: "${CONFIG.DRAFT_SUBJECT}"`);
  
  // Hole eigene E-Mail-Adresse für Selbst-Antwort-Check
  const myEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  let successCount = 0;
  let failCount = 0;
  let skippedEmojiCount = 0;
  let skippedTestCount = 0;
  let skippedSelfCount = 0;
  
  // Verarbeite jeden Thread
  threads.forEach(thread => {
    try {
      const messages = thread.getMessages();
      const latestMessage = messages[messages.length - 1];
      
      // Prüfe ob die E-Mail bereits beantwortet wurde (sollte durch Search-Query bereits gefiltert sein)
      if (!latestMessage.isUnread()) {
        return;
      }
      
      // NEU: Prüfe ob es sich um eine Emoji-Reaktion handelt
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

      // 1. Test-Mails ignorieren (erweiterte Liste)
      if (isTestEmail(subject)) {
        skippedTestCount++;
        Logger.log(`⊘ Test-Mail ignoriert: ${subject}`);
        thread.addLabel(label);
        return;
      }

      // 2. Niemals auf eigene E-Mails antworten
      const senderEmail = extractEmailAddress(sender).toLowerCase();
      if (senderEmail === myEmail) {
        skippedSelfCount++;
        Logger.log(`⊘ Eigene E-Mail ignoriert: ${anonymizeEmail(sender)}`);
        thread.addLabel(label);
        return;
      }

      // Draft & Betreff je nach Typ auswählen
      let activeDraft = defaultDraft;
      let replySubject = subject ? 'Re: ' + subject.trim() : 'Re: ';

      // Feedback-Mails mit Sonderbehandlung
      if (subject.startsWith(CONFIG.FEEDBACK_SUBJECT_PREFIX)) {
        if (!feedbackDraft) {
          Logger.log('Feedback-Draft fehlt – Mail übersprungen');
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
  Logger.log(`⊘ Emoji-Reaktionen ignoriert: ${skippedEmojiCount}`);
  Logger.log(`⊘ Test-Mails ignoriert: ${skippedTestCount}`);
  Logger.log(`⊘ Eigene E-Mails ignoriert: ${skippedSelfCount}`);
  Logger.log(`✗ Fehlgeschlagen: ${failCount}`);
  Logger.log(`=====================================\n`);
}

// ========== HILFSFUNKTIONEN ==========

/**
 * Prüft ob eine E-Mail eine Test-Mail ist
 * Unterstützt: [TEST], [Test], [test], TEST:, Test:, test:
 */
function isTestEmail(subject) {
  if (!subject) return false;
  
  const trimmedSubject = subject.trim();
  
  // Prüfe alle konfigurierten Test-Präfixe
  for (const prefix of CONFIG.TEST_SUBJECT_PREFIXES) {
    if (trimmedSubject.startsWith(prefix)) {
      return true;
    }
  }
  
  return false;
}

/**
 * Extrahiert die reine E-Mail-Adresse aus dem From-Feld
 * Beispiel: "Max Mustermann <max@example.com>" -> "max@example.com"
 */
function extractEmailAddress(fromField) {
  if (!fromField) return '';
  
  // Prüfe ob E-Mail in spitzen Klammern steht
  const matches = fromField.match(/<([^>]+)>/);
  if (matches && matches[1]) {
    return matches[1].trim();
  }
  
  // Fallback: Nutze das ganze Feld (falls keine Klammern vorhanden)
  return fromField.trim();
}

/**
 * Anonymisiert eine E-Mail-Adresse für Logs
 * Beispiel: max.mustermann@gmail.com -> m...@g...com
 */
function anonymizeEmail(email) {
  // Extrahiere zuerst die reine E-Mail-Adresse
  const cleanEmail = extractEmailAddress(email);
  
  // E-Mail in Bestandteile zerlegen
  const parts = cleanEmail.split('@');
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
