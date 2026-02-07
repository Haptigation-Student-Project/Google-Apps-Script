/**
 * Newsletter-Versand Script für Gmail
 * Versendet individuelle E-Mails an alle Kontakte mit dem Label "Newsletter Subscriber"
 * Nutze hierzu einen Email Draft in GMAIL mit den Worten "Update zu Haptigation", irgendwo im Betreff
 */

// ========== KONFIGURATION ==========
// Hier trägst du deine Werte ein:

const CONFIG = {
  // Das Label deiner Newsletter-Kontakte in Gmail
  contactLabel: "extra",
  
  // Name des Absenders (wird als "Max Mustermann von Haptigation via haptigation@gmail.com" angezeigt)
  senderName: "David von Haptigation",
  
  // Test-Modus: Wenn true, wird nur an deine eigene E-Mail gesendet
  testMode: true,
  
  // Setze Datum und Uhrzeit für automatischen Versand
  scheduledDate: "2026-02-21", // YYYY-MM-DD z.B. "2025-12-25"
  scheduledTime: "16:00",  // HH:MM z.B. "10:00" (24h-format)
};

// ========== HAUPTFUNKTIONEN ==========

/**
 * Versendet den Newsletter sofort oder wird automatisch zum geplanten Zeitpunkt aufgerufen
 */
function sendNewsletter() {
  try {
    // Log-Header für geplante Versände
    const isScheduled = isCalledByTrigger();
    if (isScheduled) {
      Logger.log(`\n========== GEPLANTER NEWSLETTER-VERSAND ==========`);
      Logger.log(`Zeitpunkt: ${new Date().toLocaleString('de-DE')}`);
    }
    
    // 1. Newsletter-Entwurf und Kontakte validieren
    const validation = validateNewsletterPrerequisites();
    if (!validation.success) {
      return;
    }
    const { draft, contacts } = validation;
    
    // 2. Versand-Info loggen
    const numRecipients = CONFIG.testMode ? 1 : contacts.length;
    const mode = CONFIG.testMode ? "TEST-MODUS" : "LIVE-VERSAND";
    
    Logger.log(`\n========== NEWSLETTER-VERSAND ==========`);
    Logger.log(`Modus: ${mode}`);
    Logger.log(`Empfänger: ${numRecipients}`);
    Logger.log(`Betreff: ${draft.getMessage().getSubject()}`);
    Logger.log(`========================================\n`);
    
    // 3. E-Mails versenden
    sendEmails(draft, contacts);
    
    // 4. Bei geplantem Versand: Trigger automatisch löschen
    if (isScheduled) {
      deleteAllNewsletterTriggers();
    }
    
    // 5. Entwurf löschen (optional - auskommentieren wenn du ihn behalten willst)
    // draft.deleteDraft();
    
  } catch (error) {
    Logger.log("FEHLER beim Versand: " + error.toString());
    showAlert("Fehler beim Versand", error.toString());
  }
}

/**
 * Plant den Newsletter-Versand für ein bestimmtes Datum und Uhrzeit
 */
function scheduleNewsletter() {
  const dateStr = CONFIG.scheduledDate;
  const timeStr = CONFIG.scheduledTime;

  try {
    // Eingaben validieren
    if (!dateStr || !timeStr) {
      showAlert("Fehler", "Bitte Datum (YYYY-MM-DD) und Uhrzeit (HH:MM) in der CONFIG angeben.");
      return;
    }
    
    // Datum und Zeit parsen
    const dateParts = dateStr.split('-');
    const timeParts = timeStr.split(':');
    
    if (dateParts.length !== 3 || timeParts.length !== 2) {
      showAlert("Fehler", "Ungültiges Format!\n\n" +
                "Datum: YYYY-MM-DD (z.B. 2024-12-25)\n" +
                "Uhrzeit: HH:MM (z.B. 14:30)");
      return;
    }
    
    // Date-Objekt erstellen
    const scheduledDateTime = new Date(
      parseInt(dateParts[0]), // Jahr
      parseInt(dateParts[1]) - 1, // Monat (0-basiert)
      parseInt(dateParts[2]), // Tag
      parseInt(timeParts[0]), // Stunde
      parseInt(timeParts[1]), // Minute
      0 // Sekunde
    );
    
    // Prüfen ob Datum in der Zukunft liegt
    const now = new Date();
    if (scheduledDateTime <= now) {
      showAlert("Fehler", "Das gewählte Datum und die Uhrzeit liegen in der Vergangenheit!\n\n" +
                `Gewählt: ${scheduledDateTime.toLocaleString('de-DE')}\n` +
                `Jetzt: ${now.toLocaleString('de-DE')}`);
      return;
    }
    
    // Prüfen ob Newsletter-Entwurf und Kontakte existieren
    const validation = validateNewsletterPrerequisites();
    if (!validation.success) {
      return;
    }
    const { draft, contacts } = validation;
    
    // Alte Trigger löschen (falls vorhanden)
    deleteAllNewsletterTriggers();
    
    // Neuen Trigger erstellen - ruft direkt sendNewsletter() auf
    ScriptApp.newTrigger('sendNewsletter')
      .timeBased()
      .at(scheduledDateTime)
      .create();
    
    // Bestätigung
    const mode = CONFIG.testMode ? " (TEST-MODUS)" : "";
    const numRecipients = CONFIG.testMode ? 1 : contacts.length;
    
    showAlert("Newsletter geplant!", 
      `Der Newsletter wird automatisch versendet${mode}:\n\n` +
      `📅 Datum: ${scheduledDateTime.toLocaleDateString('de-DE')}\n` +
      `🕐 Uhrzeit: ${scheduledDateTime.toLocaleTimeString('de-DE', {hour: '2-digit', minute: '2-digit'})}\n` +
      `📧 Empfänger: ${numRecipients}\n` +
      `📝 Betreff: ${draft.getMessage().getSubject()}\n\n` +
      (CONFIG.testMode ? "⚠️ Test-Modus ist aktiv - E-Mail wird nur an dich gesendet!\n\n" : "") +
      `Du kannst den geplanten Versand jederzeit mit deleteAllNewsletterTriggers() abbrechen.`);
    
    Logger.log(`Newsletter geplant für: ${scheduledDateTime.toLocaleString('de-DE')}, Empfänger: ${numRecipients}`);
    
  } catch (error) {
    Logger.log("Fehler beim Planen: " + error.toString());
    showAlert("Fehler beim Planen", error.toString());
  }
}

/**
 * Zeigt Informationen über geplante Newsletter-Versände an
 */
function showScheduledNewsletters() {
  const triggers = ScriptApp.getProjectTriggers();
  const newsletterTriggers = triggers.filter(t => 
    t.getHandlerFunction() === 'sendNewsletter' && t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
  );
  
  if (newsletterTriggers.length === 0) {
    showAlert("Keine geplanten Versände", 
      "Es sind aktuell keine Newsletter-Versände geplant.\n\n" +
      "Nutze scheduleNewsletter() um einen Versand zu planen.");
    return;
  }
  
  let message = `Geplante Newsletter-Versände (${newsletterTriggers.length}):\n\n`;
  
  newsletterTriggers.forEach((trigger, index) => {
    const triggerDate = new Date(trigger.getTriggerSource());
    message += `${index + 1}. 📅 ${triggerDate.toLocaleDateString('de-DE')} ` +
               `🕐 ${triggerDate.toLocaleTimeString('de-DE', {hour: '2-digit', minute: '2-digit'})}\n`;
  });
  
  message += `\n💡 Nutze deleteAllNewsletterTriggers() zum Abbrechen.`;
  
  showAlert("Geplante Versände", message);
}

/**
 * Löscht alle Newsletter-Trigger (geplante Versände)
 * @returns {number} Anzahl der gelöschten Trigger
 */
function deleteAllNewsletterTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  
  triggers.forEach(trigger => {
    // Nur zeitbasierte Trigger für sendNewsletter löschen
    if (trigger.getHandlerFunction() === 'sendNewsletter' && 
        trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });
  
  if (deleted > 0) {
    Logger.log(`${deleted} Newsletter-Trigger gelöscht`);
  }
  
  return deleted;
}

// ========== HILFSFUNKTIONEN ==========

/**
 * Prüft ob die Funktion von einem Trigger aufgerufen wurde
 * @returns {boolean} true wenn durch Trigger aufgerufen
 */
function isCalledByTrigger() {
  try {
    // Wenn es einen aktiven Trigger gibt, wurde die Funktion durch einen Trigger aufgerufen
    const triggers = ScriptApp.getProjectTriggers();
    return triggers.some(t => 
      t.getHandlerFunction() === 'sendNewsletter' && 
      t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
    );
  } catch (e) {
    return false;
  }
}

/**
 * Validiert ob Newsletter-Entwurf und Kontakte vorhanden sind
 * @returns {Object} {success: boolean, draft?: GmailDraft, contacts?: Array, error?: string}
 */
function validateNewsletterPrerequisites() {
  // 1. Newsletter-Entwurf prüfen
  const draftResult = getNewsletterDraft();
  
  if (!draftResult.success) {
    showAlert("Fehler: Newsletter-Entwurf", draftResult.error);
    return { success: false, error: draftResult.error };
  }
  
  // 2. Kontakte laden und prüfen
  const contacts = getNewsletterContacts();
  
  if (contacts.length === 0) {
    const error = `Keine Kontakte mit dem Label "${CONFIG.contactLabel}" gefunden.\n\n` +
                  `Stelle sicher, dass:\n` +
                  `1. Du Kontakte in Gmail hast\n` +
                  `2. Diese Kontakte das Label "${CONFIG.contactLabel}" haben\n` +
                  `3. Die People API aktiviert ist`;
    showAlert("Keine Kontakte gefunden", error);
    return { success: false, error };
  }
  
  Logger.log(`✓ Validierung erfolgreich: 1 Entwurf, ${contacts.length} Kontakte`);
  return { 
    success: true, 
    draft: draftResult.draft, 
    contacts 
  };
}

/**
 * Lädt den Newsletter-Entwurf aus Gmail
 * @returns {Object} {success: boolean, draft?: GmailDraft, error?: string}
 */
function getNewsletterDraft() {
  const drafts = GmailApp.getDrafts();
  
  if (drafts.length === 0) {
    return {
      success: false,
      error: 'Keine E-Mail-Entwürfe in Gmail gefunden.\n\n' +
             'Bitte erstelle einen Entwurf mit "Update zu Haptigation" im Betreff.'
    };
  }
  
  // Sammle alle passenden Entwürfe
  const matchingDrafts = [];
  
  for (let draft of drafts) {
    const subject = draft.getMessage().getSubject();
    if (subject.includes("Update zu Haptigation")) {
      matchingDrafts.push({
        draft: draft,
        subject: subject
      });
    }
  }
  
  // Prüfen ob mehrere Entwürfe gefunden wurden
  if (matchingDrafts.length > 1) {
    const subjects = matchingDrafts.map((d, i) => `${i + 1}. "${d.subject}"`).join('\n');
    return {
      success: false,
      error: `${matchingDrafts.length} Newsletter-Entwürfe gefunden!\n\n` +
             `Bitte lösche alle bis auf einen:\n${subjects}`
    };
  }
  
  // Keinen passenden Entwurf gefunden
  if (matchingDrafts.length === 0) {
    Logger.log("Verfügbare Entwürfe:");
    drafts.forEach(d => {
      Logger.log(`- "${d.getMessage().getSubject()}"`);
    });
    
    return {
      success: false,
      error: 'Kein Entwurf mit "Update zu Haptigation" im Betreff gefunden.\n\n' +
             'Bitte erstelle einen entsprechenden Entwurf in Gmail.'
    };
  }
  
  // Genau ein Entwurf gefunden
  Logger.log(`✓ Newsletter-Entwurf gefunden: "${matchingDrafts[0].subject}"`);
  return {
    success: true,
    draft: matchingDrafts[0].draft
  };
}

/**
 * Lädt alle Kontakte mit dem Newsletter-Label über die People API
 */
function getNewsletterContacts() {
  const contacts = [];
  
  try {
    // 1. Alle Contact Groups laden
    const contactGroups = People.ContactGroups.list();
    
    if (!contactGroups.contactGroups) {
      Logger.log("Keine Contact Groups gefunden");
      return contacts;
    }
    
    // 2. Die richtige Group finden
    let targetGroupResourceName = null;
    
    for (let group of contactGroups.contactGroups) {
      if (group.name === CONFIG.contactLabel) {
        targetGroupResourceName = group.resourceName;
        break;
      }
    }
    
    if (!targetGroupResourceName) {
      Logger.log(`Keine Contact Group mit Namen "${CONFIG.contactLabel}" gefunden`);
      Logger.log("Verfügbare Groups:");
      contactGroups.contactGroups.forEach(g => Logger.log(`- ${g.name}`));
      return contacts;
    }
    
    // 3. Kontakte der Gruppe direkt laden
    const response = People.ContactGroups.get(targetGroupResourceName, {
      maxMembers: 500
    });
    
    if (!response.memberResourceNames || response.memberResourceNames.length === 0) {
      Logger.log("Keine Mitglieder in dieser Group gefunden");
      return contacts;
    }
    
    // 4. Details der Mitglieder laden (in Batches von 50)
    const memberResourceNames = response.memberResourceNames;
    const batchSize = 50;
    
    Logger.log(`Lade ${memberResourceNames.length} Kontakte aus "${CONFIG.contactLabel}"...`);
    
    for (let i = 0; i < memberResourceNames.length; i += batchSize) {
      const batch = memberResourceNames.slice(i, i + batchSize);
      
      try {
        const batchResponse = People.People.getBatchGet({
          resourceNames: batch,
          personFields: 'names,emailAddresses'
        });
        
        if (batchResponse.responses) {
          batchResponse.responses.forEach(personResponse => {
            if (personResponse.person) {
              const person = personResponse.person;
              
              // Prüfen ob E-Mail existiert und valide ist
              if (!person.emailAddresses || person.emailAddresses.length === 0) {
                return; // Kontakt überspringen
              }
              
              const email = person.emailAddresses[0].value;
              
              if (!email || !email.includes('@')) {
                Logger.log(`⚠ Ungültige E-Mail übersprungen: ${email}`);
                return;
              }
              
              let name = email.split('@')[0]; // Fallback
              
              if (person.names && person.names.length > 0) {
                name = person.names[0].displayName || name;
              }
              
              contacts.push({
                email: email,
                name: name
              });
            }
          });
        }
      } catch (batchError) {
        Logger.log(`Fehler beim Laden eines Batches: ${batchError.toString()}`);
      }
    }
    
    Logger.log(`✓ ${contacts.length} Kontakte erfolgreich geladen`);
    
  } catch (error) {
    Logger.log("Fehler beim Laden der Kontakte: " + error.toString());
    throw new Error(`Kontakte konnten nicht geladen werden: ${error.toString()}\n\n` +
                   `Mögliche Ursachen:\n` +
                   `1. People API nicht aktiviert (Services > + Add service > People API)\n` +
                   `2. Keine Kontakte in der Gruppe\n` +
                   `3. Fehlende Berechtigungen`);
  }
  
  return contacts;
}

/**
 * Versendet E-Mails an alle Kontakte
 */
function sendEmails(draft, contacts) {
  const message = draft.getMessage();
  const subject = message.getSubject();
  const htmlBody = message.getBody();
  const plainBody = message.getPlainBody();
  
  const results = {
    success: 0,
    failed: 0,
    errors: []
  };
  
  // Im Test-Modus nur an dich selbst
  const recipients = CONFIG.testMode 
    ? [{ email: Session.getActiveUser().getEmail(), name: "Test" }]
    : contacts;
  
  recipients.forEach((contact, index) => {
    try {
      // Fallback für unbekannte Namen
      let displayName = contact.name;
      if (!displayName || displayName.toLowerCase() === 'unbekannt' || displayName.trim() === '') {
        displayName = 'LeserIn';
      }
      
      // Personalisierung - unterstützt [NAME], [Name] und [name]
      let personalizedSubject = subject
        .replace(/\[NAME\]/g, displayName)
        .replace(/\[Name\]/g, displayName)
        .replace(/\[name\]/g, displayName);

      // Testmodus-Präfix für Betreff
      if (CONFIG.testMode) {
        personalizedSubject = `[TEST] ${personalizedSubject}`;
      }
      
      const personalizedHtml = htmlBody
        .replace(/\[NAME\]/g, displayName)
        .replace(/\[Name\]/g, displayName)
        .replace(/\[name\]/g, displayName);
      
      const personalizedPlain = plainBody
        .replace(/\[NAME\]/g, displayName)
        .replace(/\[Name\]/g, displayName)
        .replace(/\[name\]/g, displayName);
      
      // E-Mail versenden
      GmailApp.sendEmail(
        contact.email,
        personalizedSubject,
        personalizedPlain,
        {
          htmlBody: personalizedHtml,
          name: CONFIG.senderName,
          noReply: false,
          charset: 'UTF-8'
        }
      );
      
      results.success++;
      
      if (index < 3 || index === recipients.length - 1) {
        Logger.log(`✓ E-Mail ${index + 1}/${recipients.length} gesendet an: ${anonymizeEmail(contact.email)}`);
      }
      
      // Kurze Pause alle 10 E-Mails, um Gmail-Limits zu vermeiden
      if ((index + 1) % 10 === 0) {
        Utilities.sleep(1000);
      }
      
    } catch (error) {
      results.failed++;
      const anonymizedEmail = anonymizeEmail(contact.email);
      results.errors.push(`${anonymizedEmail}: ${error.toString()}`);
      Logger.log(`✗ Fehler bei ${anonymizedEmail}: ${error.toString()}`);
    }
  });
  
  return results;
}

/**
 * Zeigt eine Nachricht an
 */
function showAlert(title, message) {
  const fullMessage = `${title}\n\n${message}`;
  
  // Versuche UI-Dialog, falls verfügbar
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    // Kein UI verfügbar - nur loggen
    Logger.log(fullMessage);
  }
}

/**
 * Zeigt alle Newsletter-Kontakte an
 */
function showContacts() {
  const contacts = getNewsletterContacts();
  
  if (contacts.length === 0) {
    showAlert("Keine Kontakte", `Keine Kontakte mit Label "${CONFIG.contactLabel}" gefunden.`);
    return;
  }
  
  let message = `Gefundene Kontakte (${contacts.length}):\n\n`;
  contacts.slice(0, 20).forEach(c => {
    message += `• ${c.name} (${anonymizeEmail(c.email)})\n`;
  });
  
  if (contacts.length > 20) {
    message += `\n... und ${contacts.length - 20} weitere`;
  }
  
  showAlert("Newsletter-Kontakte", message);
}

/**
 * Listet alle verfügbaren Contact Groups auf
 */
function listAllContactGroups() {
  try {
    const contactGroups = People.ContactGroups.list();
    
    if (!contactGroups.contactGroups) {
      showAlert("Keine Groups", "Keine Contact Groups gefunden");
      return;
    }
    
    let message = `Verfügbare Contact Groups (${contactGroups.contactGroups.length}):\n\n`;
    contactGroups.contactGroups.forEach(group => {
      message += `• ${group.name}\n`;
    });
    
    showAlert("Contact Groups", message);
    
  } catch (error) {
    showAlert("Fehler", `Fehler beim Laden der Groups: ${error.toString()}`);
  }
}

/**
 * Anonymisiert eine E-Mail-Adresse für Logs
 * Beispiel: max.mustermann@gmail.com -> m...@g...com
 */
function anonymizeEmail(email) {
  // Prüfen ob email existiert
  if (!email || typeof email !== 'string') {
    return '[keine E-Mail]';
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
