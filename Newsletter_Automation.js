/**
 * Newsletter-Versand Script für Gmail
 * Versendet individuelle E-Mails an alle Kontakte mit dem Label "Newsletter Subscriber"
 * Nutzt dafür den ersten E-Mail Draft der irgendwo im Betreff "Update zu Haptigation" stehen hat
 */

// ========== KONFIGURATION ==========
// Hier trägst du deine Werte ein:

const CONFIG = {
  // Das Label deiner Newsletter-Kontakte in Gmail
  contactLabel: "safetyFirst",
  
  // Name des Absenders (wird als "Max Mustermann von Haptigation via haptigation@gmail.com" angezeigt)
  senderName: "David von Haptigation",
  
  // Test-Modus: Wenn true, wird nur an deine eigene E-Mail gesendet
  testMode: true,
  
  // Debug-Modus: Wenn true, werden E-Mails in Logs NICHT anonymisiert
  debugMode: false,
};

// ========== HAUPTFUNKTION ==========

/**
 * Diese Funktion rufst du auf, um den Newsletter zu versenden
 */
function sendNewsletter() {
  try {
    // 1. Newsletter-Entwurf laden
    const draft = getNewsletterDraft();
    if (!draft) {
      showAlert("Fehler: Kein Newsletter-Entwurf gefunden!", 
        'Bitte erstelle einen E-Mail-Entwurf in Gmail mit "Update zu Haptigation" im Betreff.\n\n' +
        'Beispiel: "Dein Update zu Haptigation - Dezember 2025"');
      return;
    }
    
    // 2. Kontakte mit Label laden
    const contacts = getNewsletterContacts();
    if (contacts.length === 0) {
      showAlert("Keine Kontakte gefunden", 
        `Keine Kontakte mit dem Label "${CONFIG.contactLabel}" gefunden.\n\n` +
        `Stelle sicher, dass:\n` +
        `1. Du Kontakte in Gmail hast\n` +
        `2. Diese Kontakte das Label "${CONFIG.contactLabel}" haben\n` +
        `3. Die People API aktiviert ist`);
      return;
    }
    
    // 3. Bestätigung anzeigen
    const numRecipients = CONFIG.testMode ? 1 : contacts.length;
    const mode = CONFIG.testMode ? "TEST-MODUS" : "LIVE-VERSAND";
    const debugInfo = CONFIG.debugMode ? " | DEBUG-MODUS AKTIV" : "";
    
    Logger.log(`\n========== NEWSLETTER-VERSAND ==========`);
    Logger.log(`Modus: ${mode}${debugInfo}`);
    Logger.log(`Empfänger: ${numRecipients}`);
    Logger.log(`========================================\n`);
    
    // Automatisch fortfahren (da wir keine UI-Bestätigung haben)
    Logger.log("Starte Versand...");
    
    // 4. E-Mails versenden
    const results = sendEmails(draft, contacts);
    
    // 5. Entwurf löschen (optional - auskommentieren wenn du ihn behalten willst)
    // draft.deleteDraft();
    
    // 6. Ergebnis anzeigen
    showResults(results);
    
  } catch (error) {
    Logger.log("Fehler: " + error.toString());
    showAlert("Fehler beim Versand", error.toString());
  }
}

// ========== HILFSFUNKTIONEN ==========

/**
 * Lädt den Newsletter-Entwurf aus Gmail
 */
function getNewsletterDraft() {
  const drafts = GmailApp.getDrafts();
  
  if (drafts.length === 0) {
    return null;
  }
  
  // Suche nach Entwurf mit "Update zu Haptigation" im Betreff
  for (let draft of drafts) {
    const subject = draft.getMessage().getSubject();
    if (subject.includes("Update zu Haptigation")) {
      Logger.log(`✓ Newsletter-Entwurf gefunden: "${subject}"`);
      return draft;
    }
  }
  
  // Kein passender Entwurf gefunden
  Logger.log("Verfügbare Entwürfe:");
  drafts.forEach(d => {
    Logger.log(`- "${d.getMessage().getSubject()}"`);
  });
  
  return null;
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
        Logger.log(`Gefundene Group: ${group.name} (${group.resourceName})`);
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
    
    Logger.log(`Group hat ${response.memberCount || 0} Mitglieder`);
    
    if (!response.memberResourceNames || response.memberResourceNames.length === 0) {
      Logger.log("Keine Mitglieder in dieser Group gefunden");
      return contacts;
    }
    
    // 4. Details der Mitglieder laden (in Batches von 50)
    const memberResourceNames = response.memberResourceNames;
    const batchSize = 50;
    
    Logger.log(`Lade Details für ${memberResourceNames.length} Mitglieder...`);
    
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
              
              // WICHTIG: Prüfen ob E-Mail existiert
              if (!person.emailAddresses || person.emailAddresses.length === 0) {
                Logger.log(`⚠ Kontakt ohne E-Mail übersprungen`);
                return; // Diesen Kontakt überspringen
              }
              
              const email = person.emailAddresses[0].value;
              
              // Zusätzliche Validierung
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
              
              Logger.log(`✓ Kontakt geladen: ${name} (${anonymizeEmail(email)})`);
            }
          });
        }
      } catch (batchError) {
        Logger.log(`Fehler beim Laden eines Batches: ${batchError.toString()}`);
      }
    }
    
    Logger.log(`Insgesamt ${contacts.length} Kontakte mit Label "${CONFIG.contactLabel}" erfolgreich geladen`);
    
  } catch (error) {
    Logger.log("Fehler beim Laden der Kontakte: " + error.toString());
    Logger.log("Fehler-Details: " + JSON.stringify(error));
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
        personalizedSubject = `Test: ${personalizedSubject}`;
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
      Logger.log(`✓ E-Mail gesendet an: ${anonymizeEmail(contact.email)}`);
      
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
 * Zeigt das Ergebnis des Versands an
 */
function showResults(results) {
  const mode = CONFIG.testMode ? " (TEST-MODUS)" : "";
  let message = `Newsletter-Versand abgeschlossen${mode}!\n\n`;
  message += `✓ Erfolgreich: ${results.success}\n`;
  message += `✗ Fehlgeschlagen: ${results.failed}\n`;
  
  if (results.errors.length > 0) {
    message += `\nFehler:\n${results.errors.slice(0, 5).join('\n')}`;
    if (results.errors.length > 5) {
      message += `\n... und ${results.errors.length - 5} weitere`;
    }
  }
  
  if (CONFIG.testMode) {
    message += `\n\n⚠️ Test-Modus ist aktiv. Setze testMode auf false für den echten Versand.`;
  }
  
  showAlert("Versand abgeschlossen", message);
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
  // Wenn Debug-Modus aktiv ist, E-Mail nicht anonymisieren
  if (CONFIG.debugMode) {
    return email || '[keine E-Mail]';
  }
  
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