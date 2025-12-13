/**
 * Newsletter Abmelde-Automatisierung
 * Wird automatisch ausgeführt wenn jemand das Google Form abschickt
 * DSGVO-konform: E-Mails werden in Logs anonymisiert
 */

// ========== KONFIGURATION ==========
const CONFIG = {
  // Das Label deiner Newsletter-Kontakte
  contactLabel: "Newsletter Subscriber",

  // Form ID
  formID: "1rl0jU6RSHaf0tiHkmy4qUIfq8nWv2IHU6FdVy2mYSrQ",
  
  // Betreff der Bestätigungs-E-Mail (kannst du anpassen)
  confirmationSubject: "Haptigation Newsletter-Abmeldung bestätigt",
  
  // Name des Absenders
  senderName: "Haptigation Team",
  
  // Debug-Modus (nur für Tests aktivieren!)
  debugMode: false  // Auf TRUE setzen nur wenn du testest!
};

// ========== HILFSFUNKTION FÜR SICHERES LOGGING ==========

/**
 * Anonymisiert E-Mail für Logs (DSGVO)
 */
function anonymizeEmail(email) {
  if (!email) return "keine-email";
  if (CONFIG.debugMode) return email; // Nur im Debug-Modus volle E-Mail
  
  // Zeigt nur: t***@g***.de
  const parts = email.split('@');
  if (parts.length !== 2) return "ungültig";
  
  const local = parts[0].charAt(0) + '***';
  const domain = parts[1].split('.');
  const domainAnon = domain[0].charAt(0) + '***.' + domain.slice(1).join('.');
  
  return local + '@' + domainAnon;
}

/**
 * Sicheres Logging (respektiert debugMode)
 */
function log(message, includesSensitiveData = false) {
  if (CONFIG.debugMode || !includesSensitiveData) {
    Logger.log(message);
  }
}

// ========== HAUPTFUNKTION ==========

/**
 * Wird automatisch durch Form-Submit getriggert
 */
function onFormSubmit(e) {
  try {
    log("Form-Submit Event empfangen", false);
    
    // E-Mail aus Form-Response holen
    let email = null;
    
    // Methode 1: Aus e.values (wenn verfügbar)
    if (e.values && e.values.length > 1) {
      email = e.values[1]; // Index 1 ist normalerweise erste Antwort nach Timestamp
    }
    
    // Methode 2: Aus e.response (Alternative)
    if (!email && e.response) {
      const itemResponses = e.response.getItemResponses();
      if (itemResponses && itemResponses.length > 0) {
        email = itemResponses[0].getResponse();
      }
    }
    
    // Methode 3: Aus e.namedValues (noch eine Alternative)
    if (!email && e.namedValues) {
      const keys = Object.keys(e.namedValues);
      if (keys.length > 0) {
        email = e.namedValues[keys[0]][0]; // Erste Antwort
      }
    }
    
    if (!email || email.trim() === '') {
      log("Fehler: Keine E-Mail-Adresse erhalten", false);
      if (CONFIG.debugMode) {
        log("Event Structure:", false);
        log(JSON.stringify(e), true);
      }
      return;
    }
    
    email = email.trim().toLowerCase(); // Normalisieren
    const anonEmail = anonymizeEmail(email);
    
    log(`\n========== ABMELDUNG ==========`, false);
    log(`E-Mail: ${anonEmail}`, false);
    log(`================================\n`, false);
    
    // 1. Kontakt finden und Label entfernen
    const success = removeNewsletterLabel(email);
    
    if (success) {
      // 2. Bestätigungs-E-Mail senden
      sendUnsubscribeConfirmation(email);
      
      // 3. Form-Antwort löschen (Datenschutz)
      deleteFormResponse(e);
      
      log(`✓ Abmeldung erfolgreich abgeschlossen für: ${anonEmail}`, false);
    } else {
      log(`✗ Abmeldung fehlgeschlagen für: ${anonEmail}`, false);
    }
    
  } catch (error) {
    log(`Fehler bei Abmeldung: ${error.toString()}`, false);
    log(`Stack: ${error.stack}`, false);
    
    // Admin benachrichtigen
    try {
      notifyAdminAboutError(email ? anonymizeEmail(email) : "unbekannt", error);
    } catch (e2) {
      log(`Konnte Admin nicht benachrichtigen: ${e2.toString()}`, false);
    }
  }
}

// ========== HILFSFUNKTIONEN ==========

/**
 * Entfernt das Newsletter-Label von einem Kontakt
 */
function removeNewsletterLabel(email) {
  try {
    const anonEmail = anonymizeEmail(email);
    email = email.trim().toLowerCase(); // Normalisieren
    log(`Suche Kontakt mit E-Mail: ${anonEmail}`, false);
    
    // 1. Zuerst die Newsletter-Gruppe finden
    const contactGroups = People.ContactGroups.list();
    let newsletterGroupResourceName = null;
    
    if (!contactGroups.contactGroups) {
      log(`✗ Keine Contact Groups gefunden`, false);
      return false;
    }
    
    for (let group of contactGroups.contactGroups) {
      if (group.name === CONFIG.contactLabel) {
        newsletterGroupResourceName = group.resourceName;
        log(`✓ Newsletter-Gruppe gefunden: ${group.name}`, false);
        break;
      }
    }
    
    if (!newsletterGroupResourceName) {
      log(`✗ Label "${CONFIG.contactLabel}" nicht gefunden`, false);
      if (CONFIG.debugMode) {
        log("Verfügbare Groups:", false);
        contactGroups.contactGroups.forEach(g => log(`  - ${g.name}`, false));
      }
      return false;
    }
    
    // 2. Alle Mitglieder der Gruppe laden
    const response = People.ContactGroups.get(newsletterGroupResourceName, {
      maxMembers: 500
    });
    
    if (!response.memberResourceNames || response.memberResourceNames.length === 0) {
      log(`✗ Keine Mitglieder in der Newsletter-Gruppe gefunden`, false);
      return false;
    }
    
    log(`Gruppe hat ${response.memberResourceNames.length} Mitglieder`, false);
    
    // 3. Durch alle Mitglieder gehen und nach E-Mail suchen
    const memberResourceNames = response.memberResourceNames;
    const batchSize = 50;
    let foundResourceName = null;
    
    for (let i = 0; i < memberResourceNames.length; i += batchSize) {
      const batch = memberResourceNames.slice(i, i + batchSize);
      
      try {
        const batchResponse = People.People.getBatchGet({
          resourceNames: batch,
          personFields: 'names,emailAddresses'
        });
        
        if (batchResponse.responses) {
          for (let personResponse of batchResponse.responses) {
            if (personResponse.person && personResponse.person.emailAddresses) {
              const person = personResponse.person;
              
              // Alle E-Mail-Adressen des Kontakts durchgehen
              for (let emailObj of person.emailAddresses) {
                const contactEmail = emailObj.value.trim().toLowerCase();
                
                if (contactEmail === email) {
                  foundResourceName = person.resourceName;
                  log(`✓ Kontakt gefunden: ${anonEmail}`, false);
                  break;
                }
              }
              
              if (foundResourceName) break;
            }
          }
        }
        
        if (foundResourceName) break;
        
      } catch (batchError) {
        log(`Fehler beim Laden eines Batches: ${batchError.toString()}`, false);
      }
    }
    
    if (!foundResourceName) {
      log(`✗ Kontakt mit E-Mail ${anonEmail} nicht in der Newsletter-Liste gefunden`, false);
      log(`Hinweis: Stelle sicher, dass die E-Mail exakt übereinstimmt`, false);
      return false;
    }
    
    // 4. Kontakt aus der Gruppe entfernen
    People.ContactGroups.Members.modify({
      resourceNamesToRemove: [foundResourceName]
    }, newsletterGroupResourceName);
    
    log(`✓ ${anonEmail} erfolgreich aus Newsletter-Liste entfernt`, false);
    return true;
    
  } catch (error) {
    log(`✗ Fehler beim Entfernen des Labels: ${error.toString()}`, false);
    log(`Stack: ${error.stack}`, false);
    return false;
  }
}

/**
 * Sendet Bestätigungs-E-Mail
 */
function sendUnsubscribeConfirmation(email) {
  try {
    const anonEmail = anonymizeEmail(email);
    
    // Draft mit dem spezifischen Betreff suchen
    const drafts = GmailApp.getDrafts();
    let confirmationDraft = null;
    
    for (let draft of drafts) {
      const subject = draft.getMessage().getSubject();
      if (subject.includes("Unsubscribe Newsletter - DO NOT DELETE")) {
        confirmationDraft = draft;
        log(`✓ Bestätigungs-Draft gefunden`, false);
        break;
      }
    }
    
    if (!confirmationDraft) {
      log(`✗ Kein Draft mit "Unsubscribe Newsletter - DO NOT DELETE" im Betreff gefunden`, false);
      if (CONFIG.debugMode) {
        log("Verfügbare Drafts:", false);
        drafts.forEach(d => {
          log(`- "${d.getMessage().getSubject()}"`, false);
        });
      }
      return;
    }
    
    // Draft-Inhalt holen
    const message = confirmationDraft.getMessage();
    const htmlBody = message.getBody();
    const plainBody = message.getPlainBody();
    
    // Personalisierung: [EMAIL] durch tatsächliche E-Mail ersetzen
    const personalizedHtml = htmlBody.replace(/\[EMAIL\]/g, email);
    const personalizedPlain = plainBody.replace(/\[EMAIL\]/g, email);
    
    // E-Mail versenden
    GmailApp.sendEmail(
      email,
      CONFIG.confirmationSubject,
      personalizedPlain,
      {
        htmlBody: personalizedHtml,
        name: CONFIG.senderName,
        noReply: false
      }
    );
    
    log(`✓ Bestätigungs-E-Mail gesendet an: ${anonEmail}`, false);
    
  } catch (error) {
    log(`✗ Fehler beim Senden der Bestätigung: ${error.toString()}`, false);
  }
}

/**
 * Löscht die Form-Antwort (Datenschutz)
 */
function deleteFormResponse(e) {
  try {
    // Form über die CONFIG-ID öffnen (nicht getActiveForm!)
    const form = FormApp.openById(CONFIG.formID);
    
    // Response ID aus dem Event-Objekt holen
    if (e.response) {
      const responseId = e.response.getId();
      form.deleteResponse(responseId);
      log(`✓ Form-Antwort gelöscht (Datenschutz)`, false);
      return;
    }
    
    // Fallback: Letzte Response löschen (weniger präzise)
    const formResponses = form.getResponses();
    if (formResponses.length > 0) {
      const lastResponse = formResponses[formResponses.length - 1];
      form.deleteResponse(lastResponse.getId());
      log(`✓ Form-Antwort gelöscht (Datenschutz) - Fallback zur letzten Response`, false);
    }
    
  } catch (error) {
    log(`⚠ Warnung: Form-Antwort konnte nicht gelöscht werden: ${error.toString()}`, false);
    // Nicht kritisch, daher kein throw
  }
}

/**
 * Benachrichtigt Admin über Fehler
 */
function notifyAdminAboutError(anonEmail, error) {
  try {
    const adminEmail = "haptigation@gmail.com"
    
    GmailApp.sendEmail(
      adminEmail,
      "⚠️ Fehler bei Newsletter-Abmeldung",
      `Es gab einen Fehler bei der Abmeldung von: ${anonEmail}\n\n` +
      `Fehler: ${error.toString()}\n\n` +
      `Hinweis: E-Mail wurde anonymisiert. Für Details debugMode aktivieren.\n\n` +
      `Bitte manuell überprüfen.`,
      {
        name: "Haptigation Newsletter System"
      }
    );
    
  } catch (e) {
    log(`Fehler beim Senden der Admin-Benachrichtigung: ${e.toString()}`, false);
  }
}

// ========== SETUP-FUNKTIONEN ==========

/**
 * Einmalig ausführen um den Trigger zu erstellen
 */
function setupFormTrigger() {
  try {
    const formId = CONFIG.formID;
    
    const form = FormApp.openById(formId);
    
    // Alte Trigger löschen (falls vorhanden)
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onFormSubmit') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Neuen Trigger erstellen
    ScriptApp.newTrigger('onFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();
    
    Logger.log("✓ Trigger erfolgreich erstellt!");
    Logger.log(`Form: ${form.getTitle()}`);
    
  } catch (error) {
    Logger.log(`✗ Fehler beim Erstellen des Triggers: ${error.toString()}`);
    throw error;
  }
}

/**
 * Trigger entfernen (falls nötig)
 */
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("✓ Trigger entfernt");
    }
  });
}

/**
 * Test-Funktion - simuliert eine Abmeldung
 */
function testUnsubscribe() {
  const testEmail = "test@gmx.de"; // HIER DEINE TEST-EMAIL
  
  Logger.log("===== TEST-MODUS =====");
  Logger.log(`Test mit E-Mail: ${testEmail}`);
  
  const success = removeNewsletterLabel(testEmail);
  
  if (success) {
    sendUnsubscribeConfirmation(testEmail);
    Logger.log("✓ Test erfolgreich!");
  } else {
    Logger.log("✗ Test fehlgeschlagen");
  }
}

/**
 * Debug: Zeigt alle Newsletter-Kontakte
 */
function listNewsletterContacts() {
  try {
    Logger.log(`\n=== Suche nach Label: "${CONFIG.contactLabel}" ===\n`);
    
    const contactGroups = People.ContactGroups.list();
    
    if (!contactGroups.contactGroups) {
      Logger.log("Keine Contact Groups gefunden!");
      return;
    }
    
    Logger.log("Verfügbare Contact Groups:");
    contactGroups.contactGroups.forEach(g => {
      Logger.log(`  - ${g.name} (${g.memberCount || 0} Mitglieder)`);
    });
    
    let newsletterGroupResourceName = null;
    
    for (let group of contactGroups.contactGroups) {
      if (group.name === CONFIG.contactLabel) {
        newsletterGroupResourceName = group.resourceName;
        Logger.log(`\n✓ Newsletter-Gruppe gefunden: ${group.name}`);
        break;
      }
    }
    
    if (!newsletterGroupResourceName) {
      Logger.log(`\n✗ Label "${CONFIG.contactLabel}" nicht gefunden!`);
      return;
    }
    
    const response = People.ContactGroups.get(newsletterGroupResourceName, {
      maxMembers: 500
    });
    
    if (!response.memberResourceNames || response.memberResourceNames.length === 0) {
      Logger.log("Keine Mitglieder in dieser Gruppe");
      return;
    }
    
    Logger.log(`\n=== Newsletter-Kontakte (${response.memberResourceNames.length}) ===\n`);
    
    const batchResponse = People.People.getBatchGet({
      resourceNames: response.memberResourceNames,
      personFields: 'names,emailAddresses'
    });
    
    if (batchResponse.responses) {
      batchResponse.responses.forEach((personResponse, index) => {
        if (personResponse.person && personResponse.person.emailAddresses) {
          const person = personResponse.person;
          const name = person.names && person.names.length > 0 
            ? person.names[0].displayName 
            : "Unbekannt";
          const email = person.emailAddresses[0].value;
          const anonEmail = anonymizeEmail(email);
          
          // Im Debug-Modus volle Info, sonst anonymisiert
          if (CONFIG.debugMode) {
            Logger.log(`${index + 1}. ${name} <${email}>`);
          } else {
            Logger.log(`${index + 1}. Kontakt <${anonEmail}>`);
          }
        }
      });
    }
    
  } catch (error) {
    Logger.log(`Fehler: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}
