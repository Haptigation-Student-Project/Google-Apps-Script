/**
 * Delete User Data Automation
 * Löscht alle in Gmail gespeicherten Daten und Emails der angegebenen EMAIL und sendet vorher eine Bestätigung
 * Nutzt dafür den ersten E-Mail Draft der irgendwo im Betreff "Confirm User Data Deletion Draft - DO NOT DELETE" stehen hat
 */

// Konfiguration
const CONFIG = {
  DEBUG_MODE: false, // Auf true setzen für vollständige Email-Anzeige im Log
  MAX_THREADS_PER_BATCH: 500, // Maximale Anzahl von Threads pro Batch-Löschung
  EMAIL: "beispiel@gmail.com", // Ersetze mit echter E-Mail die gelöscht werden muss
  TESTMAIL: "test@gmail.com", // Ersetze mit echter E-Mail mit der du testen möchtest
  DRAFT_SUBJECT: 'Confirm User Data Deletion Draft - DO NOT DELETE', // Draft-Betreff in Gmail
  FINAL_EMAIL_SUBJECT: 'Wir haben ihre Nutzerdaten gelöscht', // Betreff der versendeten Email
};

/**
 * Hauptfunktion: Manuell ausführen mit Email-Adresse als Parameter
 */
function deleteByEmail() {
  const EMAIL_TO_DELETE = CONFIG.EMAIL;
  
  if (!EMAIL_TO_DELETE || EMAIL_TO_DELETE === 'beispiel@gmail.com') {
    log('ERROR', 'Bitte trage eine gültige Email-Adresse in der Variable EMAIL_TO_DELETE ein!');
    throw new Error('Keine Email-Adresse angegeben. Bitte CODE anpassen.');
  }
  
  if (!isValidEmail(EMAIL_TO_DELETE)) {
    log('ERROR', 'Ungültige Email-Adresse: ' + anonymizeEmail(EMAIL_TO_DELETE));
    throw new Error('Ungültige Email-Adresse angegeben.');
  }
  
  log('WARNING', '⚠️ ACHTUNG: Löschvorgang wird in 3 Sekunden gestartet...');
  log('WARNING', `Ziel: ${anonymizeEmail(EMAIL_TO_DELETE)}`);
  Utilities.sleep(3000);
  
  deleteAllDataForEmail(EMAIL_TO_DELETE);
}

/**
 * Hauptfunktion: Löscht alle Daten für eine Email-Adresse
 */
function deleteAllDataForEmail(email) {
  log('INFO', `========================================`);
  log('INFO', `Starte Löschvorgang für: ${anonymizeEmail(email)}`);
  log('INFO', `========================================`);
  
  const results = {
    contactsDeleted: 0,
    threadsDeleted: 0,
    errors: []
  };
  
  try {
    // Schritt 1: Sende Bestätigungsmail
    log('INFO', 'SCHRITT 1: Bestätigungsmail versenden...');
    const emailSent = sendConfirmationEmail(email);
    
    if (!emailSent) {
      log('WARNING', '⚠️ Bestätigungsmail konnte nicht versendet werden - Löschvorgang wird trotzdem fortgesetzt');
    }
    
    // Schritt 2: Suche und lösche Kontakte
    log('INFO', 'SCHRITT 2: Kontakte durchsuchen...');
    results.contactsDeleted = deleteContactsByEmail(email);
    
    // Schritt 3: Suche und lösche E-Mails
    log('INFO', 'SCHRITT 3: E-Mails durchsuchen...');
    results.threadsDeleted = deleteEmailsByAddress(email);
    
    // Zusammenfassung
    log('SUCCESS', `========================================`);
    log('SUCCESS', `✅ Löschvorgang abgeschlossen für: ${anonymizeEmail(email)}`);
    log('SUCCESS', `📊 Gelöschte Kontakte: ${results.contactsDeleted}`);
    log('SUCCESS', `📧 Gelöschte E-Mail-Threads: ${results.threadsDeleted}`);
    log('SUCCESS', `========================================`);
    
    return results;
    
  } catch (error) {
    log('ERROR', `========================================`);
    log('ERROR', `❌ Fehler beim Löschvorgang: ${error.message}`);
    log('ERROR', `Stack: ${error.stack}`);
    log('ERROR', `========================================`);
    throw error;
  }
}

/**
 * Sendet eine Bestätigungsmail aus dem Draft-Template mit HTML und Inline-Bildern
 */
function sendConfirmationEmail(email) {
  log('INFO', `📧 Suche Draft mit Betreff: "${CONFIG.DRAFT_SUBJECT}"`);
  
  try {
    // Suche nach dem Draft
    const drafts = GmailApp.getDrafts();
    let templateDraft = null;
    
    for (let draft of drafts) {
      const message = draft.getMessage();
      if (message.getSubject() === CONFIG.DRAFT_SUBJECT) {
        templateDraft = draft;
        log('SUCCESS', `✓ Draft gefunden: "${CONFIG.DRAFT_SUBJECT}"`);
        break;
      }
    }
    
    if (!templateDraft) {
      log('ERROR', `❌ Draft mit Betreff "${CONFIG.DRAFT_SUBJECT}" nicht gefunden!`);
      log('ERROR', '💡 Bitte erstelle einen Draft mit exakt diesem Betreff.');
      return false;
    }
    
    // Hole Draft-Nachricht
    const draftMessage = templateDraft.getMessage();
    
    // Hole HTML-Body (falls vorhanden) oder Plain Text als Fallback
    let htmlBody = draftMessage.getBody(); // HTML-Body
    let plainBody = draftMessage.getPlainBody(); // Plain Text Fallback
    
    // Verwende den konfigurierten finalen Betreff
    let subject = CONFIG.FINAL_EMAIL_SUBJECT;
    
    // Ersetze Platzhalter in beiden Versionen
    htmlBody = htmlBody.replace(/\{EMAIL\}/g, email);
    plainBody = plainBody.replace(/\{EMAIL\}/g, email);
    subject = subject.replace(/\{EMAIL\}/g, email);
    
    // Hole alle Anhänge (Inline-Bilder sind als Attachments gespeichert)
    const attachments = draftMessage.getAttachments();
    
    log('INFO', `📤 Sende Bestätigungsmail an: ${anonymizeEmail(email)}`);
    log('INFO', `📋 Betreff: "${subject}"`);
    log('INFO', `🖼️  Anhänge/Inline-Bilder: ${attachments.length}`);
    
    // Erstelle erweiterte Mail-Optionen
    const mailOptions = {
      htmlBody: htmlBody,           // HTML-Version mit Inline-Bildern
      inlineImages: {},             // Container für Inline-Bilder
      attachments: []               // Echte Anhänge (keine Inline-Bilder)
    };
    
    // Verarbeite Anhänge: Unterscheide zwischen Inline-Bildern und echten Anhängen
    attachments.forEach((attachment, index) => {
      const contentId = attachment.getName();
      const isInlineImage = htmlBody.includes(`cid:${contentId}`) || 
                           htmlBody.includes(`src="${contentId}"`);
      
      if (isInlineImage) {
        // Dies ist ein Inline-Bild
        mailOptions.inlineImages[contentId] = attachment;
        log('INFO', `   🖼️  Inline-Bild: ${contentId}`);
      } else {
        // Dies ist ein echter Anhang
        mailOptions.attachments.push(attachment);
        log('INFO', `   📎 Anhang: ${contentId}`);
      }
    });
    
    // Sende E-Mail mit HTML und Inline-Bildern
    GmailApp.sendEmail(
      email,
      subject,
      plainBody,      // Plain Text Version (Fallback)
      mailOptions     // Erweiterte Optionen mit HTML und Inline-Bildern
    );
    
    log('SUCCESS', `✓ Bestätigungsmail erfolgreich versendet an ${anonymizeEmail(email)}`);
    log('SUCCESS', `   📧 Format: HTML mit ${Object.keys(mailOptions.inlineImages).length} Inline-Bild(ern)`);
    
    // Kurze Pause nach dem Versand
    Utilities.sleep(1000);
    
    return true;
    
  } catch (error) {
    log('ERROR', `❌ Fehler beim Versenden der Bestätigungsmail: ${error.message}`);
    log('ERROR', `Stack: ${error.stack}`);
    return false;
  }
}

/**
 * Löscht alle Kontakte mit der angegebenen Email-Adresse
 */
function deleteContactsByEmail(email) {
  log('INFO', `🔍 Suche Kontakte für: ${anonymizeEmail(email)}`);
  let deletedCount = 0;
  
  try {
    // Suche Kontakte mit People API
    const contacts = People.People.Connections.list('people/me', {
      personFields: 'names,emailAddresses',
      pageSize: 1000
    });
    
    if (!contacts.connections || contacts.connections.length === 0) {
      log('INFO', '📭 Keine Kontakte in Google Contacts gefunden');
      return 0;
    }
    
    log('INFO', `📋 ${contacts.connections.length} Kontakte gefunden, durchsuche nach Übereinstimmungen...`);
    
    // Filtere Kontakte nach Email-Adresse
    const matchingContacts = contacts.connections.filter(person => {
      if (!person.emailAddresses) return false;
      return person.emailAddresses.some(emailObj => 
        emailObj.value.toLowerCase() === email.toLowerCase()
      );
    });
    
    if (matchingContacts.length === 0) {
      log('INFO', '📭 Keine passenden Kontakte gefunden');
      return 0;
    }
    
    log('INFO', `🎯 ${matchingContacts.length} passende Kontakte gefunden`);
    
    // Lösche jeden gefundenen Kontakt
    matchingContacts.forEach((contact, index) => {
      try {
        const resourceName = contact.resourceName;
        const name = contact.names && contact.names[0] ? contact.names[0].displayName : 'Unbekannt';
        
        log('INFO', `🗑️  Lösche Kontakt ${index + 1}/${matchingContacts.length}: "${name}" (${anonymizeEmail(email)})`);
        
        People.People.deleteContact(resourceName);
        deletedCount++;
        
        log('SUCCESS', `✓ Kontakt erfolgreich gelöscht: "${name}"`);
        
        // Kurze Pause zwischen Löschvorgängen
        Utilities.sleep(300);
        
      } catch (error) {
        log('ERROR', `❌ Fehler beim Löschen von Kontakt: ${error.message}`);
        results.errors.push(`Kontakt: ${error.message}`);
      }
    });
    
  } catch (error) {
    log('ERROR', `❌ Fehler bei Kontaktsuche: ${error.message}`);
    throw error;
  }
  
  log('INFO', `✅ Kontakt-Löschung abgeschlossen: ${deletedCount} Kontakt(e) gelöscht`);
  return deletedCount;
}

/**
 * Löscht alle E-Mail-Threads von/an eine Email-Adresse
 */
function deleteEmailsByAddress(email) {
  log('INFO', `🔍 Suche E-Mails für: ${anonymizeEmail(email)}`);
  let deletedCount = 0;
  
  try {
    // Suche nach E-Mails von oder an diese Adresse
    const searchQuery = `from:${email} OR to:${email}`;
    log('INFO', `🔎 Verwende Suchquery: from:${anonymizeEmail(email)} OR to:${anonymizeEmail(email)}`);
    
    let start = 0;
    let hasMore = true;
    let batchNumber = 0;
    
    while (hasMore) {
      batchNumber++;
      const threads = GmailApp.search(searchQuery, start, CONFIG.MAX_THREADS_PER_BATCH);
      
      if (threads.length === 0) {
        hasMore = false;
        break;
      }
      
      log('INFO', `📦 Batch #${batchNumber}: ${threads.length} Thread(s) gefunden (Start-Position: ${start})`);
      
      // Lösche Threads in diesem Batch
      threads.forEach((thread, index) => {
        try {
          const subject = thread.getFirstMessageSubject();
          const messageCount = thread.getMessageCount();
          const threadId = start + index + 1;
          
          log('INFO', `🗑️  Lösche Thread #${threadId}: "${subject}" (${messageCount} Nachricht(en))`);
          
          thread.moveToTrash();
          deletedCount++;
          
          log('SUCCESS', `✓ Thread #${threadId} in Papierkorb verschoben`);
          
        } catch (error) {
          log('ERROR', `❌ Fehler beim Löschen von Thread: ${error.message}`);
        }
      });
      
      // Prüfe ob weitere Threads existieren
      start += threads.length;
      
      if (threads.length < CONFIG.MAX_THREADS_PER_BATCH) {
        hasMore = false;
      }
      
      // Kurze Pause zwischen Batches
      if (hasMore) {
        log('INFO', `⏸️  Pause vor nächstem Batch...`);
        Utilities.sleep(500);
      }
    }
    
    if (deletedCount === 0) {
      log('INFO', '📭 Keine E-Mails gefunden');
    }
    
  } catch (error) {
    log('ERROR', `❌ Fehler bei E-Mail-Suche: ${error.message}`);
    throw error;
  }
  
  log('INFO', `✅ E-Mail-Löschung abgeschlossen: ${deletedCount} Thread(s) in Papierkorb verschoben`);
  return deletedCount;
}

/**
 * Validiert eine Email-Adresse
 */
function isValidEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

/**
 * Anonymisiert eine Email-Adresse für Logs
 */
function anonymizeEmail(email, partial = false) {
  if (CONFIG.DEBUG_MODE) {
    return email;
  }
  
  if (!email || !email.includes('@')) {
    return '***';
  }
  
  const [localPart, domain] = email.split('@');
  const domainParts = domain.split('.');
  
  if (partial) {
    // Zeigt ersten Buchstaben: m...@g...
    return `${localPart[0]}...@${domainParts[0][0]}...`;
  }
  
  // Vollständige Anonymisierung mit TLD: m...@g....de
  const tld = domainParts[domainParts.length - 1];
  return `${localPart[0]}...@${domainParts[0][0]}....${tld}`;
}

/**
 * Zentrales Logging mit Zeitstempel
 */
function log(level, message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const logMessage = `[${timestamp}] [${level.padEnd(7)}] ${message}`;
  Logger.log(logMessage);  // ← Nur diese Zeile behalten
  // console.log(logMessage); ← Diese Zeile löschen oder auskommentieren
}

/**
 * Zeigt alle Logs an
 */
function showLogs() {
  const logs = Logger.getLog();
  console.log(logs);
  return logs;
}

/**
 * Testfunktion: Anonymisierung testen
 */
function testAnonymization() {
  console.log('=== TEST: Email-Anonymisierung ===\n');
  
  console.log('DEBUG Mode OFF:');
  CONFIG.DEBUG_MODE = false;
  console.log('Voll:', anonymizeEmail('max.mustermann@gmail.com', false));
  console.log('Partial:', anonymizeEmail('max.mustermann@gmail.com', true));
  console.log('');
  
  console.log('DEBUG Mode ON:');
  CONFIG.DEBUG_MODE = true;
  console.log('Voll:', anonymizeEmail('max.mustermann@gmail.com', false));
  console.log('Partial:', anonymizeEmail('max.mustermann@gmail.com', true));
  
  // Zurücksetzen
  CONFIG.DEBUG_MODE = false;
}

/**
 * Testfunktion: Bestätigungsmail testen (ohne zu löschen)
 */
function testConfirmationEmail() {
  const TEST_EMAIL = CONFIG.TESTMAIL; 
  
  if (TEST_EMAIL === 'test@gmail.com') {
    log('ERROR', 'Bitte TEST_EMAIL Variable anpassen!');
    return;
  }
  
  log('INFO', '=== TEST: Bestätigungsmail-Versand ===');
  log('INFO', `Ziel-Email: ${anonymizeEmail(TEST_EMAIL)}`);
  
  const success = sendConfirmationEmail(TEST_EMAIL);
  
  if (success) {
    log('SUCCESS', '✅ Test erfolgreich! Prüfe dein Postfach.');
  } else {
    log('ERROR', '❌ Test fehlgeschlagen. Siehe Logs oben für Details.');
  }
}

/**
 * Testfunktion: Kontakte auflisten (ohne zu löschen)
 */
function listContactsForEmail() {
  const TEST_EMAIL = CONFIG.TESTMAIL;
  
  if (TEST_EMAIL === 'test@gmail.com') {
    log('ERROR', 'Bitte TEST_EMAIL Variable anpassen!');
    return;
  }
  
  log('INFO', `Suche Kontakte für: ${anonymizeEmail(TEST_EMAIL)}`);
  
  const contacts = People.People.Connections.list('people/me', {
    personFields: 'names,emailAddresses',
    pageSize: 1000
  });
  
  if (!contacts.connections) {
    log('INFO', 'Keine Kontakte gefunden');
    return;
  }
  
  const matchingContacts = contacts.connections.filter(person => {
    if (!person.emailAddresses) return false;
    return person.emailAddresses.some(emailObj => 
      emailObj.value.toLowerCase() === TEST_EMAIL.toLowerCase()
    );
  });
  
  log('INFO', `Gefundene Kontakte: ${matchingContacts.length}`);
  
  matchingContacts.forEach((contact, i) => {
    const name = contact.names && contact.names[0] ? contact.names[0].displayName : 'Unbekannt';
    const emails = contact.emailAddresses.map(e => anonymizeEmail(e.value)).join(', ');
    log('INFO', `${i + 1}. ${name} - Emails: ${emails}`);
  });
}