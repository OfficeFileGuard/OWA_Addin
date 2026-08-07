import JSZip from "jszip";
declare const ADDIN_BASE_URL: string; // Variabile definita da un plugin di webpack.config.js - disponibile a runtime

console.log("commands.ts caricato e pronto!");
console.debug("commands.ts caricato e pronto!");
console.error("commands.ts caricato e pronto!");
console.warn("commands.ts caricato e pronto!");
console.info("commands.ts caricato e pronto!");


Office.onReady(() => {
  console.log("Office ready");
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

/** 
 * Starting point quando la mail viene inviata
 * Wrappa la funzione onSendHandler (è la funzione che realmente svolge il lavoro) in una nuova promise con un timeout
 * Se si pianta il server officefileguard, al timeout restuisce sempre true (permette l'invio della mail)
 * 
 * @param event - L'evento associato al click sul pulsante di invio.
 * @returns - Void
 */
async function onMessageSendHandler(event: Office.AddinCommands.Event) {
  console.debug("onMessageSendHandler triggered");
  console.log("onMessageSendHandler INIZIATA!");

  const timeoutMS = 60000;  // Timeout in millisecondi

  try { await withTimeout(onSendHandler(event), timeoutMS); }    // Wraps all the logic in a new Promise with timeout  
  catch (error) {
    console.error("Timeout or error during execution:", error);
    event.completed({ allowEvent: true });  // In case of timeout, we allow the email to be sent anyway
  }
}


/**
 * Funzione che esegue tutta la logica principale dell'add-in
 * 
 * @param event - L'evento associato al click sul pulsante di invio.
 * @returns - Void
 */
async function onSendHandler(event: Office.AddinCommands.Event): Promise<void> {
  // Ottiene item l'oggetto email in composizione 
  const item = Office.context.mailbox.item as Office.MessageCompose;
  // se non l'ottiene consente l'invio e ritorna
  if (!item) {
    event.completed({ allowEvent: true });
    return;
  }

  let activeDialog: Office.Dialog | null = null;

  const closeDialogSafely = () => {
    if (activeDialog) {
      try {
        activeDialog.close();
      } catch (e) {
        console.warn("Exception closing dialog:", e);
      }
      activeDialog = null;
    }
  };

  try {
    // Verifica se ci sono destinatari esterni. Se non ci sono permette invio e esce
    console.debug("Checking for external recipients...");
    const hasExternalRecipients = await checkExternalRecipients(item);
    if (!hasExternalRecipients) {
      event.completed({ allowEvent: true });
      return;
    }

    // Verifica se ci sono allegati riservati. Se non ci sono permette invio e esce
    console.debug("Checking for reserved attachments...");
    const hasReservedAttachments = await checkReservedAttachments(item);
    if (!hasReservedAttachments) {
      event.completed({ allowEvent: true });
      return;
    }

    // Entrambe le condizioni vere — apre il dialog e attende la risposta
    await new Promise<void>((resolveDialog) => {   //Promise attende la risposta del dialog
      console.warn("warning dialog : ", `${ADDIN_BASE_URL}/guardalert.html`);
      Office.context.ui.displayDialogAsync(
        `${ADDIN_BASE_URL}/guardalert.html`,    //ADDIN_BASE_URL è una variabile definita da un plugin di webpack.config.js
        { height: 44, width: 40, displayInIframe: false },
        (asyncResult) => {
          // Se il dialog non si apre, permette l'invio della mail e esce
          console.warn("sto per fare verifica stato Office UI")
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog:", asyncResult.error);
            event.completed({ allowEvent: true });
            resolveDialog();  // risolve la Promise per terminare l'attesa
            return;
          }

          activeDialog = asyncResult.value;     // activeDialog contiene l'oggetto dialog aperto
          let dialogHandled = false;          // indica se il dialog è stato gestito

          // creazione di un event handler di tipo DialogMessageReceived (viene attivato quando il dialog invia un messaggio)
          activeDialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (args: { message: string; origin: string | undefined; } | { error: number; }) => {
              // se il dialog è già stato gestito, esce
              if (dialogHandled) return;
              dialogHandled = true; // altrimenti, imposta che il dialogo è stato gestito

              // se non c'è il messaggio, chiude in sicurezza il dialog, sblocca l'invio e esce
              if (!('message' in args)) {
                console.warn('DialogMessageReceived event triggered without a "message" property. Proceeding with send');
                closeDialogSafely();
                event.completed({ allowEvent: true });
                resolveDialog();
                return;
              }

              // se invece c'è il messaggio
              const message = args.message;
              closeDialogSafely();   // chiude sempre in sicurezza il dialog

              // se il messaggio è "send", chiude l'evento e permette l'invio della mail
              if (message === "send") {
                event.completed({ allowEvent: true });  // addin ha terminato il suo intervento e permette l'invio della mail
              }
              // altrimenti chiude l'evento e blocca l'invio della mail e mostra un messaggio di errore
              else {
                (event as any).completed({ allowEvent: false, errorMessage: "⚠ This email contains confidential attachments addressed to external recipients. Please review before sending." });
              }
              resolveDialog();
            }
          );

          // creazione di un event handler di tipo DialogEventReceived (viene attivato quando il dialog viene chiuso dall'utente con la X)
          activeDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            if (dialogHandled) return;    // se il dialog è già stato gestito, esce
            dialogHandled = true;        // altrimenti, imposta che il dialogo è stato gestito
            activeDialog = null;         // il dialog è già stato chiuso dall'utente via UI

            // Si arriva qui se utente chiude dialog senza scegliere uno dei pulsanti
            // quindi blocca l'invio della mail e mostra un messaggio di errore
            (event as any).completed({ allowEvent: false, errorMessage: "⚠ This email contains confidential attachments addressed to external recipients. Please review before sending." });
            resolveDialog();
          });
        }
      );
    });

  }
  catch (error) {  // In caso di errore imprevisto, chiude il dialog se aperto, permette l'invio della mail e esce
    console.error("An error occurred while checking the email: sending anyway", error);
    closeDialogSafely();
    event.completed({ allowEvent: true });
  }
}


/**
 * Verifica se i destinatari dell'email sono esterni.
 * 
 * @param item - L'oggetto email da controllare.
 * @returns Un promise che risolve con true se ci sono destinatari esterni, false altrimenti.
 */
async function checkExternalRecipients(item: Office.MessageCompose): Promise<boolean> {
  const senderEmail = Office.context.mailbox.userProfile.emailAddress;
  const senderDomain = senderEmail.split("@")[1]?.toLowerCase() ?? "";

  const getRecipients = (field: Office.Recipients): Promise<Office.EmailAddressDetails[]> => {
    return new Promise((resolve) => {
      field.getAsync((result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
      });
    });
  };

  const [to, cc, bcc] = await Promise.all([getRecipients(item.to), getRecipients(item.cc), getRecipients(item.bcc)]);

  const allRecipients = [...to, ...cc, ...bcc]; // Combina tutti i destinatari in un unico array
  return allRecipients.some(r => {
    const domain = r.emailAddress.split("@")[1]?.toLowerCase() ?? "";
    return domain !== senderDomain && domain !== "";
  });
}

/**
 * Verifica se tra i file allegati all'email ce ne sono di riservati.
 * 
 * @param item - L'oggetto email da controllare.
 * @returns Un promise che risolve con true se ci sono allegati riservati, false altrimenti.
 */
async function checkReservedAttachments(item: Office.MessageCompose): Promise<boolean> {
  return new Promise((resolve) => {
    item.getAttachmentsAsync(async (result) => {
      // Se si verifica un errore o non ci sono allegati, risolve la promise con false e ritorna (dal callback getAttachmentsAsync non dalla funzione checkReservedAttachments)
      console.log("there are ", result, " attachments");
      if (result.status !== Office.AsyncResultStatus.Succeeded || result.value.length === 0) {
        resolve(false);
        return;
      }

      const officeExtensions = [".docx", ".xlsx", ".pptx"];
      // crea la lista dei soli allegati con estensione Office
      const officeAttachments = result.value.filter(a =>
        officeExtensions.some(ext => a.name.toLowerCase().endsWith(ext))
      );
      console.log("office attachments :", officeAttachments);
      // Se non ci sono allegati Office, risolve la promise con false e ritorna (dal callback getAttachmentsAsync non dalla funzione checkReservedAttachments)
      if (officeAttachments.length === 0) {
        resolve(false);
        return;
      }
      // Cicla su tutti gli allegati Office, ma al primo che è reserved risolve la promise con true e ritorna (dal callback getAttachmentsAsync non dalla funzione checkReservedAttachments)
      for (const attachment of officeAttachments) {
        const isReserved = await isAttachmentReserved(item, attachment.id);
        if (isReserved) {
          resolve(true);
          return;
        }
      }
      // se nel ciclo sopra non ha trovato allegati riservati, risolve la promise con false e ritorna (dal callback getAttachmentsAsync non dalla funzione checkReservedAttachments)
      console.log("no reserved attachments found");
      resolve(false);
      return;
    });
  });
}

async function isAttachmentReserved(item: Office.MessageCompose, attachmentId: string): Promise<boolean> {
  return new Promise((resolve) => {
    item.getAttachmentContentAsync(attachmentId, async (result) => {
      console.log("getAttachmentContentAsync result:", result);
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("getAttachmentContentAsync failed:", result.error);
        resolve(false);
        return;
      }

      try {
        const rawContent = result.value.content;
        // Rimuove eventuali prefissi data-url e spazi/a capo dal base64
        const cleanBase64 = rawContent.replace(/^data:.*;base64,/, "").replace(/\s/g, "");

        // Carica lo zip direttamente dalla stringa base64 tramite JSZip
        const zip = await JSZip.loadAsync(cleanBase64, { base64: true });

        // Cerca docProps/custom.xml in modo case-insensitive e tollerante ai separatori / e \
        const customXmlFiles = zip.file(/docProps[\/\\]custom\.xml$/i);
        const customXmlFile = customXmlFiles.length > 0 ? customXmlFiles[0] : null;

        console.log("customXmlFile :", customXmlFile);
        if (!customXmlFile) {
          console.log("customXmlFile is null (non trovato nello zip)");
          resolve(false);
          return;
        }

        const xmlContent = await customXmlFile.async("string");  // metadati dell'allegato in formato stringa
        console.log("xmlContent :", xmlContent);

        // ricerca stringa "reserved" usando una regex tollerante per l'attributo name="reserved" o name='reserved'
        if (/name=["']reserved["']/i.test(xmlContent)) {
          console.log("Found 'reserved' in metadata");
          resolve(true);
          return;
        }

        resolve(false);

      } catch (error) {
        console.error("isAttachmentReserved error: ", error);
        resolve(false);
      }
    });
  });
}


/**
 * Aggiunge un timeout ad una promise
 * 
 * @param promise La promise in input
 * @param ms Timeout in millisecondi
 * @returns La nuova promise con timeout
 */
function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`Timeout after ${ms}ms`)), ms);
    // se non è scaduto il timeout, ritorna il result di promise oppure il suo errore, e cancella il timeout
    promise
      .then((result) => { clearTimeout(timer); resolve(result); })
      .catch((error) => { clearTimeout(timer); reject(error); });
  });
}