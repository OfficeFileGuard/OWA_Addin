import JSZip from "jszip";

Office.onReady(() => {});

// Starting point quando la mail viene inviata
async function onMessageSendHandler(event: Office.AddinCommands.Event) {
  // Ottiene item come oggetto email in composizione 
  const item = Office.context.mailbox.item as Office.MessageCompose;
  // se non l'ottine conste l'invio e ritorna
  if (!item) {
    event.completed({ allowEvent: true });
    return;
  }

  try {
    // Verifica se ci sono destinatari esterni. Se non ci sono permette invio e esce
    const hasExternalRecipients = await checkExternalRecipients(item);
    if (!hasExternalRecipients) {
      event.completed({ allowEvent: true });
      return;
    }

    // Verifica se ci sono allegati riservati. Se non ci sono permette invio e esce
    const hasReservedAttachments = await checkReservedAttachments(item);
    if (!hasReservedAttachments) {
      event.completed({ allowEvent: true });
      return;
    }

    // Entrambe le condizioni vere — apri il dialog e attendi la risposta
    await new Promise<void>((resolve) => {   //Promise attende la risposta del dialog
      Office.context.ui.displayDialogAsync(
        "${ADDIN_BASE_URL}/guardalert.html",    //ADDIN_BASE_URL è una variabile definita da un plugin di webpack.config.js
        { height: 50, width: 30, displayInIframe: false },
        (asyncResult) => {
          // Se il dialog non si apre, permette l'invio della mail e esce
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            event.completed({ allowEvent: true });
            resolve();  // risolve la Promise per terminare l'attesa
            return;
          }

          const dialog = asyncResult.value;     // dialog contiene l'oggetto dialog

          // creazione di un event handler di tipo DialogMessageReceived (viene attivato quando il dialog invia un messaggio)
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (args: { message: string; origin: string | undefined; } | { error: number; }) => {
              // se non c'è il messaggio, chiude l'evento, risolve la Promise e esce
              if (!('message' in args)) {
                console.warn('DialogMessageReceived event triggered without a "message" property. Proceeding with send');
                event.completed({ allowEvent: true });
                resolve();
                return;
              }

            // se invece c'è il messaggio
              dialog.close();   // chiude il dialog
              // se il messaggio è "send", chiude l'evento e permette l'invio della mail
              if (args.message === "send") {
                event.completed({ allowEvent: true });  // addin ha terminato il suo intervento e permette l'invio della mail
              } 
              // altrimenti chiude l'evento e blocca l'invio della mail e mostra un messaggio di errore
              else {
                (event as any).completed({ allowEvent: false, errorMessage: "⚠ This email contains confidential attachments addressed to external recipients. Please review before sending." });
              }
              resolve();
            }
          );

          dialog.addEventHandler(
            Office.EventType.DialogEventReceived,
            () => {
              (event as any).completed({ allowEvent: false, errorMessage: "⚠ This email contains confidential attachments addressed to external recipients. Please review before sending." });
              resolve();
            }
          );
        }
      );
    });

  } catch {
    // In caso di errore, permette l'invio della mail e esce
    event.completed({ allowEvent: true });
    return;
  }
}

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

  const [to, cc, bcc] = await Promise.all([
    getRecipients(item.to),
    getRecipients(item.cc),
    getRecipients(item.bcc)
  ]);

  const allRecipients = [...to, ...cc, ...bcc];
  return allRecipients.some(r => {
    const domain = r.emailAddress.split("@")[1]?.toLowerCase() ?? "";
    return domain !== senderDomain && domain !== "";
  });
}

async function checkReservedAttachments(item: Office.MessageCompose): Promise<boolean> {
  return new Promise((resolve) => {
    item.getAttachmentsAsync(async (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || result.value.length === 0) {
        resolve(false);
        return;
      }

      const officeExtensions = [".docx", ".xlsx", ".pptx"];
      const officeAttachments = result.value.filter(a =>
        officeExtensions.some(ext => a.name.toLowerCase().endsWith(ext))
      );

      if (officeAttachments.length === 0) {
        resolve(false);
        return;
      }

      for (const attachment of officeAttachments) {
        const isReserved = await isAttachmentReserved(item, attachment.id);
        if (isReserved) {
          resolve(true);
          return;
        }
      }

      resolve(false);
    });
  });
}

async function isAttachmentReserved(item: Office.MessageCompose, attachmentId: string): Promise<boolean> {
  return new Promise((resolve) => {
    item.getAttachmentContentAsync(attachmentId, async (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        resolve(false);
        return;
      }

      try {
        const base64 = result.value.content;
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++)
          bytes[i] = binary.charCodeAt(i);

        const zip = await JSZip.loadAsync(bytes.buffer);
        const customXmlFile = zip.file("docProps/custom.xml");

        if (!customXmlFile) {
          resolve(false);
          return;
        }

        const xmlContent = await customXmlFile.async("string");
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlContent, "application/xml");
        const properties = doc.getElementsByTagName("property");

        for (let i = 0; i < properties.length; i++) {
          if (properties[i].getAttribute("name")?.toLowerCase() === "reserved") {
            resolve(true);
            return;
          }
        }

        resolve(false);
      } catch {
        resolve(false);
      }
    });
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);