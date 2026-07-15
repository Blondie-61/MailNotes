/* global Office, document, window, navigator */

const AgentUrl = "/api";

const DEBUG = true;

let itemChangeSequence = 0;
let currentMailNotesId = "";
let shlStatusResetTimer: number | undefined;

function log(...args: any[]) {
  if (DEBUG) {
    console.log(...args);
  }
}

Office.onReady((info) => {
  log("Office.onReady", info);

  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  Office.context.mailbox.addHandlerAsync(
    Office.EventType.ItemChanged,
    handleItemChanged,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(
          "ItemChanged konnte nicht registriert werden:",
          result.error
        );

        return;
      }

      log("ItemChanged wurde registriert.");
      setShlStatus("active");
    }
  );

  void runOutlook();
});

async function runOutlook() {
  setupButtons();

  const sequence = ++itemChangeSequence;

  clearCurrentMailDisplay();
  showMailInformation();
  await loadNote(sequence);
}

async function handleItemChanged() {
  const sequence = ++itemChangeSequence;

  log("ItemChanged", sequence);

  clearCurrentMailDisplay();
  showMailInformation();
  await loadNote(sequence);
}

type CurrentMailIdentity = {
  messageId: string;
  itemId: string;
  conversationId: string;
  mailboxAddress: string;
  subject: string;
  senderName: string;
  senderAddress: string;
  mailDate: string;
};

function getCurrentMailSnapshot(): CurrentMailIdentity | null {
  const item = Office.context.mailbox.item;

  if (!item) {
    return null;
  }

  const messageId =
    ((item as any).internetMessageId || "").toString();

  if (!messageId) {
    return null;
  }

  return {
    messageId,
    itemId: ((item as any).itemId || "").toString(),
    conversationId:
      ((item as any).conversationId || "").toString(),
    mailboxAddress:
      (Office.context.mailbox.userProfile?.emailAddress || "").toString(),
    subject: (item.subject || "").toString(),
    senderName:
      ((item as any).from?.displayName || "").toString(),
    senderAddress:
      ((item as any).from?.emailAddress || "").toString(),
    mailDate:
      ((item as any).dateTimeCreated || "").toString()
  };
}

async function refreshKnownMailIdentity(
  identity: CurrentMailIdentity,
  showRepairStatus: boolean
): Promise<any> {
  const body = new URLSearchParams();

  if (currentMailNotesId) {
    body.append("mailNotesId", currentMailNotesId);
  }

  body.append("messageId", identity.messageId);
  body.append("itemId", identity.itemId);
  body.append("conversationId", identity.conversationId);
  body.append("mailboxAddress", identity.mailboxAddress);
  body.append("subject", identity.subject);
  body.append("senderName", identity.senderName);
  body.append("senderAddress", identity.senderAddress);
  body.append("mailDate", identity.mailDate);

  const response = await fetch(
    AgentUrl + "/mail/refresh",
    {
      method: "POST",
      body
    }
  );

  if (!response.ok) {
    throw new Error(
      "Agent returned HTTP " + response.status
    );
  }

  const result = await response.json();

  if (result.found && result.mailNotesId) {
    currentMailNotesId = result.mailNotesId.toString();
  }

  if (showRepairStatus && result.updated) {
    setShlStatus("repaired");
  }

  return result;
}

function setShlStatus(
  state: "active" | "repaired" | "error"
) {
  const element = document.getElementById("shl-status");

  if (!element) {
    return;
  }

  if (shlStatusResetTimer !== undefined) {
    window.clearTimeout(shlStatusResetTimer);
    shlStatusResetTimer = undefined;
  }

  element.className = "shl-status " + state;

  if (state === "repaired") {
    element.textContent = "🟢 SHL aktiv · Mail aktualisiert";
    element.title =
      "Die technische Outlook-ID der verschobenen Mail wurde aktualisiert. Links und Backlinks sind wieder gültig.";

    shlStatusResetTimer = window.setTimeout(() => {
      setShlStatus("active");
    }, 3500);

    return;
  }

  if (state === "error") {
    element.textContent = "🟠 SHL gestört";
    element.title =
      "Die automatische Aktualisierung konnte den MailNotesAgent nicht erreichen.";
    return;
  }

  element.textContent = "🟢 SHL aktiv";
  element.title =
    "MailNotes aktualisiert bekannte Mails automatisch bei einem echten Outlook-Kontextwechsel.";
}

function clearCurrentMailDisplay() {
  currentMailNotesId = "";

  setText("mail-subject", "");
  setText("mail-from", "");
  setText("mail-date", "");

  setText("mail-message-id", "");
  setText("mail-conversation-id", "");
  setText("mail-item-id", "");

  setEditorText("note-content", "");
  setEditorText("note-links", "");

  setNoteMeta("", "");
  setText("mail-link-status", "");

  const linkInput =
    document.getElementById("link-input") as HTMLInputElement;

  if (linkInput) {
    linkInput.value = "";
  }

  const linksList =
    document.getElementById("links-list");

  if (linksList) {
    linksList.innerHTML = "";
  }

  clearBacklinks();
}

function setupButtons() {
  const btnSave =
    document.getElementById("btn-save");

  const btnAddLink =
    document.getElementById("btn-add-link");

  const btnCopyMailLink =
    document.getElementById("btn-copy-mail-link");

  const linkInput =
    document.getElementById("link-input") as HTMLInputElement;

  if (btnSave) {
    btnSave.onclick = async () => {
      try {
        await saveNote();
        await loadNote();
      } catch (error) {
        console.error(error);

        setText(
          "mail-link-status",
          "Speichern fehlgeschlagen."
        );
      }
    };
  }

  if (btnAddLink) {
    btnAddLink.onclick = () => {
    void addLinkFromInput();
    };
  }

  if (btnCopyMailLink) {
    btnCopyMailLink.onclick = () => {
      void rememberCurrentMailLink();
    };
  }

if (linkInput) {
  linkInput.onkeydown = (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      void addLinkFromInput();
    }
  };
}

}

function showMailInformation() {
  const item =
    Office.context.mailbox.item;

  if (!item) {
    setText("mail-subject", "");
    setText("mail-from", "");
    setText("mail-date", "");

    setText("mail-message-id", "");
    setText("mail-conversation-id", "");
    setText("mail-item-id", "");

    return;
  }

  setText(
    "mail-subject",
    item.subject
  );

  setText(
    "mail-from",
    (item as any).from?.displayName
  );

  setText(
    "mail-date",
    formatDate(
      (item as any).dateTimeCreated
    )
  );

  setText(
    "mail-message-id",
    (item as any).internetMessageId
  );

  setText(
    "mail-conversation-id",
    (item as any).conversationId
  );

  setText(
    "mail-item-id",
    (item as any).itemId
  );
}

async function loadNote(
  expectedSequence: number = itemChangeSequence
) {
  const item =
    Office.context.mailbox.item;

  if (!item) {
    clearCurrentMailDisplay();
    return;
  }

  const messageId =
    (item as any).internetMessageId;

  if (!messageId) {
    setEditorText(
      "note-content",
      "Keine Message-ID vorhanden."
    );

    setEditorText(
      "note-links",
      ""
    );

    setNoteMeta("", "");

    await renderLinks(expectedSequence);
    clearBacklinks();

    return;
  }

  try {
    const note =
      await getNote(messageId);

    if (expectedSequence !== itemChangeSequence) {
      log(
        "Veraltete Notizantwort verworfen:",
        messageId
      );

      return;
    }

    const currentItem =
      Office.context.mailbox.item;

        const currentMessageId =
      currentItem
        ? (currentItem as any).internetMessageId
        : "";

    if (currentMessageId !== messageId) {
      log(
        "Mail wurde während des Ladens gewechselt:",
        messageId
      );

      return;
    }

    // Die MailNotesID wird auch für bekannte Mails ohne eigene Notiz
    // vom Agent geliefert. Erst danach darf SHL die technischen IDs
    // aktualisieren.
    currentMailNotesId =
      note.mailNotesId || "";

    const identity = getCurrentMailSnapshot();

    // /mail/refresh dient zugleich als Ensure-Mail-Aufruf: Eine bisher
    // unbekannte Mail wird in der Mail-Tabelle registriert und erhält eine
    // MailNotesID. Eine leere Notiz wird dabei ausdrücklich nicht erzeugt.
    if (identity) {
      try {
        await refreshKnownMailIdentity(identity, true);
      } catch (refreshError) {
        log("SHL refresh fehlgeschlagen:", refreshError);
        setShlStatus("error");
      }
    }

    if (note.found) {
      setEditorText(
        "note-content",
        note.content
      );

      setEditorText(
        "note-links",
        note.links
      );

      setNoteMeta(
        note.createdAt,
        note.modifiedAt
      );
    } else {
      setEditorText(
        "note-content",
        ""
      );

      setEditorText(
        "note-links",
        ""
      );

      setNoteMeta("", "");
    }

    await renderLinks(expectedSequence);

    if (expectedSequence !== itemChangeSequence) {
      return;
    }

    await renderBacklinks(
      currentMailNotesId || messageId,
      expectedSequence
    );

  } catch (error) {
    if (expectedSequence !== itemChangeSequence) {
      return;
    }

    setEditorText(
      "note-content",
      "MailNotesAgent nicht erreichbar."
    );

    setEditorText(
      "note-links",
      ""
    );

    setNoteMeta("", "");

    await renderLinks(expectedSequence);
    clearBacklinks();

    console.error(error);
  }
}

async function saveNote() {
  const item =
    Office.context.mailbox.item;

  if (!item) {
    throw new Error(
      "Keine Mail ausgewählt."
    );
  }

  const messageId =
    (item as any).internetMessageId;

  const conversationId =
    (item as any).conversationId;

  const itemId =
    (item as any).itemId || "";

  const subject =
    item.subject || "";

  const senderName =
    (item as any).from?.displayName || "";

  const senderAddress =
    (item as any).from?.emailAddress || "";

  const mailboxAddress =
    Office.context.mailbox.userProfile?.emailAddress || "";

  const mailDate =
    (item as any).dateTimeCreated || "";

  const content =
    getEditorText("note-content");

  const links =
    getEditorText("note-links");

  if (!messageId) {
    throw new Error(
      "Keine Message-ID vorhanden."
    );
  }

  const body =
    new URLSearchParams();

  if (currentMailNotesId) {
    body.append(
      "mailNotesId",
      currentMailNotesId
    );
  }

  body.append(
    "messageId",
    messageId
  );

  body.append(
    "conversationId",
    conversationId || ""
  );

  body.append(
    "itemId",
    itemId
  );

  body.append(
    "subject",
    subject
  );

  body.append(
    "senderName",
    senderName
  );

  body.append(
    "senderAddress",
    senderAddress
  );

  body.append(
    "mailboxAddress",
    mailboxAddress
  );

  body.append(
    "mailDate",
    mailDate
  );

  body.append(
    "content",
    content
  );

  body.append(
    "links",
    links
  );

  log(
    "POST body:",
    body.toString()
  );

  const response =
    await fetch(
      AgentUrl + "/note",
      {
        method: "POST",
        body: body
      }
    );

  const responseText =
    await response.text();

  log(
    "SAVE status:",
    response.status
  );

  log(
    "SAVE response:",
    responseText
  );

  if (!response.ok) {
    throw new Error(
      "Agent returned HTTP " +
      response.status
    );
  }

  const result = JSON.parse(responseText);

  if (result.mailNotesId) {
    currentMailNotesId =
      result.mailNotesId.toString();
  }

  return result;
}

async function getNote(
  messageId: string
): Promise<any> {
  const url =
    AgentUrl +
    "/note?messageId=" +
    encodeURIComponent(messageId);

  log(
    "GET note URL:",
    url
  );

  const response =
    await fetch(url);

  const responseText =
    await response.text();

  log(
    "GET status:",
    response.status
  );

  log(
    "GET response:",
    responseText
  );

  if (!response.ok) {
    throw new Error(
      "Agent returned HTTP " +
      response.status
    );
  }

  return JSON.parse(responseText);
}


async function addLinkFromInput() {
  const input =
    document.getElementById(
      "link-input"
    ) as HTMLInputElement;

  if (!input) {
    return;
  }

  let value =
    input.value.trim();

  if (!value) {
    const bufferedLink =
      await getBufferedMailLink();

    if (!bufferedLink) {
      return;
    }

    value = bufferedLink;
  }

  const links =
    getLinks();

  if (links.includes(value)) {
    setText(
      "mail-link-status",
      "Dieser Link ist bereits vorhanden."
    );

    return;
  }

  links.push(value);

  setLinks(links);

  input.value = "";

  await renderLinks();

  try {
    await saveNote();
    await loadNote();

    setText(
      "mail-link-status",
      "Link hinzugefügt und gespeichert."
    );

    window.setTimeout(() => {
      setText(
        "mail-link-status",
        ""
      );
    }, 2000);

  } catch (error) {
    setText(
      "mail-link-status",
      "Link hinzugefügt, Speichern fehlgeschlagen."
    );

    console.error(error);
  }
}

async function rememberCurrentMailLink() {
  const item =
    Office.context.mailbox.item;

  if (!item) {
    setText(
      "mail-link-status",
      "Keine Mail ausgewählt."
    );

    return;
  }

  const messageId =
    (item as any).internetMessageId;

  if (!messageId) {
    setText(
      "mail-link-status",
      "Keine Message-ID vorhanden."
    );

    return;
  }

  const itemId =
    (item as any).itemId || "";

  const conversationId =
    (item as any).conversationId || "";

  const subject =
    item.subject || "";

  const senderName =
    (item as any).from?.displayName || "";

  const senderAddress =
    (item as any).from?.emailAddress || "";

  const mailDate =
    (item as any).dateTimeCreated || "";

  try {
    setText(
      "mail-link-status",
      "Mail-Link wird gemerkt …"
    );

    // Der LinkBuffer registriert eine bisher unbekannte Mail selbst
    // in der Mail-Tabelle und erzeugt dabei ihre MailNotesID. Eine
    // sichtbare oder leere Notiz ist dafür ausdrücklich nicht nötig.
    await setLinkBuffer({
      mailNotesId: currentMailNotesId || undefined,
      messageId,
      itemId,
      conversationId,
      subject,
      senderName,
      senderAddress,
      mailDate
    });

    setText(
      "mail-link-status",
      "Mail-Link gemerkt."
    );

    window.setTimeout(() => {
      setText(
        "mail-link-status",
        ""
      );
    }, 2500);

  } catch (error) {
    setText(
      "mail-link-status",
      "Mail-Link konnte nicht gemerkt werden."
    );

    console.error(error);
  }
}

async function renderLinks(
  expectedSequence: number = itemChangeSequence
) {
  const list =
    document.getElementById(
      "links-list"
    );

  if (!list) {
    return;
  }

  const links =
    getLinks();

  const fragment =
    document.createDocumentFragment();

  for (
    let index = 0;
    index < links.length;
    index++
  ) {
    if (expectedSequence !== itemChangeSequence) {
      return;
    }

    const url =
      links[index];

    const row =
      document.createElement("div");

    row.className =
      "link-row";

    const open =
      document.createElement("a");

    open.className =
      "link-open";

    open.href =
      url;

    open.title =
      url;

    const title =
      document.createElement("span");

    title.className =
      "link-title";

    title.textContent =
      getLinkCaption(url);

    const subtitle =
      document.createElement("span");

    subtitle.className =
      "link-subtitle";

    open.appendChild(title);
    open.appendChild(subtitle);

    open.onclick = (event) => {
      event.preventDefault();
      openLink(url);
    };

    const actions =
      document.createElement("div");

    actions.className =
      "link-actions";

    const copyButton =
      createCopyButton(
        url,
        "Link kopieren"
      );

    const deleteButton =
      document.createElement("button");

    deleteButton.className =
      "link-action";

    deleteButton.type =
      "button";

    deleteButton.title =
      "Link entfernen";

    deleteButton.innerHTML =
      '<span class="icon-delete">×</span>';

    deleteButton.onclick = async () => {
      deleteButton.disabled = true;
      await deleteLink(index);
    };

    actions.appendChild(
      copyButton
    );

    actions.appendChild(
      deleteButton
    );

    row.appendChild(open);
    row.appendChild(actions);

    fragment.appendChild(row);

    if (
      url
        .toLowerCase()
        .startsWith("mailnotes:")
    ) {
      await resolveRenderedMailLink(
        url,
        title,
        subtitle,
        expectedSequence
      );
    }
  }

  if (expectedSequence !== itemChangeSequence) {
    return;
  }

  list.replaceChildren(fragment);
}

async function resolveRenderedMailLink(
  url: string,
  title: HTMLElement,
  subtitle: HTMLElement,
  expectedSequence: number
) {
  try {
    const resolved =
      await resolveLink(url);

    if (expectedSequence !== itemChangeSequence) {
      return;
    }

    if (
      resolved.found &&
      resolved.type === "mail"
    ) {
      title.textContent =
        "📧 " +
        (
          resolved.title ||
          "Verknüpfte Mail"
        );

      subtitle.textContent =
        formatResolvedSubtitle(
          resolved.subtitle
        );

      return;
    }

    subtitle.textContent =
      "Mail konnte nicht aufgelöst werden.";

  } catch (error) {
    subtitle.textContent =
      "MailNotesAgent nicht erreichbar.";

    console.error(error);
  }
}

function formatResolvedSubtitle(
  value: any
): string {
  if (!value) {
    return "";
  }

  const text =
    value.toString();

  const separatorPosition =
    text.indexOf(" · ");

  if (separatorPosition < 0) {
    return text;
  }

  const senderName =
    text
      .substring(
        0,
        separatorPosition
      )
      .trim();

  const mailDate =
    text
      .substring(
        separatorPosition + 3
      )
      .trim();

  const parts: string[] = [];

  if (senderName) {
    parts.push(senderName);
  }

  if (mailDate) {
    parts.push(
      formatDate(mailDate)
    );
  }

  return parts.join(" · ");
}

async function getBacklinks(
  mailIdentity: string
): Promise<any> {
  const isMailNotesId =
    isMailNotesIdValue(mailIdentity);

  const parameterName =
    isMailNotesId
      ? "mailNotesId"
      : "messageId";

  const url =
    AgentUrl +
    "/backlinks?" +
    parameterName +
    "=" +
    encodeURIComponent(mailIdentity);

  log(
    "GET backlinks URL:",
    url
  );

  const response =
    await fetch(url);

  const responseText =
    await response.text();

  log(
    "BACKLINKS status:",
    response.status
  );

  log(
    "BACKLINKS response:",
    responseText
  );

  if (!response.ok) {
    throw new Error(
      "Agent returned HTTP " +
      response.status
    );
  }

  return JSON.parse(responseText);
}

async function renderBacklinks(
  mailIdentity: string,
  expectedSequence: number = itemChangeSequence
) {
  const section =
    document.getElementById(
      "backlinks-section"
    );

  const list =
    document.getElementById(
      "backlinks-list"
    );

  if (!section || !list) {
    return;
  }

  try {
    const result =
      await getBacklinks(mailIdentity);

    if (
      expectedSequence !== itemChangeSequence ||
      getCurrentMailIdentity() !== mailIdentity
    ) {
      return;
    }

    const items =
      Array.isArray(result.items)
        ? result.items
        : [];

    const fragment =
      document.createDocumentFragment();

    for (const item of items) {
      const sourceMailIdentity =
        item.mailNotesId ||
        item.messageId ||
        "";

      const mailLink =
        "mailnotes:" +
        encodeURIComponent(
          sourceMailIdentity
        );

      const row =
        document.createElement("div");

      row.className =
        "link-row backlink-row";

      const open =
        document.createElement("a");

      open.className =
        "link-open";

      open.href =
        mailLink;

      open.title =
        item.subject || sourceMailIdentity;

      const title =
        document.createElement("span");

      title.className =
        "link-title";

      title.textContent =
        "📧 " +
        (
          item.subject ||
          "Verknüpfte Mail"
        );

      const subtitle =
        document.createElement("span");

      subtitle.className =
        "link-subtitle";

      const subtitleParts: string[] = [];

      if (item.senderName) {
        subtitleParts.push(
          item.senderName
        );
      }

      if (item.mailDate) {
        subtitleParts.push(
          formatDate(item.mailDate)
        );
      }

      subtitle.textContent =
        subtitleParts.join(" · ");

      open.appendChild(title);
      open.appendChild(subtitle);

      open.onclick = (event) => {
        event.preventDefault();
        openLink(mailLink);
      };

      const actions =
        document.createElement("div");

      actions.className =
        "link-actions";

      const copyButton =
        createCopyButton(
          mailLink,
          "Mail-Link kopieren"
        );

      actions.appendChild(
        copyButton
      );

      row.appendChild(open);
      row.appendChild(actions);

      fragment.appendChild(row);
    }

    if (
      expectedSequence !== itemChangeSequence ||
      getCurrentMailIdentity() !== mailIdentity
    ) {
      return;
    }

    list.replaceChildren(fragment);
    section.hidden = items.length === 0;

  } catch (error) {
    if (expectedSequence !== itemChangeSequence) {
      return;
    }

    clearBacklinks();
    console.error(error);
  }
}

function clearBacklinks() {
  const section =
    document.getElementById(
      "backlinks-section"
    );

  const list =
    document.getElementById(
      "backlinks-list"
    );

  if (list) {
    list.innerHTML = "";
  }

  if (section) {
    section.hidden = true;
  }
}

function createCopyButton(
  value: string,
  title: string
): HTMLButtonElement {
  const button =
    document.createElement("button");

  button.className =
    "link-action";

  button.type =
    "button";

  button.title =
    title;

  button.innerHTML =
    '<span class="icon-copy"></span>';

  button.onclick = () => {
    void copyLink(value);
  };

  return button;
}

function getLinks(): string[] {
  return getEditorText(
    "note-links"
  )
    .split(/\r?\n/)
    .map((value) => value.trim())
    .filter((value) => value.length > 0);
}

function setLinks(
  links: string[]
) {
  setEditorText(
    "note-links",
    links.join("\n")
  );
}

function openLink(
  url: string
) {
  if (
    url
      .toLowerCase()
      .startsWith("mailnotes:")
  ) {
    void openMailNotesLink(url);
    return;
  }

  window.open(
    url,
    "_blank"
  );
}

async function openMailNotesLink(
  url: string
) {
  try {
    setText(
      "mail-link-status",
      "Mail wird geöffnet …"
    );

    const resolved =
      await resolveLink(url);

    if (
      !resolved.found ||
      resolved.type !== "mail"
    ) {
      await copyMessageIdFallback(
        url,
        "Mail konnte nicht aufgelöst werden."
      );

      return;
    }

    const storedItemId =
      resolved.itemId
        ? resolved.itemId.toString()
        : "";

    if (!storedItemId) {
      await copyMessageIdFallback(
        url,
        "Keine Item-ID vorhanden."
      );

      return;
    }

    const mailbox =
      Office.context.mailbox as any;

    const candidates: string[] = [storedItemId];

    // displayMessageFormAsync erwartet je nach Outlook-Client eine
    // Exchange-/EWS-ID. Falls im Datensatz eine REST-ID gelandet ist,
    // versuchen wir deshalb zusätzlich die konvertierte EWS-ID.
    if (
      typeof mailbox.convertToEwsId ===
      "function"
    ) {
      try {
        const ewsItemId =
          mailbox.convertToEwsId(
            storedItemId
          );

        if (
          ewsItemId &&
          !candidates.includes(ewsItemId)
        ) {
          candidates.push(ewsItemId);
        }
      } catch (conversionError) {
        log(
          "Item-ID konnte nicht in EWS-ID konvertiert werden:",
          conversionError
        );
      }
    }

    log(
      "Öffne Mail mit Item-ID-Kandidaten:",
      candidates
    );

    if (
      typeof mailbox.displayMessageFormAsync ===
      "function"
    ) {
      const tryCandidate =
        (index: number) => {
          if (index >= candidates.length) {
            setText(
              "mail-link-status",
              "Öffnen fehlgeschlagen – gespeicherte Item-ID ist ungültig."
            );
            return;
          }

          const candidate = candidates[index];

          mailbox.displayMessageFormAsync(
            candidate,
            (result: Office.AsyncResult<void>) => {
              log(
                "displayMessageFormAsync:",
                index,
                result
              );

              if (
                result.status ===
                Office.AsyncResultStatus.Failed
              ) {
                console.error(
                  "Mail konnte mit Item-ID-Kandidat nicht geöffnet werden:",
                  candidate,
                  result.error
                );

                tryCandidate(index + 1);
                return;
              }

              setText(
                "mail-link-status",
                ""
              );
            }
          );
        };

      tryCandidate(0);
      return;
    }

    if (
      typeof mailbox.displayMessageForm ===
      "function"
    ) {
      mailbox.displayMessageForm(
        candidates[0]
      );

      setText(
        "mail-link-status",
        ""
      );

      return;
    }

    await copyMessageIdFallback(
      url,
      "Öffnen wird von Outlook nicht unterstützt."
    );

  } catch (error) {
    console.error(error);

    await copyMessageIdFallback(
      url,
      "Öffnen fehlgeschlagen."
    );
  }
}

async function copyMessageIdFallback(
  _url: string,
  statusText: string
) {
  setText(
    "mail-link-status",
    statusText.replace(
      " – Message-ID wurde kopiert.",
      "."
    )
  );

  window.setTimeout(() => {
    setText(
      "mail-link-status",
      ""
    );
  }, 5000);
}

function getCurrentMailIdentity(): string {
  return (
    currentMailNotesId ||
    getCurrentMessageId()
  );
}

function isMailNotesIdValue(
  value: string
): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    value
  );
}

function decodeMailNotesMessageId(
  url: string
): string {
  const raw =
    url.substring(
      "mailnotes:".length
    );

  try {
    return decodeURIComponent(raw);
  } catch {
    return raw;
  }
}

async function copyLink(
  url: string
) {
  try {
    await navigator.clipboard.writeText(
      url
    );
  } catch (error) {
    console.error(error);
  }
}

type LinkBufferData = {
  mailNotesId?: string;
  messageId?: string;
  itemId?: string;
  conversationId?: string;
  subject?: string;
  senderName?: string;
  senderAddress?: string;
  mailDate?: string;
};

async function setLinkBuffer(
  data: LinkBufferData
) {
  const body =
    new URLSearchParams();

  if (data.mailNotesId) {
    body.append("mailNotesId", data.mailNotesId);
  }
  body.append("messageId", data.messageId || "");
  body.append("itemId", data.itemId || "");
  body.append("conversationId", data.conversationId || "");
  body.append("subject", data.subject || "");
  body.append("senderName", data.senderName || "");
  body.append("senderAddress", data.senderAddress || "");
  body.append("mailDate", data.mailDate || "");

  const response =
    await fetch(
      AgentUrl + "/linkbuffer",
      {
        method: "POST",
        body
      }
    );

  if (!response.ok) {
    throw new Error(
      "Agent returned HTTP " +
      response.status
    );
  }
}

async function getBufferedMailLink(): Promise<string> {
  try {
    setText(
      "mail-link-status",
      "Gemerkter Mail-Link wird geladen …"
    );

    const response =
      await fetch(
        AgentUrl + "/linkbuffer"
      );

    if (!response.ok) {
      throw new Error(
        "Agent returned HTTP " +
        response.status
      );
    }

    const buffer =
      await response.json();

    if (!buffer.found || !buffer.mailNotesId) {
      setText(
        "mail-link-status",
        "Kein Mail-Link gemerkt."
      );

      return "";
    }

    return (
      "mailnotes:" +
      encodeURIComponent(
        buffer.mailNotesId.toString()
      )
    );
  } catch (error) {
    setText(
      "mail-link-status",
      "Gemerkter Mail-Link konnte nicht geladen werden."
    );

    console.error(error);
    return "";
  }
}

async function resolveLink(
  url: string
): Promise<any> {
  const response =
    await fetch(
      AgentUrl +
      "/resolve?link=" +
      encodeURIComponent(url)
    );

  if (!response.ok) {
    throw new Error(
      "Agent returned HTTP " +
      response.status
    );
  }

  return await response.json();
}

async function deleteLink(
  index: number
) {
  const expectedSequence =
    itemChangeSequence;

  const links =
    getLinks();

  const deletedLinks =
    links.splice(
      index,
      1
    );

  if (deletedLinks.length === 0) {
    return;
  }

  setLinks(links);

  await renderLinks(
    expectedSequence
  );

  try {
    await saveNote();

    if (
      expectedSequence !==
      itemChangeSequence
    ) {
      return;
    }

    setText(
      "mail-link-status",
      "Link gelöscht."
    );

    window.setTimeout(() => {
      if (
        expectedSequence ===
        itemChangeSequence
      ) {
        setText(
          "mail-link-status",
          ""
        );
      }
    }, 2000);

  } catch (error) {
    if (
      expectedSequence !==
      itemChangeSequence
    ) {
      console.error(error);
      return;
    }

    const restoredLinks =
      getLinks();

    restoredLinks.splice(
      Math.min(
        index,
        restoredLinks.length
      ),
      0,
      deletedLinks[0]
    );

    setLinks(restoredLinks);

    await renderLinks(
      expectedSequence
    );

    setText(
      "mail-link-status",
      "Löschen konnte nicht gespeichert werden."
    );

    console.error(error);
  }
}

function getLinkCaption(
  url: string
): string {
  try {
    const lower =
      url.toLowerCase();

    if (
      lower.startsWith("mailnotes:")
    ) {
      return "📧 Verknüpfte Mail";
    }

    if (
      lower.startsWith("hook://")
    ) {
      return "🔗 Hookmark";
    }

    if (
      lower.startsWith("file://")
    ) {
      const parts =
        url.split("/");

      const fileName =
        parts[parts.length - 1];

      return (
        "📄 " +
        decodeURIComponent(
          fileName || url
        )
      );
    }

    if (
      lower.startsWith("http://") ||
      lower.startsWith("https://")
    ) {
      const parsed =
        new URL(url);

      return (
        "🌐 " +
        parsed.hostname
      );
    }

    if (
      lower.startsWith("mailto:")
    ) {
      return (
        "✉️ " +
        url.substring(7)
      );
    }

    return "🔗 " + url;

  } catch {
    return "🔗 " + url;
  }
}

function getCurrentMessageId(): string {
  const item =
    Office.context.mailbox.item;

  return item
    ? (item as any).internetMessageId || ""
    : "";
}

function getEditorText(
  id: string
): string {
  const element =
    document.getElementById(
      id
    ) as HTMLTextAreaElement;

  if (!element) {
    return "";
  }

  return element.value;
}

function setEditorText(
  id: string,
  value: any
) {
  const element =
    document.getElementById(
      id
    ) as HTMLTextAreaElement;

  if (!element) {
    return;
  }

  element.value =
    value
      ? value.toString()
      : "";
}

function setNoteMeta(
  createdAt: any,
  modifiedAt: any
) {
  setText(
    "note-created",
    formatDate(createdAt)
  );

  setText(
    "note-modified",
    formatDate(modifiedAt)
  );
}

function formatDate(
  value: any
): string {
  if (!value) {
    return "–";
  }

  const date =
    new Date(value);

  if (
    isNaN(date.getTime())
  ) {
    return value.toString();
  }

  return date
    .toLocaleString(
      "de-DE",
      {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit"
      }
    )
    .replace(",", "");
}

function setText(
  id: string,
  value: any
) {
  const element =
    document.getElementById(id);

  if (!element) {
    return;
  }

  element.textContent =
    value
      ? value.toString()
      : "–";
}