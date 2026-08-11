/* global Office, document, window, navigator */

import { MailNotesConfig } from "../config";

const AgentUrl = MailNotesConfig.agentUrl;
const ENABLE_LOGGING = MailNotesConfig.enableLogging;

let itemChangeSequence = 0;
let currentMailNotesId = "";
let shlStatusResetTimer: number | undefined;
let repairQueueCount = 0;
let repairHintDismissed = false;
let searchTimer: number | undefined;
let searchSequence = 0;

const AUTOSAVE_DELAY_MS = 800;
let autosaveTimer: number | undefined;
let autosaveRevision = 0;
let autosaveCompletedRevision = 0;
let autosaveInProgress = false;
let autosaveActivePromise: Promise<boolean> | null = null;
let autosaveSuppressed = false;
let autosavePendingSnapshot: NoteSaveSnapshot | null = null;

function log(...args: any[]) {
  if (ENABLE_LOGGING) {
    console.log(...args);
  }
}

function fileLog(event: string, data?: unknown): void {
  if (!ENABLE_LOGGING) {
    return;
  }

  const body = new URLSearchParams();
  body.append("event", event);

  if (data !== undefined) {
    try {
      body.append("data", JSON.stringify(data));
    } catch {
      body.append("data", String(data));
    }
  }

  void fetch(AgentUrl + "/log", {
    method: "POST",
    body
  }).catch(() => {
    // Diagnose-Logging darf die eigentliche Funktion niemals beeinflussen.
  });
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
  void refreshRepairQueueNotice(true);

  const sequence = ++itemChangeSequence;

  clearCurrentMailDisplay();
  showMailInformation();
  await loadNote(sequence);
}

async function handleItemChanged() {
  await requestAutosave(true);
  const sequence = ++itemChangeSequence;

  log("ItemChanged", sequence);

  clearCurrentMailDisplay();
  showMailInformation();
  await loadNote(sequence);
  await refreshRepairQueueAfterMailChange();
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
    fileLog("SHL repaired", {
      mailNotesId: result.mailNotesId,
      oldItemId: result.oldItemId,
      itemId: result.itemId,
      messageId: identity.messageId,
      subject: identity.subject
    });
    setShlStatus("repaired");
  }

  return result;
}

function setShlStatus(
  state: "active" | "repaired" | "error" | "pending"
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
      setShlStatus(repairQueueCount > 0 ? "pending" : "active");
    }, 3500);

    return;
  }

  if (state === "pending") {
    element.textContent = repairQueueCount === 1
      ? "🟡 SHL aktiv · 1 Reparatur offen"
      : "🟡 SHL aktiv · " + repairQueueCount + " Reparaturen offen";
    element.title = repairQueueCount === 1
      ? "Ein MailLink wartet auf Reparatur. Klicken Sie auf den gelben Hinweis, sobald Sie die Reparatur fortsetzen möchten."
      : repairQueueCount + " MailLinks warten auf Reparatur.";
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
  cancelPendingAutosave();
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

  const btnOpenRepairQueue = document.getElementById("btn-open-repair-queue");
  const btnDismissRepairQueue = document.getElementById("btn-dismiss-repair-queue");
  const btnCloseRepairQueue = document.getElementById("btn-close-repair-queue");

  if (btnOpenRepairQueue) {
    btnOpenRepairQueue.onclick = () => void showRepairQueue();
  }

  if (btnDismissRepairQueue) {
    btnDismissRepairQueue.onclick = () => {
      repairHintDismissed = true;
      hideElement("repair-queue-notice");
      setShlStatus(repairQueueCount > 0 ? "pending" : "active");
    };
  }

  if (btnCloseRepairQueue) {
    btnCloseRepairQueue.onclick = () => {
      repairHintDismissed = true;
      hideElement("repair-queue-card");
      setShlStatus(repairQueueCount > 0 ? "pending" : "active");
    };
  }

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
  const noteContent =
    document.getElementById("note-content") as HTMLTextAreaElement;

  const btnAddLink =
    document.getElementById("btn-add-link");

  const btnCopyMailLink =
    document.getElementById("btn-copy-mail-link");

  const linkInput =
    document.getElementById("link-input") as HTMLInputElement;

  const searchInput =
    document.getElementById("search-input") as HTMLInputElement;

  const btnClearSearch =
    document.getElementById("btn-clear-search");

  if (noteContent) {
    noteContent.oninput = () => {
      scheduleAutosave();
    };

    noteContent.onblur = () => {
      void requestAutosave(true);
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

  if (searchInput) {
    searchInput.oninput = () => {
      scheduleSearch(searchInput.value);
    };

    searchInput.onkeydown = (event) => {
      if (event.key === "Escape") {
        clearSearch();
      }
    };
  }

  if (btnClearSearch) {
    btnClearSearch.onclick = clearSearch;
  }
}

type SearchResultItem = {
  mailNotesId: string;
  messageId: string;
  itemId: string;
  subject: string;
  senderName: string;
  senderAddress: string;
  mailDate: string;
  modifiedAt: string;
  snippet: string;
};

function scheduleSearch(searchText: string): void {
  if (searchTimer !== undefined) {
    window.clearTimeout(searchTimer);
  }

  const trimmed = searchText.trim();
  if (!trimmed) {
    clearSearchResults();
    return;
  }

  setText("search-status", "Suche …");
  searchTimer = window.setTimeout(() => {
    void searchNotes(trimmed);
  }, 250);
}

function clearSearch(): void {
  const input = document.getElementById("search-input") as HTMLInputElement;
  if (input) {
    input.value = "";
    input.focus();
  }

  if (searchTimer !== undefined) {
    window.clearTimeout(searchTimer);
    searchTimer = undefined;
  }

  searchSequence++;
  clearSearchResults();
}

function clearSearchResults(): void {
  const results = document.getElementById("search-results");
  if (results) {
    results.innerHTML = "";
    results.hidden = true;
  }
  setText("search-status", "");
}

async function searchNotes(searchText: string): Promise<void> {
  const sequence = ++searchSequence;

  try {
    const response = await fetch(
      AgentUrl + "/search?q=" + encodeURIComponent(searchText) + "&limit=50"
    );

    if (!response.ok) {
      throw new Error("Agent returned HTTP " + response.status);
    }

    const result = await response.json();
    if (sequence !== searchSequence) {
      return;
    }

    const items: SearchResultItem[] = Array.isArray(result.items)
      ? result.items
      : [];

    renderSearchResults(items);
  } catch (error) {
    console.error("MailNotes-Suche fehlgeschlagen:", error);
    if (sequence === searchSequence) {
      clearSearchResults();
      setText("search-status", "Suche nicht verfügbar.");
    }
  }
}

function renderSearchResults(items: SearchResultItem[]): void {
  const results = document.getElementById("search-results");
  if (!results) {
    return;
  }

  results.innerHTML = "";
  results.hidden = false;

  if (items.length === 0) {
    setText("search-status", "Keine Treffer.");
    return;
  }

  setText(
    "search-status",
    items.length === 1 ? "1 Treffer" : items.length + " Treffer"
  );

  const fragment = document.createDocumentFragment();

  for (const item of items) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "search-result";

    const subject = document.createElement("span");
    subject.className = "search-result-subject";
    subject.textContent = item.subject || "(Ohne Betreff)";
    button.appendChild(subject);

    const meta = document.createElement("span");
    meta.className = "search-result-meta";
    const sender = item.senderName || item.senderAddress || "Unbekannter Absender";
    meta.textContent = sender + (item.mailDate ? " · " + formatDate(item.mailDate) : "");
    button.appendChild(meta);

    if (item.snippet) {
      const snippet = document.createElement("span");
      snippet.className = "search-result-snippet";
      snippet.textContent = item.snippet;
      button.appendChild(snippet);
    }

    button.onclick = () => {
      const token = item.mailNotesId || item.messageId;
      if (token) {
        void openMailNotesLink("mailnotes:" + encodeURIComponent(token));
      }
    };

    fragment.appendChild(button);
  }

  results.appendChild(fragment);
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
        fileLog("mail/refresh error", {
          message: refreshError instanceof Error
            ? refreshError.message
            : String(refreshError)
        });
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

    autosaveCompletedRevision = autosaveRevision;
    setAutosaveStatus("");

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

type NoteSaveSnapshot = CurrentMailIdentity & {
  mailNotesId: string;
  content: string;
  links: string;
  sequence: number;
  revision: number;
};

function cancelPendingAutosave(): void {
  if (autosaveTimer !== undefined) {
    window.clearTimeout(autosaveTimer);
    autosaveTimer = undefined;
  }

  autosaveRevision++;
  autosaveCompletedRevision = autosaveRevision;
  autosavePendingSnapshot = null;
  setAutosaveStatus("");
}

function scheduleAutosave(): void {
  if (autosaveSuppressed) {
    return;
  }

  autosaveRevision++;
  autosavePendingSnapshot = createNoteSaveSnapshot(autosaveRevision);
  setAutosaveStatus("Nicht gespeichert");

  if (autosaveTimer !== undefined) {
    window.clearTimeout(autosaveTimer);
  }

  autosaveTimer = window.setTimeout(() => {
    autosaveTimer = undefined;
    void requestAutosave(false);
  }, AUTOSAVE_DELAY_MS);
}

async function requestAutosave(immediate: boolean): Promise<boolean> {
  if (autosaveSuppressed || autosaveRevision <= autosaveCompletedRevision) {
    return true;
  }

  if (immediate && autosaveTimer !== undefined) {
    window.clearTimeout(autosaveTimer);
    autosaveTimer = undefined;
  }

  if (autosaveActivePromise) {
    const previousResult = await autosaveActivePromise;

    if (!previousResult) {
      return false;
    }

    if (autosaveRevision <= autosaveCompletedRevision) {
      return true;
    }
  }

  autosaveActivePromise = drainAutosave();

  try {
    return await autosaveActivePromise;
  } finally {
    autosaveActivePromise = null;
  }
}

async function drainAutosave(): Promise<boolean> {
  autosaveInProgress = true;

  try {
    while (autosaveCompletedRevision < autosaveRevision) {
      const revision = autosaveRevision;
      const snapshot = autosavePendingSnapshot;

      if (!snapshot) {
        autosaveCompletedRevision = revision;
        return true;
      }

      setAutosaveStatus("Speichert …");

      try {
        const result = await saveNoteSnapshot(snapshot);
        autosaveCompletedRevision = revision;

        if (autosavePendingSnapshot?.revision === revision) {
          autosavePendingSnapshot = null;
        }

        if (snapshot.sequence === itemChangeSequence) {
          if (result.mailNotesId) {
            currentMailNotesId = result.mailNotesId.toString();
          }

          updateNoteMetaAfterSave(result);
          setAutosaveStatus(
            autosaveCompletedRevision < autosaveRevision
              ? "Speichert …"
              : "Gespeichert"
          );
        }
      } catch (error) {
        if (snapshot.sequence === itemChangeSequence) {
          setAutosaveStatus("Speichern fehlgeschlagen");
        }

        console.error(error);
        return false;
      }
    }

    return true;
  } finally {
    autosaveInProgress = false;
  }
}

function createNoteSaveSnapshot(revision: number): NoteSaveSnapshot | null {
  const identity = getCurrentMailSnapshot();

  if (!identity) {
    return null;
  }

  return {
    ...identity,
    mailNotesId: currentMailNotesId,
    content: getEditorText("note-content"),
    links: getEditorText("note-links"),
    sequence: itemChangeSequence,
    revision
  };
}


function updateNoteMetaAfterSave(result: any): void {
  const createdElement = document.getElementById("note-created");
  const currentCreated = createdElement?.textContent || "";
  const now = new Date().toISOString();

  if (result.createdAt) {
    setText("note-created", formatDate(result.createdAt));
  } else if (!currentCreated || currentCreated === "–") {
    setText("note-created", formatDate(now));
  }

  setText("note-modified", formatDate(result.modifiedAt || now));
}

function setAutosaveStatus(text: string): void {
  setText("autosave-status", text);
}

async function saveNoteSnapshot(snapshot: NoteSaveSnapshot): Promise<any> {
  const body = new URLSearchParams();

  if (snapshot.mailNotesId) {
    body.append("mailNotesId", snapshot.mailNotesId);
  }

  body.append("messageId", snapshot.messageId);
  body.append("conversationId", snapshot.conversationId);
  body.append("itemId", snapshot.itemId);
  body.append("subject", snapshot.subject);
  body.append("senderName", snapshot.senderName);
  body.append("senderAddress", snapshot.senderAddress);
  body.append("mailboxAddress", snapshot.mailboxAddress);
  body.append("mailDate", snapshot.mailDate);
  body.append("content", snapshot.content);
  body.append("links", snapshot.links);

  log("POST body:", body.toString());

  const response = await fetch(AgentUrl + "/note", {
    method: "POST",
    body
  });

  const responseText = await response.text();

  log("SAVE status:", response.status);
  log("SAVE response:", responseText);

  if (!response.ok) {
    throw new Error("Agent returned HTTP " + response.status);
  }

  return JSON.parse(responseText);
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
    autosaveRevision++;
    autosavePendingSnapshot = createNoteSaveSnapshot(autosaveRevision);
    const saved = await requestAutosave(true);

    setText(
      "mail-link-status",
      saved ? "Link hinzugefügt." : "Link hinzugefügt, Speichern fehlgeschlagen."
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
              "Öffnen fehlgeschlagen – Link wurde zur Reparatur vorgemerkt."
            );
            void addRepairQueueItem(resolved, storedItemId);
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
    autosaveRevision++;
    autosavePendingSnapshot = createNoteSaveSnapshot(autosaveRevision);
    const saved = await requestAutosave(true);

    if (!saved) {
      throw new Error("Link konnte nicht gespeichert werden.");
    }

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

type RepairQueueItem = {
  id: number;
  mailNotesId: string;
  oldItemId: string;
  messageId: string;
  subject: string;
  senderName: string;
  senderAddress: string;
  mailDate: string;
  retryCount: number;
  status: number;
};

async function addRepairQueueItem(resolved: any, oldItemId: string): Promise<void> {
  const body = new URLSearchParams();
  body.append("mailNotesId", (resolved.mailNotesId || "").toString());
  body.append("oldItemId", oldItemId);
  body.append("messageId", (resolved.messageId || "").toString());
  const queueMeta = parseResolvedRepairMetadata(resolved.subtitle);

  body.append("subject", (resolved.title || "").toString());
  body.append("senderName", queueMeta.senderName);
  body.append("mailDate", queueMeta.mailDate);
  body.append("reason", "stored_item_id_invalid");

  const response = await fetch(AgentUrl + "/repairqueue", { method: "POST", body });
  if (!response.ok) {
    throw new Error("Repair queue returned HTTP " + response.status);
  }
  repairHintDismissed = false;
  await refreshRepairQueueNotice(true);
}

async function refreshRepairQueueNotice(forceShow = false): Promise<number> {
  try {
    const response = await fetch(AgentUrl + "/repairqueue/count");
    if (!response.ok) return -1;

    const result = await response.json();
    const count = Number(result.count || 0);
    repairQueueCount = count;
    const notice = document.getElementById("repair-queue-notice");

    if (count === 0) {
      finishRepairQueue();
      return 0;
    }

    setShlStatus("pending");

    if (forceShow) {
      repairHintDismissed = false;
    }

    if (notice) {
      notice.hidden = repairHintDismissed;
    }

    setText("repair-queue-notice-text", count === 1
      ? "Ein Mail-Link ist derzeit nicht erreichbar. Outlook vor der Reparatur neu starten."
      : count + " Mail-Links sind derzeit nicht erreichbar. Outlook vor der Reparatur neu starten.");

    return count;
  } catch {
    // Der Hinweis ist Komfortfunktion; Agentfehler werden über SHL angezeigt.
    return -1;
  }
}

async function refreshRepairQueueAfterMailChange(): Promise<void> {
  const card = document.getElementById("repair-queue-card");
  const repairQueueWasOpen = !!card && !card.hidden;
  const count = await refreshRepairQueueNotice(false);

  if (count <= 0) {
    return;
  }

  if (repairQueueWasOpen) {
    await showRepairQueue();
  }
}

function finishRepairQueue(): void {
  repairQueueCount = 0;
  repairHintDismissed = false;
  hideElement("repair-queue-notice");
  hideElement("repair-queue-card");

  // Eine gerade gesetzte Erfolgsmeldung darf durch die anschließende
  // Queue-Bereinigung nicht sofort wieder überschrieben werden.
  // Der Timer in setShlStatus("repaired") stellt den normalen Status
  // nach 3,5 Sekunden selbst wieder her.
  if (shlStatusResetTimer === undefined) {
    setShlStatus("active");
  }

  const list = document.getElementById("repair-queue-list");
  if (list) {
    list.innerHTML = "";
  }
}

async function showRepairQueue(): Promise<void> {
  const response = await fetch(AgentUrl + "/repairqueue");
  if (!response.ok) throw new Error("Repair queue returned HTTP " + response.status);
  const result = await response.json();
  const items = (result.items || []) as RepairQueueItem[];
  const list = document.getElementById("repair-queue-list");
  if (!list) return;

  if (items.length === 0) {
    finishRepairQueue();
    return;
  }

  repairHintDismissed = true;
  hideElement("repair-queue-notice");
  list.innerHTML = "";

  for (const item of items) {
    const row = document.createElement("div");
    row.className = "repair-item";

    const title = document.createElement("div");
    title.className = "repair-item-title";
    title.textContent = item.subject || "(ohne Betreff)";
    row.appendChild(title);

    const meta = document.createElement("div");
    meta.className = "repair-item-meta";
    meta.textContent = [item.senderName, formatDate(item.mailDate)].filter(Boolean).join(" · ");
    row.appendChild(meta);

    const actions = document.createElement("div");
    actions.className = "repair-item-actions";

    const searchFeedback = document.createElement("div");
    searchFeedback.className = "repair-search-feedback";
    searchFeedback.hidden = true;

    const searchButton = document.createElement("button");
    searchButton.type = "button";
    searchButton.textContent = "Mail für Reparatur auswählen";

    const searchText = buildOutlookSearchText(item);
    if (!searchText) {
      searchButton.disabled = true;
      searchButton.title = "Für diese Mail sind weder Betreff, Absender noch Empfangsdatum gespeichert.";
      searchFeedback.textContent =
        "Für diese Mail konnte kein brauchbarer Outlook-Suchtext erzeugt werden. " +
        "Bitte anhand der angezeigten Angaben manuell suchen.";
      searchFeedback.hidden = false;
    } else {
      searchButton.onclick = async () => {
        if (!navigator.clipboard) {
          searchFeedback.textContent =
            "Der Suchtext konnte nicht automatisch kopiert werden: „" + searchText + "“.";
          searchFeedback.hidden = false;
          return;
        }

        await navigator.clipboard.writeText(searchText);

        searchButton.textContent = "Suchtext erneut kopieren";
        searchFeedback.textContent =
          "Für die Outlook-Suche kopiert: „" + searchText + "“. " +
          "Oben in Outlook einfügen und die passende Mail öffnen. " +
          "MailNotes repariert den Link anschließend automatisch.";
        searchFeedback.hidden = false;
      };
    }
    actions.appendChild(searchButton);

    const skipButton = document.createElement("button");
    skipButton.type = "button";
    skipButton.textContent = "Überspringen";
    skipButton.onclick = async () => {
      await setRepairQueueStatus(item.id, 3);
      await showRepairQueue();
      await refreshRepairQueueNotice();
    };
    actions.appendChild(skipButton);

    row.appendChild(actions);
    row.appendChild(searchFeedback);
    list.appendChild(row);
  }

  const card = document.getElementById("repair-queue-card");
  if (card) card.hidden = false;
}


function buildOutlookSearchText(item: RepairQueueItem): string {
  const subject = sanitizeOutlookSearchValue(item.subject || "");
  const sender = sanitizeOutlookSearchValue(
    extractRepairQueueSender(item.senderName || item.senderAddress || "")
  );

  const parts: string[] = [];

  if (subject) {
    parts.push('subject:"' + subject + '"');
  }

  if (sender) {
    parts.push('from:"' + sender + '"');
  }

  // Sind Betreff oder Absender vorhanden, werden ausschließlich diese
  // verwertbaren Parameter verwendet. Das Datum ist nur der letzte Fallback.
  if (parts.length > 0) {
    return parts.join(" ");
  }

  const receivedDate = buildOutlookReceivedDate(
    item.mailDate || extractRepairQueueDate(item.senderName || "")
  );

  return receivedDate
    ? 'received:"' + receivedDate + '"'
    : "";
}

function parseResolvedRepairMetadata(value: any): {
  senderName: string;
  mailDate: string;
} {
  const subtitle = value ? value.toString().trim() : "";
  if (!subtitle) {
    return { senderName: "", mailDate: "" };
  }

  const separatorIndex = subtitle.indexOf(" · ");
  if (separatorIndex < 0) {
    return { senderName: subtitle, mailDate: "" };
  }

  const senderName = subtitle.substring(0, separatorIndex).trim();
  const dateText = subtitle.substring(separatorIndex + 3).trim();
  const parsedDate = new Date(dateText);

  return {
    senderName,
    mailDate: Number.isNaN(parsedDate.getTime()) ? dateText : parsedDate.toISOString()
  };
}

function extractRepairQueueSender(value: string): string {
  const sender = value.trim();
  const separatorIndex = sender.indexOf(" · ");

  return separatorIndex >= 0
    ? sender.substring(0, separatorIndex).trim()
    : sender;
}

function extractRepairQueueDate(value: string): string {
  const separatorIndex = value.indexOf(" · ");
  return separatorIndex >= 0
    ? value.substring(separatorIndex + 3).trim()
    : "";
}

function buildOutlookReceivedDate(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  const date = new Date(trimmed);
  if (Number.isNaN(date.getTime())) {
    return sanitizeOutlookSearchValue(trimmed);
  }

  return String(date.getDate()).padStart(2, "0") + "." +
    String(date.getMonth() + 1).padStart(2, "0") + "." +
    date.getFullYear();
}

function sanitizeOutlookSearchValue(value: string): string {
  return value
    .trim()
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s{2,}/g, " ")
    .replace(/"/g, "");
}

async function setRepairQueueStatus(id: number, status: number): Promise<void> {
  const body = new URLSearchParams();
  body.append("id", id.toString());
  body.append("status", status.toString());
  const response = await fetch(AgentUrl + "/repairqueue/status", { method: "POST", body });
  if (!response.ok) throw new Error("Repair queue returned HTTP " + response.status);
}

function hideElement(id: string): void {
  const element = document.getElementById(id);
  if (element) element.hidden = true;
}
