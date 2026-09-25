/* global Office, document, window, navigator */

declare const __MAILNOTES_BUILD_TIME__: string;
declare const __MAILNOTES_BUILD_COMMIT__: string;

import { MailNotesConfig } from "../config";

const AgentUrl = MailNotesConfig.agentUrl;
const ENABLE_LOGGING = MailNotesConfig.enableLogging;

let itemChangeSequence = 0;
let currentMailNotesId = "";
let shlStatusResetTimer: number | undefined;
let repairQueueCount = 0;
let currentGmlActive = false;
let currentMailKnown = false;
let repairHintDismissed = false;
let repairWatchTimer: number | undefined;
let repairWatchDeadline = 0;
let repairWatchItemIds = new Set<number>();
const REPAIR_WATCH_INTERVAL_MS = 1500;
const REPAIR_WATCH_TIMEOUT_MS = 2 * 60 * 1000;
let searchTimer: number | undefined;
let searchSequence = 0;
let favoritesFilterActive = false;
let activeTags: TagInfo[] = [];
let tagPanelOpen = false;
let activePersons: PersonInfo[] = [];
let personPanelOpen = false;
let currentNoteExists = false;
let currentNoteIsFavorite = false;


const CATEGORY_SETTING_ENABLED = "mailnotes.category.enabled";
const CATEGORY_SETTING_NAME = "mailnotes.category.name";
const CATEGORY_SETTING_COLOR = "mailnotes.category.color";
const DEFAULT_CATEGORY_NAME = "MailNotes";
const DEFAULT_CATEGORY_COLOR = "Preset3"; // Gelb

type MailNotesCategorySettings = {
  enabled: boolean;
  name: string;
  color: string;
};

let mailNotesCategorySettings: MailNotesCategorySettings = {
  enabled: false,
  name: DEFAULT_CATEGORY_NAME,
  color: DEFAULT_CATEGORY_COLOR
};
let categorySyncInProgress = false;
let categorySyncPending = false;
let categoryMasterCacheName = "";
let categoryLastSyncedSequence = 0;
let categoryLastSyncedHasNote: boolean | null = null;

type NoteAutocompleteKind = "tag" | "person";
type NoteAutocompleteItem = {
  name: string;
  normalizedName: string;
  count?: number;
};
type NoteAutocompleteContext = {
  kind: NoteAutocompleteKind;
  start: number;
  end: number;
  query: string;
};

let noteAutocompleteKind: NoteAutocompleteKind | null = null;
let noteAutocompleteItems: NoteAutocompleteItem[] = [];
let noteAutocompleteSelectedIndex = 0;
let noteAutocompleteTags: TagInfo[] | null = null;
let noteAutocompletePersons: PersonInfo[] | null = null;
let noteAutocompleteRequestSequence = 0;

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

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

async function fetchAgentWithStartupRetry(
  path: string,
  attempts: number = 8,
  delayMs: number = 500
): Promise<Response> {
  let lastError: unknown;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      const response = await fetch(AgentUrl + path);
      if (response.ok) {
        return response;
      }

      lastError = new Error("Agent returned HTTP " + response.status);
    } catch (error) {
      lastError = error;
    }

    if (attempt < attempts) {
      await delay(delayMs);
    }
  }

  throw lastError instanceof Error
    ? lastError
    : new Error("MailNotes Agent ist nicht erreichbar.");
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

  const sequence = ++itemChangeSequence;

  clearCurrentMailDisplay();
  showMailInformation();
  await loadNote(sequence);
  await refreshRepairQueueNotice(false);
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

function normalizeOfficeText(value: unknown): string {
  return (value || "").toString().normalize("NFC");
}

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
    // Outlook can return visually identical Unicode in decomposed form
    // (for example "u" + COMBINING DIAERESIS instead of "ü").
    // Normalize Office text before it crosses the HTTP/SQLite boundary.
    subject: normalizeOfficeText(item.subject),
    senderName:
      normalizeOfficeText((item as any).from?.displayName),
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

  if (showRepairStatus && (result.updated || result.repairCompleted)) {
    fileLog("SHL repaired", {
      mailNotesId: result.mailNotesId,
      oldItemId: result.oldItemId,
      itemId: result.itemId,
      messageId: identity.messageId,
      subject: identity.subject,
      updated: !!result.updated,
      repairCompleted: !!result.repairCompleted
    });

    // repairCompleted ist das eindeutige Server-Signal dafür, dass ein
    // offener SHL-Reparaturauftrag durch diese Mail abgeschlossen wurde.
    // Die Queue sofort nachziehen; der Repair-Watch bleibt als Absicherung
    // für Reparaturen in einer anderen Outlook-Fensterinstanz aktiv.
    if (result.repairCompleted) {
      await refreshRepairQueueNotice(false);
    }

    setShlStatus("repaired");
  }

  return result;
}

function currentMailLinkMode(): "GML" | "SRL" {
  return currentGmlActive ? "GML" : "SRL";
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
    element.textContent = "🟢 " + currentMailLinkMode() + " aktiv · Mail aktualisiert";
    element.title =
      "Die technische Outlook-ID der verschobenen Mail wurde aktualisiert. Links und Backlinks sind wieder gültig.";

    shlStatusResetTimer = window.setTimeout(() => {
      setShlStatus(repairQueueCount > 0 ? "pending" : "active");
    }, 3500);

    return;
  }

  if (state === "pending") {
    element.textContent = repairQueueCount === 1
      ? "🟡 " + currentMailLinkMode() + " aktiv · 1 Reparatur offen"
      : "🟡 " + currentMailLinkMode() + " aktiv · " + repairQueueCount + " Reparaturen offen";
    element.title = repairQueueCount === 1
      ? "Eine Mail-Zuordnung muss repariert werden. Klicken Sie auf den gelben Hinweis, sobald Sie die Reparatur fortsetzen möchten."
      : repairQueueCount + " Mail-Zuordnungen müssen repariert werden.";
    return;
  }

  if (state === "error") {
    element.textContent = "🟠 " + currentMailLinkMode() + " gestört";
    element.title =
      "Die automatische Aktualisierung konnte den MailNotesAgent nicht erreichen.";
    return;
  }

  element.textContent = "🟢 " + currentMailLinkMode() + " aktiv";
  element.title = currentGmlActive
    ? "Für diese Mail ist GML verfügbar. Die SRL bleibt parallel erhalten und kann aus der GML aktualisiert werden."
    : "Für diese Mail ist SRL aktiv. GML ist für diese Mail oder dieses Konto derzeit nicht verfügbar.";
}

function clearCurrentMailDisplay() {
  autosaveSuppressed = true;
  currentNoteExists = false;
  currentNoteIsFavorite = false;
  updateFavoriteButton();
  cancelPendingAutosave();
  currentMailNotesId = "";
  currentGmlActive = false;
  currentMailKnown = false;

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
    btnOpenRepairQueue.onclick = async () => {
      await showRepairQueue();
      await startRepairWatch();
    };
  }

  if (btnDismissRepairQueue) {
    btnDismissRepairQueue.onclick = () => {
      stopRepairWatch();
      repairHintDismissed = true;
      hideElement("repair-queue-notice");
      setShlStatus(repairQueueCount > 0 ? "pending" : "active");
    };
  }

  if (btnCloseRepairQueue) {
    btnCloseRepairQueue.onclick = () => {
      stopRepairWatch();
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
  const noteTags = document.getElementById("note-tags");
  if (noteTags) {
    noteTags.innerHTML = "";
    noteTags.hidden = true;
  }
  const notePersons = document.getElementById("note-persons");
  if (notePersons) {
    notePersons.innerHTML = "";
    notePersons.hidden = true;
  }
}


function categoryApiSupported(): boolean {
  try {
    return Boolean(
      Office.context.requirements?.isSetSupported("Mailbox", "1.8") &&
      Office.context.mailbox.masterCategories &&
      (Office.context.mailbox.item as any)?.categories
    );
  } catch {
    return false;
  }
}

function readCategorySettings(): MailNotesCategorySettings {
  const roaming = Office.context.roamingSettings;
  const enabled = roaming.get(CATEGORY_SETTING_ENABLED);
  const name = roaming.get(CATEGORY_SETTING_NAME);
  const color = roaming.get(CATEGORY_SETTING_COLOR);

  return {
    enabled: enabled === true,
    name: typeof name === "string" && name.trim() ? name.trim() : DEFAULT_CATEGORY_NAME,
    color: typeof color === "string" && color ? color : DEFAULT_CATEGORY_COLOR
  };
}

function saveRoamingSettings(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error?.message || "Einstellungen konnten nicht gespeichert werden."));
      }
    });
  });
}

function setCategoryStatus(text: string): void {
  setText("category-status", text);
}

function setupCategorySettings(): void {
  mailNotesCategorySettings = readCategorySettings();

  const enabled = document.getElementById("category-enabled") as HTMLInputElement | null;
  const name = document.getElementById("category-name") as HTMLInputElement | null;
  const color = document.getElementById("category-color") as HTMLSelectElement | null;

  if (!enabled || !name || !color) {
    return;
  }

  enabled.checked = mailNotesCategorySettings.enabled;
  name.value = mailNotesCategorySettings.name;
  color.value = mailNotesCategorySettings.color;

  const supported = categoryApiSupported();
  enabled.disabled = !supported;
  name.disabled = !supported || !enabled.checked;
  color.disabled = !supported || !enabled.checked;

  if (!supported) {
    setCategoryStatus("Outlook-Kategorien werden von dieser Umgebung nicht unterstützt.");
    return;
  }

  const updateDependentControls = () => {
    name.disabled = !enabled.checked;
    color.disabled = !enabled.checked;
  };

  const persist = async (syncCurrentItem: boolean) => {
    const oldName = mailNotesCategorySettings.name;
    const newName = name.value.trim() || DEFAULT_CATEGORY_NAME;

    name.value = newName;
    mailNotesCategorySettings = {
      enabled: enabled.checked,
      name: newName,
      color: color.value || DEFAULT_CATEGORY_COLOR
    };

    // Einstellungsänderungen erzwingen genau eine neue Synchronisation.
    categoryLastSyncedHasNote = null;
    if (oldName !== newName) {
      categoryMasterCacheName = "";
    }

    Office.context.roamingSettings.set(CATEGORY_SETTING_ENABLED, mailNotesCategorySettings.enabled);
    Office.context.roamingSettings.set(CATEGORY_SETTING_NAME, mailNotesCategorySettings.name);
    Office.context.roamingSettings.set(CATEGORY_SETTING_COLOR, mailNotesCategorySettings.color);

    setCategoryStatus("Speichert …");

    try {
      await saveRoamingSettings();
      setCategoryStatus("");

      if (mailNotesCategorySettings.enabled) {
        // Kategorie beim Aktivieren/Ändern der Einstellung sofort sicherstellen.
        // Dadurch erhalten wir auch unmittelbar eine aussagekräftige API-Fehlermeldung,
        // falls Outlook das Anlegen der Master-Kategorie verweigert.
        await ensureMasterCategoryExists(mailNotesCategorySettings);
      }

      if (syncCurrentItem) {
        await synchronizeMailNotesCategoryForCurrentItem(oldName !== newName ? oldName : "");
      }
    } catch (error) {
      setCategoryStatus(error instanceof Error ? error.message : String(error));
    }
  };

  enabled.onchange = () => {
    updateDependentControls();
    void persist(true);
  };

  name.onchange = () => {
    void persist(true);
  };

  name.onblur = () => {
    if (name.value.trim() !== mailNotesCategorySettings.name) {
      void persist(true);
    }
  };

  color.onchange = () => {
    void persist(true);
  };
}

function noteHasMailNotesContent(content: string, links: string): boolean {
  return Boolean(content.trim() || links.trim());
}

function getCurrentEditorHasMailNotesContent(): boolean {
  return noteHasMailNotesContent(
    getEditorText("note-content"),
    getEditorText("note-links")
  );
}

function officeAsyncToPromise<T>(
  invoke: (callback: (result: Office.AsyncResult<T>) => void) => void
): Promise<T> {
  return new Promise((resolve, reject) => {
    invoke((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message || "Outlook-Kategorie konnte nicht aktualisiert werden."));
      }
    });
  });
}

async function ensureMasterCategoryExists(settings: MailNotesCategorySettings): Promise<void> {
  // Die Master-Kategorienliste ist eine relativ teure Outlook-API-Abfrage.
  // Innerhalb einer MailNotes-Sitzung müssen wir sie für denselben Namen nur
  // einmal prüfen. Das verhindert Office.js-Aufrufe bei jedem Autosave.
  if (categoryMasterCacheName === settings.name) {
    return;
  }

  const masterCategories = Office.context.mailbox.masterCategories;
  const categories = await officeAsyncToPromise<Office.CategoryDetails[]>((callback) =>
    masterCategories.getAsync(callback)
  );

  const exists = (categories || []).some(
    (category) => category.displayName.localeCompare(settings.name, undefined, { sensitivity: "accent" }) === 0
  );

  if (exists) {
    categoryMasterCacheName = settings.name;
    return;
  }

  const colorEnum = (Office.MailboxEnums.CategoryColor as any)[settings.color] ||
    Office.MailboxEnums.CategoryColor.Preset3;

  try {
    await officeAsyncToPromise<void>((callback) =>
      masterCategories.addAsync([
        {
          displayName: settings.name,
          color: colorEnum
        }
      ], callback)
    );
  } catch (error) {
    // Parallel geöffnete MailNotes-Fenster können dieselbe Kategorie gleichzeitig
    // anlegen. In diesem Fall reicht es, die Masterliste erneut zu prüfen.
    const after = await officeAsyncToPromise<Office.CategoryDetails[]>((callback) =>
      masterCategories.getAsync(callback)
    );
    const nowExists = (after || []).some(
      (category) => category.displayName.localeCompare(settings.name, undefined, { sensitivity: "accent" }) === 0
    );
    if (!nowExists) {
      throw error;
    }
  }

  categoryMasterCacheName = settings.name;
}

async function synchronizeMailNotesCategoryForCurrentItem(previousCategoryName: string = ""): Promise<void> {
  if (!categoryApiSupported()) {
    return;
  }

  if (categorySyncInProgress) {
    categorySyncPending = true;
    return;
  }

  categorySyncInProgress = true;

  try {
    do {
      categorySyncPending = false;

      const item = Office.context.mailbox.item as any;
      const categories = item?.categories;
      if (!categories) {
        return;
      }

      const settings = mailNotesCategorySettings;
      const hasNote = getCurrentEditorHasMailNotesContent();

      // Für die laufende Mail interessiert nur der Zustandswechsel
      // "keine Notiz" <-> "Notiz vorhanden". Normales Weiterschreiben in
      // einer bereits vorhandenen Notiz darf keinerlei Outlook-Category-API
      // mehr auslösen.
      if (!previousCategoryName &&
          categoryLastSyncedSequence === itemChangeSequence &&
          categoryLastSyncedHasNote === hasNote) {
        return;
      }

      // Ist die Option abgeschaltet, lassen wir vorhandene Kategorien bewusst
      // unangetastet. MailNotes entfernt dann weder alte Zuordnungen noch die
      // Master-Kategorie aus Outlook.
      if (!settings.enabled) {
        return;
      }

      if (hasNote) {
        await ensureMasterCategoryExists(settings);
      }

      const itemCategories = await officeAsyncToPromise<Office.CategoryDetails[]>((callback) =>
        categories.getAsync(callback)
      );
      const assigned = (itemCategories || []).map((category) => category.displayName);

      if (previousCategoryName && previousCategoryName !== settings.name &&
          assigned.some((value) => value === previousCategoryName)) {
        await officeAsyncToPromise<void>((callback) =>
          categories.removeAsync([previousCategoryName], callback)
        );
      }

      const isAssigned = assigned.some((value) => value === settings.name);
      const shouldBeAssigned = hasNote;

      if (shouldBeAssigned && !isAssigned) {
        await officeAsyncToPromise<void>((callback) =>
          categories.addAsync([settings.name], callback)
        );
      } else if (!shouldBeAssigned && isAssigned) {
        await officeAsyncToPromise<void>((callback) =>
          categories.removeAsync([settings.name], callback)
        );
      }

      categoryLastSyncedSequence = itemChangeSequence;
      categoryLastSyncedHasNote = shouldBeAssigned;
      previousCategoryName = "";
    } while (categorySyncPending);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    log("Kategorie konnte nicht synchronisiert werden:", message);
    fileLog("category sync error", { message });
    setCategoryStatus("Kategorie: " + message);
  } finally {
    categorySyncInProgress = false;
  }
}

function setupButtons() {
  setupInfoPanel();
  setupDatabasePathButton();
  setupCategorySettings();

  const shlStatus = document.getElementById("shl-status");
  if (shlStatus) {
    shlStatus.onclick = async () => {
      if (repairQueueCount <= 0) return;
      await showRepairQueue();
      await startRepairWatch();
    };
  }

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

  const btnFavorites =
    document.getElementById("btn-favorites");

  const btnTags =
    document.getElementById("btn-tags");

  const btnPersons =
    document.getElementById("btn-persons");

  const btnNoteFavorite =
    document.getElementById("btn-note-favorite");

  const btnNoteDelete =
    document.getElementById("btn-note-delete");

  const searchCard =
    document.querySelector(".search-card");

  if (noteContent) {
    noteContent.oninput = () => {
      scheduleAutosave();
      void updateNoteAutocomplete(noteContent);
    };

    // Outlook für Windows verwendet TAB selbst zur Fokusnavigation.
    // Deshalb behandeln wir TAB zusätzlich in der Capture-Phase, bevor der
    // normale Keydown-Handler und die Host-Navigation zum Zug kommen.
    noteContent.addEventListener(
      "keydown",
      (event) => {
        if (event.key === "Tab" &&
            handleNoteAutocompleteKeyDown(noteContent, event)) {
          event.stopImmediatePropagation();
        }
      },
      true
    );

    noteContent.onkeydown = (event) => {
      handleNoteAutocompleteKeyDown(noteContent, event);
    };

    noteContent.onclick = () => {
      void updateNoteAutocomplete(noteContent);
    };

    noteContent.onkeyup = (event) => {
      if (event.key === "ArrowLeft" || event.key === "ArrowRight" ||
          event.key === "Home" || event.key === "End") {
        void updateNoteAutocomplete(noteContent);
      }
    };

    noteContent.onblur = () => {
      window.setTimeout(closeNoteAutocomplete, 120);
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

  if (btnFavorites) {
    btnFavorites.onclick = () => {
      void toggleFavoritesFilter();
    };
  }

  if (btnTags) {
    btnTags.onclick = () => {
      void toggleTagPanel();
    };
  }

  if (btnPersons) {
    btnPersons.onclick = () => {
      void togglePersonPanel();
    };
  }

  if (btnNoteFavorite) {
    btnNoteFavorite.onclick = () => {
      void toggleCurrentNoteFavorite();
    };
  }

  if (btnNoteDelete) {
    btnNoteDelete.onclick = () => {
      void deleteCurrentNote();
    };
  }

  document.addEventListener("click", (event) => {
    if (!searchCard) return;

    const target = event.target as Node | null;
    if (target && !searchCard.contains(target)) {
      const hadFilter = favoritesFilterActive || activeTags.length > 0 || activePersons.length > 0;
      setFavoritesFilterActive(false);
      setActiveTags([]);
      setActivePersons([]);
      closeTagPanel();
      closePersonPanel();
      if (hadFilter) void refreshSearchForCurrentState();
    }
  });

  void refreshStatisticsCounts();
}

function updateFavoriteButton(): void {
  const button = document.getElementById("btn-note-favorite") as HTMLButtonElement | null;
  if (!button) return;

  button.disabled = !currentNoteExists;
  button.textContent = currentNoteIsFavorite ? "♥" : "♡";

  const deleteButton = document.getElementById("btn-note-delete") as HTMLButtonElement | null;
  if (deleteButton) deleteButton.disabled = !currentNoteExists;
  button.classList.toggle("active", currentNoteIsFavorite);
  button.setAttribute("aria-pressed", currentNoteIsFavorite ? "true" : "false");
  button.title = currentNoteIsFavorite ? "Favorit entfernen" : "Als Favorit markieren";
  button.setAttribute("aria-label", button.title);
}

async function toggleCurrentNoteFavorite(): Promise<void> {
  if (!currentNoteExists) return;
  const identity = getCurrentMailSnapshot();
  if (!identity) return;

  const nextValue = !currentNoteIsFavorite;
  const body = new URLSearchParams();
  if (currentMailNotesId) body.append("mailNotesId", currentMailNotesId);
  body.append("messageId", identity.messageId);
  body.append("favorite", nextValue ? "true" : "false");

  const button = document.getElementById("btn-note-favorite") as HTMLButtonElement | null;
  if (button) button.disabled = true;

  try {
    const response = await fetch(AgentUrl + "/favorite", { method: "POST", body });
    if (!response.ok) throw new Error("Agent returned HTTP " + response.status);

    currentNoteIsFavorite = nextValue;
    updateFavoriteButton();
    await refreshStatisticsCounts();
    if (favoritesFilterActive) await refreshSearchForCurrentState();
  } catch (error) {
    console.error("Favorit konnte nicht geändert werden:", error);
    updateFavoriteButton();
  }
}

function confirmNoteDeletion(
  titleText = "Notiz wirklich löschen?",
  bodyText = "Inhalt, Links, Favorit sowie Tag-/Personen-Zuordnungen werden entfernt."
): Promise<boolean> {
  const overlay = document.getElementById("note-delete-confirm");
  const cancelButton = document.getElementById("btn-note-delete-cancel") as HTMLButtonElement | null;
  const confirmButton = document.getElementById("btn-note-delete-confirm") as HTMLButtonElement | null;
  const title = document.getElementById("note-delete-confirm-title");
  const text = document.getElementById("note-delete-confirm-text");

  if (!overlay || !cancelButton || !confirmButton || !title || !text) {
    console.error("Löschdialog ist im Taskpane nicht vorhanden.");
    return Promise.resolve(false);
  }

  title.textContent = titleText;
  text.textContent = bodyText;

  return new Promise<boolean>((resolve) => {
    let finished = false;

    const finish = (result: boolean) => {
      if (finished) return;
      finished = true;
      overlay.hidden = true;
      cancelButton.onclick = null;
      confirmButton.onclick = null;
      overlay.onclick = null;
      document.removeEventListener("keydown", onKeyDown, true);
      resolve(result);
    };

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        finish(false);
      }
    };

    cancelButton.onclick = () => finish(false);
    confirmButton.onclick = () => finish(true);
    overlay.onclick = (event) => {
      if (event.target === overlay) finish(false);
    };
    document.addEventListener("keydown", onKeyDown, true);

    overlay.hidden = false;
    window.setTimeout(() => confirmButton.focus(), 0);
  });
}

async function deleteCurrentNote(): Promise<void> {
  if (!currentNoteExists) return;

  const identity = getCurrentMailSnapshot();
  if (!identity) return;

  const confirmed = await confirmNoteDeletion();
  if (!confirmed) return;

  // Noch nicht gestartete Autosaves verwerfen und einen bereits laufenden
  // Speichervorgang zuerst sauber beenden. So kann ein verspäteter POST die
  // eben gelöschte Notiz nicht unmittelbar wiederherstellen.
  cancelPendingAutosave();
  if (autosaveActivePromise) {
    await autosaveActivePromise;
  }

  const params = new URLSearchParams();
  if (currentMailNotesId) params.append("mailNotesId", currentMailNotesId);
  params.append("messageId", identity.messageId);

  const deleteButton = document.getElementById("btn-note-delete") as HTMLButtonElement | null;
  const favoriteButton = document.getElementById("btn-note-favorite") as HTMLButtonElement | null;
  if (deleteButton) deleteButton.disabled = true;
  if (favoriteButton) favoriteButton.disabled = true;

  try {
    const response = await fetch(AgentUrl + "/note?" + params.toString(), {
      method: "DELETE"
    });

    if (!response.ok) {
      throw new Error("Agent returned HTTP " + response.status);
    }

    currentNoteExists = false;
    currentNoteIsFavorite = false;
    invalidateNoteAutocompleteCache();

    setEditorText("note-content", "");
    setEditorText("note-links", "");
    setNoteMeta("", "");
    setAutosaveStatus("");
    updateFavoriteButton();

    await renderLinks(itemChangeSequence);
    await renderBacklinks(currentMailNotesId || identity.messageId, itemChangeSequence);
    await refreshCurrentNoteTags();
    await refreshCurrentNotePersons();
    await refreshStatisticsCounts();
    await refreshSearchForCurrentState();
    void synchronizeMailNotesCategoryForCurrentItem();
  } catch (error) {
    console.error("Notiz konnte nicht gelöscht werden:", error);
    window.alert("Die Notiz konnte nicht gelöscht werden.");
    updateFavoriteButton();
  }
}

async function refreshStatisticsCounts(): Promise<void> {
  try {
    const response = await fetchAgentWithStartupRetry("/stats");
    const stats = await response.json() as MailNotesStats;
    setInfoText("favorites-count", stats.favorites);
    setInfoText("tags-count", stats.tags);
    setInfoText("persons-count", stats.persons);
  } catch {
    // Komfortinformation: Fehler beeinflussen das Add-in nicht.
  }
}

function closeNoteAutocomplete(): void {
  noteAutocompleteKind = null;
  noteAutocompleteItems = [];
  noteAutocompleteSelectedIndex = 0;
  noteAutocompleteRequestSequence++;

  const panel = document.getElementById("note-autocomplete");
  if (panel) {
    panel.hidden = true;
    panel.innerHTML = "";
  }
}

function invalidateNoteAutocompleteCache(): void {
  noteAutocompleteTags = null;
  noteAutocompletePersons = null;
}

function getNoteAutocompleteContext(
  editor: HTMLTextAreaElement
): NoteAutocompleteContext | null {
  const caret = editor.selectionStart ?? editor.value.length;
  const textBeforeCaret = editor.value.substring(0, caret);
  const lineStart = Math.max(
    textBeforeCaret.lastIndexOf("\n"),
    textBeforeCaret.lastIndexOf("\r")
  ) + 1;
  const line = textBeforeCaret.substring(lineStart);

  const hashPos = line.lastIndexOf("#");
  const atPos = line.lastIndexOf("@");
  const markerPos = Math.max(hashPos, atPos);

  if (markerPos < 0) return null;

  const marker = line.charAt(markerPos);
  const absoluteMarkerPos = lineStart + markerPos;

  if (absoluteMarkerPos > 0) {
    const previous = editor.value.charAt(absoluteMarkerPos - 1);
    if (previous !== " " && previous !== "\t" && previous !== "\n" && previous !== "\r") {
      return null;
    }
  }

  const query = line.substring(markerPos + 1);

  if (marker === "#") {
    // Tags enthalten keine Leerzeichen. Sobald eines getippt wurde,
    // ist der Tag abgeschlossen und Auto-Complete verschwindet.
    if (/\s/.test(query) || query.indexOf("@") >= 0 || query.indexOf("#") >= 0) {
      return null;
    }

    return {
      kind: "tag",
      start: absoluteMarkerPos,
      end: caret,
      query
    };
  }

  // Personen dürfen Leerzeichen, Bindestriche, Apostrophe und Punkte
  // enthalten. TAB beendet den Namen explizit und darf daher nicht
  // innerhalb der aktuellen Query vorkommen.
  if (query.indexOf("\t") >= 0 || query.indexOf("#") >= 0 || query.indexOf("@") >= 0) {
    return null;
  }

  return {
    kind: "person",
    start: absoluteMarkerPos,
    end: caret,
    query
  };
}

async function loadNoteAutocompleteItems(
  kind: NoteAutocompleteKind
): Promise<NoteAutocompleteItem[]> {
  if (kind === "tag" && noteAutocompleteTags) {
    return noteAutocompleteTags;
  }

  if (kind === "person" && noteAutocompletePersons) {
    return noteAutocompletePersons;
  }

  const response = await fetchAgentWithStartupRetry(kind === "tag" ? "/tags" : "/persons");
  const result = await response.json();
  const items: NoteAutocompleteItem[] = Array.isArray(result.items) ? result.items : [];

  if (kind === "tag") {
    noteAutocompleteTags = items as TagInfo[];
  } else {
    noteAutocompletePersons = items as PersonInfo[];
  }

  return items;
}

function rankNoteAutocompleteItems(
  items: NoteAutocompleteItem[],
  query: string
): NoteAutocompleteItem[] {
  const normalizedQuery = query.trim().toLocaleLowerCase();

  return items
    .filter((item) => {
      if (!normalizedQuery) return true;
      return item.name.toLocaleLowerCase().includes(normalizedQuery);
    })
    .sort((a, b) => {
      const aName = a.name.toLocaleLowerCase();
      const bName = b.name.toLocaleLowerCase();
      const aStarts = normalizedQuery && aName.startsWith(normalizedQuery) ? 0 : 1;
      const bStarts = normalizedQuery && bName.startsWith(normalizedQuery) ? 0 : 1;
      if (aStarts !== bStarts) return aStarts - bStarts;

      const countDiff = (b.count ?? 0) - (a.count ?? 0);
      if (countDiff !== 0) return countDiff;

      return a.name.localeCompare(b.name, "de", { sensitivity: "base" });
    })
    .slice(0, 8);
}

async function updateNoteAutocomplete(editor: HTMLTextAreaElement): Promise<void> {
  const context = getNoteAutocompleteContext(editor);
  if (!context) {
    closeNoteAutocomplete();
    return;
  }

  const requestSequence = ++noteAutocompleteRequestSequence;

  try {
    const allItems = await loadNoteAutocompleteItems(context.kind);
    if (requestSequence !== noteAutocompleteRequestSequence) return;

    const currentContext = getNoteAutocompleteContext(editor);
    if (!currentContext || currentContext.kind !== context.kind ||
        currentContext.start !== context.start || currentContext.query !== context.query) {
      return;
    }

    const items = rankNoteAutocompleteItems(allItems, context.query);
    if (items.length === 0) {
      closeNoteAutocomplete();
      return;
    }

    noteAutocompleteKind = context.kind;
    noteAutocompleteItems = items;
    noteAutocompleteSelectedIndex = 0;
    renderNoteAutocomplete(editor);
  } catch (error) {
    console.error("Auto-Complete konnte nicht geladen werden:", error);
    closeNoteAutocomplete();
  }
}

function renderNoteAutocomplete(editor: HTMLTextAreaElement): void {
  const panel = document.getElementById("note-autocomplete");
  if (!panel || !noteAutocompleteKind || noteAutocompleteItems.length === 0) {
    closeNoteAutocomplete();
    return;
  }

  panel.innerHTML = "";
  panel.hidden = false;

  noteAutocompleteItems.forEach((item, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "note-autocomplete-item" +
      (index === noteAutocompleteSelectedIndex ? " active" : "");

    const label = document.createElement("span");
    label.textContent = (noteAutocompleteKind === "tag" ? "#" : "@") + item.name;
    button.appendChild(label);

    if ((item.count ?? 0) > 0) {
      const count = document.createElement("span");
      count.className = "note-autocomplete-count";
      count.textContent = String(item.count);
      button.appendChild(count);
    }

    button.onmousedown = (event) => {
      // Fokus im Texteditor behalten, damit selectionStart/selectionEnd
      // für das Ersetzen des Tokens erhalten bleiben.
      event.preventDefault();
    };

    button.onclick = (event) => {
      event.preventDefault();
      event.stopPropagation();
      applyNoteAutocompleteSelection(editor, index);
    };

    panel.appendChild(button);
  });
}

function applyNoteAutocompleteSelection(
  editor: HTMLTextAreaElement,
  index: number
): void {
  const context = getNoteAutocompleteContext(editor);
  const item = noteAutocompleteItems[index];

  if (!context || !item || context.kind !== noteAutocompleteKind) {
    closeNoteAutocomplete();
    return;
  }

  const marker = context.kind === "tag" ? "#" : "@";
  const terminator = context.kind === "person" ? "\t" : " ";
  const replacement = marker + item.name + terminator;

  editor.value =
    editor.value.substring(0, context.start) +
    replacement +
    editor.value.substring(context.end);

  const newCaret = context.start + replacement.length;
  editor.focus();
  editor.setSelectionRange(newCaret, newCaret);

  closeNoteAutocomplete();
  scheduleAutosave();
}

function insertPersonTerminator(editor: HTMLTextAreaElement): boolean {
  const context = getNoteAutocompleteContext(editor);
  if (!context || context.kind !== "person" || context.query.trim().length < 2) {
    return false;
  }

  editor.value =
    editor.value.substring(0, context.end) +
    "\t" +
    editor.value.substring(context.end);

  const newCaret = context.end + 1;
  editor.setSelectionRange(newCaret, newCaret);
  closeNoteAutocomplete();
  scheduleAutosave();
  return true;
}

function insertTagTerminator(editor: HTMLTextAreaElement): boolean {
  const context = getNoteAutocompleteContext(editor);
  if (!context || context.kind !== "tag" || context.query.trim().length < 1) {
    return false;
  }

  editor.value =
    editor.value.substring(0, context.end) +
    " " +
    editor.value.substring(context.end);

  const newCaret = context.end + 1;
  editor.setSelectionRange(newCaret, newCaret);
  closeNoteAutocomplete();
  scheduleAutosave();
  return true;
}

function handleNoteAutocompleteKeyDown(
  editor: HTMLTextAreaElement,
  event: KeyboardEvent
): boolean {
  const panel = document.getElementById("note-autocomplete");
  const autocompleteVisible = Boolean(panel && !panel.hidden && noteAutocompleteItems.length > 0);

  if (autocompleteVisible) {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      noteAutocompleteSelectedIndex =
        (noteAutocompleteSelectedIndex + 1) % noteAutocompleteItems.length;
      renderNoteAutocomplete(editor);
      return true;
    }

    if (event.key === "ArrowUp") {
      event.preventDefault();
      noteAutocompleteSelectedIndex =
        (noteAutocompleteSelectedIndex - 1 + noteAutocompleteItems.length) %
        noteAutocompleteItems.length;
      renderNoteAutocomplete(editor);
      return true;
    }

    if (event.key === "Enter" || event.key === "Tab") {
      event.preventDefault();
      applyNoteAutocompleteSelection(editor, noteAutocompleteSelectedIndex);
      return true;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      closeNoteAutocomplete();
      return true;
    }
  }

  // Auch ohne ausgewählten Vorschlag schließt TAB das aktuelle Token ab:
  // @Person mit echtem TAB als eindeutigem Trenner,
  // #Tag mit einem normalen Leerzeichen.
  if (event.key === "Tab") {
    if (insertPersonTerminator(editor) || insertTagTerminator(editor)) {
      event.preventDefault();
      return true;
    }
  }

  return false;
}

type TagInfo = {
  name: string;
  normalizedName: string;
  count?: number;
};

function isTagActive(normalizedName: string): boolean {
  return activeTags.some((tag) => tag.normalizedName === normalizedName);
}

function setActiveTags(tags: TagInfo[]): void {
  activeTags = tags.filter(
    (tag, index, items) =>
      Boolean(tag.normalizedName) &&
      items.findIndex((item) => item.normalizedName === tag.normalizedName) === index
  );

  const button = document.getElementById("btn-tags");
  const label = document.getElementById("tags-filter-label");
  const count = document.getElementById("tags-count");
  const active = activeTags.length > 0;

  button?.classList.toggle("active", active);
  button?.setAttribute("aria-pressed", active ? "true" : "false");

  if (activeTags.length === 1) {
    const tag = activeTags[0];
    button?.setAttribute("title", "Tagfilter: #" + tag.name);
    if (label) label.textContent = "#" + tag.name;
  } else if (activeTags.length > 1) {
    button?.setAttribute(
      "title",
      "Tagfilter: " + activeTags.map((tag) => "#" + tag.name).join(" + ")
    );
    if (label) label.textContent = "#" + activeTags.length;
  } else {
    button?.setAttribute("title", "Tags auswählen");
    if (label) label.textContent = "#";
  }

  if (count) count.hidden = active;
}

function toggleTagSelection(tag: TagInfo): void {
  if (isTagActive(tag.normalizedName)) {
    setActiveTags(activeTags.filter((item) => item.normalizedName !== tag.normalizedName));
  } else {
    setActiveTags([...activeTags, tag]);
  }
}

function closeTagPanel(): void {
  tagPanelOpen = false;
  closePersonPanel();
  const panel = document.getElementById("tag-filter-panel");
  const button = document.getElementById("btn-tags");
  if (panel) panel.hidden = true;
  button?.setAttribute("aria-expanded", "false");
}

async function toggleTagPanel(): Promise<void> {
  if (tagPanelOpen) {
    closeTagPanel();
    return;
  }

  closePersonPanel();
  const panel = document.getElementById("tag-filter-panel");
  const button = document.getElementById("btn-tags");
  if (!panel) return;

  tagPanelOpen = true;
  panel.hidden = false;
  panel.textContent = "Tags werden geladen …";
  button?.setAttribute("aria-expanded", "true");

  try {
    const response = await fetch(AgentUrl + "/tags");
    if (!response.ok) throw new Error("Agent returned HTTP " + response.status);
    const result = await response.json();
    const tags: TagInfo[] = Array.isArray(result.items) ? result.items : [];
    setInfoText("tags-count", result.count ?? tags.length);
    renderTagPanel(tags);
  } catch (error) {
    console.error("Tags konnten nicht geladen werden:", error);
    panel.textContent = "Tags nicht verfügbar.";
  }
}

function renderTagPanel(tags: TagInfo[]): void {
  const panel = document.getElementById("tag-filter-panel");
  if (!panel) return;
  panel.innerHTML = "";

  const allButton = document.createElement("button");
  allButton.type = "button";
  allButton.className = "tag-filter-item" + (activeTags.length === 0 ? " active" : "");
  const allLabel = document.createElement("span");
  allLabel.textContent = "Alle Tags";
  allButton.appendChild(allLabel);
  allButton.onclick = (event) => {
    event.stopPropagation();
    setActiveTags([]);
    renderTagPanel(tags);
    void refreshSearchForCurrentState();
  };
  panel.appendChild(allButton);

  for (const tag of tags) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tag-filter-item" +
      (isTagActive(tag.normalizedName) ? " active" : "");

    const name = document.createElement("span");
    name.textContent = "#" + tag.name;
    button.appendChild(name);

    const count = document.createElement("span");
    count.className = "tag-filter-item-count";
    count.textContent = String(tag.count ?? 0);
    button.appendChild(count);

    button.onclick = (event) => {
      event.stopPropagation();
      toggleTagSelection(tag);
      renderTagPanel(tags);
      void refreshSearchForCurrentState();
    };

    panel.appendChild(button);
  }
}

async function refreshCurrentNoteTags(): Promise<void> {
  const container = document.getElementById("note-tags");
  if (!container) return;

  if (!currentNoteExists || !currentMailNotesId) {
    if (container.childElementCount > 0 || !container.hidden) {
      container.innerHTML = "";
      container.hidden = true;
    }
    return;
  }

  try {
    const response = await fetch(
      AgentUrl + "/note/tags?mailNotesId=" + encodeURIComponent(currentMailNotesId)
    );
    if (!response.ok) return;

    const result = await response.json();
    const tags: TagInfo[] = Array.isArray(result.items) ? result.items : [];

    // Die Anzeige nur neu aufbauen, wenn sich die Tag-Liste wirklich geändert hat.
    // So bleibt die Höhe unter dem Editor beim Autosave stabil und es ruckelt nicht.
    const newSignature = tags.map((tag) => tag.normalizedName || tag.name).join("\u0001");
    const currentSignature = Array.from(
      container.querySelectorAll<HTMLButtonElement>(".note-tag")
    )
      .map((button) => (button.dataset.normalizedName || button.textContent || "").replace(/^#/, ""))
      .join("\u0001");

    if (newSignature === currentSignature) {
      return;
    }

    if (tags.length === 0) {
      container.innerHTML = "";
      container.hidden = true;
      return;
    }

    const fragment = document.createDocumentFragment();
    for (const tag of tags) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "note-tag";
      button.dataset.normalizedName = tag.normalizedName || tag.name;
      button.textContent = "#" + tag.name;
      button.title = "Nach #" + tag.name + " filtern";
      button.onclick = (event) => {
        event.stopPropagation();
        toggleTagSelection(tag);
        void refreshSearchForCurrentState();
        document.querySelector(".search-card")?.scrollIntoView({ block: "nearest" });
      };
      fragment.appendChild(button);
    }

    container.replaceChildren(fragment);
    container.hidden = false;
  } catch {
    // Tag-Anzeige ist Komfortfunktion.
  }
}

type PersonInfo = {
  name: string;
  normalizedName: string;
  count?: number;
};

function isPersonActive(normalizedName: string): boolean {
  return activePersons.some((person) => person.normalizedName === normalizedName);
}

function setActivePersons(persons: PersonInfo[]): void {
  activePersons = persons.filter(
    (person, index, items) =>
      Boolean(person.normalizedName) &&
      items.findIndex((item) => item.normalizedName === person.normalizedName) === index
  );

  const button = document.getElementById("btn-persons");
  const label = document.getElementById("persons-filter-label");
  const count = document.getElementById("persons-count");
  const active = activePersons.length > 0;

  button?.classList.toggle("active", active);
  button?.setAttribute("aria-pressed", active ? "true" : "false");

  if (activePersons.length === 1) {
    const person = activePersons[0];
    button?.setAttribute("title", "Personenfilter: @" + person.name);
    if (label) label.textContent = "@" + person.name;
  } else if (activePersons.length > 1) {
    button?.setAttribute(
      "title",
      "Personenfilter: " + activePersons.map((person) => "@" + person.name).join(" + ")
    );
    if (label) label.textContent = "@" + activePersons.length;
  } else {
    button?.setAttribute("title", "Personen auswählen");
    if (label) label.textContent = "@";
  }

  if (count) count.hidden = active;
}

function togglePersonSelection(person: PersonInfo): void {
  if (isPersonActive(person.normalizedName)) {
    setActivePersons(activePersons.filter((item) => item.normalizedName !== person.normalizedName));
  } else {
    setActivePersons([...activePersons, person]);
  }
}

function closePersonPanel(): void {
  personPanelOpen = false;
  const panel = document.getElementById("person-filter-panel");
  const button = document.getElementById("btn-persons");
  if (panel) panel.hidden = true;
  button?.setAttribute("aria-expanded", "false");
}

async function togglePersonPanel(): Promise<void> {
  if (personPanelOpen) {
    closePersonPanel();
    return;
  }

  closeTagPanel();
  const panel = document.getElementById("person-filter-panel");
  const button = document.getElementById("btn-persons");
  if (!panel) return;

  personPanelOpen = true;
  panel.hidden = false;
  panel.textContent = "Personen werden geladen …";
  button?.setAttribute("aria-expanded", "true");

  try {
    const response = await fetchAgentWithStartupRetry("/persons");
    if (!personPanelOpen) return;

    const result = await response.json();
    const persons: PersonInfo[] = Array.isArray(result.items) ? result.items : [];
    setInfoText("persons-count", result.count ?? persons.length);
    renderPersonPanel(persons);
  } catch (error) {
    console.error("Personen konnten nicht geladen werden:", error);
    if (personPanelOpen) {
      panel.textContent = "Personen nicht verfügbar.";
    }
  }
}

function renderPersonPanel(persons: PersonInfo[]): void {
  const panel = document.getElementById("person-filter-panel");
  if (!panel) return;
  panel.innerHTML = "";

  const allButton = document.createElement("button");
  allButton.type = "button";
  allButton.className = "person-filter-item" + (activePersons.length === 0 ? " active" : "");
  const allLabel = document.createElement("span");
  allLabel.textContent = "Alle Personen";
  allButton.appendChild(allLabel);
  allButton.onclick = (event) => {
    event.stopPropagation();
    setActivePersons([]);
    renderPersonPanel(persons);
    void refreshSearchForCurrentState();
  };
  panel.appendChild(allButton);

  for (const person of persons) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "person-filter-item" +
      (isPersonActive(person.normalizedName) ? " active" : "");

    const name = document.createElement("span");
    name.textContent = "@" + person.name;
    button.appendChild(name);

    const count = document.createElement("span");
    count.className = "person-filter-item-count";
    count.textContent = String(person.count ?? 0);
    button.appendChild(count);

    button.onclick = (event) => {
      event.stopPropagation();
      togglePersonSelection(person);
      renderPersonPanel(persons);
      void refreshSearchForCurrentState();
    };

    panel.appendChild(button);
  }
}

async function refreshCurrentNotePersons(): Promise<void> {
  const container = document.getElementById("note-persons");
  if (!container) return;

  if (!currentNoteExists || !currentMailNotesId) {
    if (container.childElementCount > 0 || !container.hidden) {
      container.innerHTML = "";
      container.hidden = true;
    }
    return;
  }

  try {
    const response = await fetch(
      AgentUrl + "/note/persons?mailNotesId=" + encodeURIComponent(currentMailNotesId)
    );
    if (!response.ok) return;

    const result = await response.json();
    const persons: PersonInfo[] = Array.isArray(result.items) ? result.items : [];

    const newSignature = persons.map((person) => person.normalizedName || person.name).join("\u0001");
    const currentSignature = Array.from(
      container.querySelectorAll<HTMLButtonElement>(".note-person")
    )
      .map((button) => (button.dataset.normalizedName || button.textContent || "").replace(/^@/, ""))
      .join("\u0001");

    if (newSignature === currentSignature) {
      return;
    }

    if (persons.length === 0) {
      container.innerHTML = "";
      container.hidden = true;
      return;
    }

    const fragment = document.createDocumentFragment();
    for (const person of persons) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "note-person";
      button.dataset.normalizedName = person.normalizedName || person.name;
      button.textContent = "@" + person.name;
      button.title = "Nach @" + person.name + " filtern";
      button.onclick = (event) => {
        event.stopPropagation();
        togglePersonSelection(person);
        void refreshSearchForCurrentState();
        document.querySelector(".search-card")?.scrollIntoView({ block: "nearest" });
      };
      fragment.appendChild(button);
    }

    container.replaceChildren(fragment);
    container.hidden = false;
  } catch {
    // Personen-Anzeige ist Komfortfunktion.
  }
}

async function toggleFavoritesFilter(): Promise<void> {
  // Favoriten ist wie Tags/Personen ein eigener Suchfilter.
  // Eine eventuell geöffnete Auswahlliste gehört beim Wechsel hierher geschlossen.
  closeTagPanel();
  closePersonPanel();

  setFavoritesFilterActive(!favoritesFilterActive);
  await refreshSearchForCurrentState();

  const input = document.getElementById("search-input") as HTMLInputElement | null;
  input?.focus();
}

function setFavoritesFilterActive(active: boolean): void {
  favoritesFilterActive = active;
  const button = document.getElementById("btn-favorites");
  button?.classList.toggle("active", favoritesFilterActive);
  button?.setAttribute("aria-pressed", favoritesFilterActive ? "true" : "false");
}

async function refreshSearchForCurrentState(): Promise<void> {
  if (searchTimer !== undefined) {
    window.clearTimeout(searchTimer);
    searchTimer = undefined;
  }

  const input = document.getElementById("search-input") as HTMLInputElement | null;
  const searchText = input?.value.trim() ?? "";

  if (searchText || favoritesFilterActive || activeTags.length > 0 || activePersons.length > 0) {
    await searchNotes(
      searchText,
      favoritesFilterActive,
      activeTags.map((tag) => tag.normalizedName),
      activePersons.map((person) => person.normalizedName)
    );
  } else {
    searchSequence++;
    clearSearchResults();
  }
}

async function loadFavorites(): Promise<void> {
  const sequence = ++searchSequence;
  setText("search-status", "Favoriten …");

  try {
    const response = await fetch(AgentUrl + "/favorites?limit=100");
    if (!response.ok) throw new Error("Agent returned HTTP " + response.status);
    const result = await response.json();
    if (sequence !== searchSequence || !favoritesFilterActive) return;

    const items: SearchResultItem[] = Array.isArray(result.items) ? result.items : [];
    setInfoText("favorites-count", result.count ?? items.length);
    renderSearchResults(items, "Favorit");
  } catch (error) {
    console.error("Favoriten konnten nicht geladen werden:", error);
    if (sequence === searchSequence) {
      clearSearchResults();
      setText("search-status", "Favoriten nicht verfügbar.");
    }
  }
}

type MailNotesStats = {
  notes: number;
  mailLinks: number;
  favorites: number;
  tags: number;
  persons: number;
  version: string;
};


async function loadDatabasePath(): Promise<string> {
  const response = await fetch(AgentUrl + "/version");
  if (!response.ok) throw new Error("Agent returned HTTP " + response.status);
  const info = await response.json();
  const path = (info.databasePath || "").toString();
  setInfoText("info-database-path", path || "–");
  return path;
}

function setupDatabasePathButton(): void {
  const button = document.getElementById("btn-database-path") as HTMLButtonElement | null;
  if (!button) return;

  button.onclick = async () => {
    try {
      const currentPath = await loadDatabasePath();
      const requestedPath = window.prompt(
        "Neuer Datenbankpfad oder Zielordner:\n\n" +
        "Existiert dort bereits eine MailNotes.sqlite, wird diese übernommen. " +
        "Andernfalls wird die aktuelle Datenbank dorthin kopiert.",
        currentPath
      );

      if (requestedPath === null || requestedPath.trim() === "") return;

      button.disabled = true;
      setInfoText("info-status", "Datenbank wird umgestellt …");

      const body = new URLSearchParams();
      body.append("path", requestedPath.trim());
      const response = await fetch(AgentUrl + "/database/path", { method: "POST", body });
      const result = await response.json();

      if (!response.ok || !result.ok) {
        throw new Error(result.message || "Datenbankpfad konnte nicht geändert werden.");
      }

      setInfoText("info-database-path", result.databasePath || requestedPath.trim());
      if (result.mode === "adopted") {
        setInfoText("info-status", "Vorhandene Datenbank übernommen.");
      } else if (result.mode === "moved") {
        setInfoText("info-status", "Datenbank verschoben.");
      } else {
        setInfoText("info-status", "Datenbankpfad unverändert.");
      }

      await loadNote(itemChangeSequence);
      void refreshStatisticsCounts();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setInfoText("info-status", message);
    } finally {
      button.disabled = false;
    }
  };
}

function setupInfoPanel(): void {
  const button = document.getElementById("btn-info");
  const card = document.getElementById("info-card");

  if (!button || !card) {
    return;
  }

  const closeInfoPanel = () => {
    card.hidden = true;
    button.setAttribute("aria-expanded", "false");
    button.classList.remove("active");
  };

  button.onclick = () => {
    const opening = card.hidden;

    if (opening) {
      card.hidden = false;
      button.setAttribute("aria-expanded", "true");
      button.classList.add("active");
      void loadStatistics();
      void loadDatabasePath();
    } else {
      closeInfoPanel();
    }
  };

  document.addEventListener("click", (event) => {
    const target = event.target as Node | null;
    if (!target || card.hidden) {
      return;
    }

    if (!card.contains(target) && !button.contains(target)) {
      closeInfoPanel();
    }
  });
}

async function loadStatistics(): Promise<void> {
  setInfoText("info-status", "Wird aktualisiert …");

  try {
    const response = await fetch(AgentUrl + "/stats");

    if (!response.ok) {
      throw new Error("Agent returned HTTP " + response.status);
    }

    const stats = await response.json() as MailNotesStats;
    setInfoText("info-notes", stats.notes);
    setInfoText("info-mail-links", stats.mailLinks);
    setInfoText("info-favorites", stats.favorites);
    setInfoText("info-tags", stats.tags);
    setInfoText("info-persons", stats.persons);
    setInfoText("favorites-count", stats.favorites);
    setInfoText("tags-count", stats.tags);
    setInfoText("persons-count", stats.persons);
    setInfoText("info-version", stats.version || "–");
    setInfoText("info-build", __MAILNOTES_BUILD_TIME__ || "–");
    setInfoText("info-commit", __MAILNOTES_BUILD_COMMIT__ || "–");
    setInfoText("info-status", "");
  } catch (error) {
    console.error("MailNotes-Statistik konnte nicht geladen werden:", error);
    setInfoText("info-status", "Informationen nicht verfügbar.");
  }
}

function setInfoText(id: string, value: unknown): void {
  const element = document.getElementById(id);
  if (!element) {
    return;
  }

  element.textContent = value === null || value === undefined || value === ""
    ? "–"
    : String(value);
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
  if (!trimmed && !favoritesFilterActive && activeTags.length === 0 && activePersons.length === 0) {
    clearSearchResults();
    return;
  }

  setText(
    "search-status",
    favoritesFilterActive || activeTags.length > 0 || activePersons.length > 0 ? "Gefilterte Suche …" : "Suche …"
  );
  searchTimer = window.setTimeout(() => {
    void searchNotes(trimmed, favoritesFilterActive, activeTags.map((tag) => tag.normalizedName), activePersons.map((person) => person.normalizedName));
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

  setFavoritesFilterActive(false);
  setActiveTags([]);
  setActivePersons([]);
  closeTagPanel();
  closePersonPanel();
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

async function searchNotes(
  searchText: string,
  favoriteOnly: boolean = false,
  tagNormalizedNames: string[] = [],
  personNormalizedNames: string[] = []
): Promise<void> {
  const sequence = ++searchSequence;

  try {
    const response = await fetch(
      AgentUrl + "/search?q=" + encodeURIComponent(searchText) + "&limit=50" +
      (favoriteOnly ? "&favorite=1" : "") +
      (tagNormalizedNames.length > 0 ? "&tags=" + encodeURIComponent(tagNormalizedNames.join("|")) : "") +
      (personNormalizedNames.length > 0 ? "&persons=" + encodeURIComponent(personNormalizedNames.join("|")) : "")
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

    renderSearchResults(
      items,
      favoriteOnly && tagNormalizedNames.length === 0 && personNormalizedNames.length === 0 ? "Favorit" : "Treffer"
    );
  } catch (error) {
    console.error("MailNotes-Suche fehlgeschlagen:", error);
    if (sequence === searchSequence) {
      clearSearchResults();
      setText("search-status", "Suche nicht verfügbar.");
    }
  }
}

function renderSearchResults(items: SearchResultItem[], singularLabel: string = "Treffer"): void {
  const results = document.getElementById("search-results");
  if (!results) {
    return;
  }

  results.innerHTML = "";
  results.hidden = false;

  if (items.length === 0) {
    setText("search-status", singularLabel === "Favorit" ? "Keine Favoriten." : "Keine Treffer.");
    return;
  }

  setText(
    "search-status",
    items.length === 1
      ? "1 " + singularLabel
      : items.length + " " + (singularLabel === "Favorit" ? "Favoriten" : "Treffer")
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
  autosaveSuppressed = true;

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
    currentMailKnown = Boolean(currentMailNotesId);
    currentGmlActive = Boolean(note.gmlActive);
    setShlStatus(repairQueueCount > 0 ? "pending" : "active");

    const identity = getCurrentMailSnapshot();

    // Technische IDs nur fuer bereits in MailNotes bekannte Mails
    // aktualisieren. Die reine Auswahl einer unbekannten Outlook-Mail darf
    // keinen Mail-Datensatz erzeugen.
    if (identity && currentMailKnown) {
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

    currentNoteExists = Boolean(note.found);
    currentNoteIsFavorite = Boolean(note.isFavorite);
    updateFavoriteButton();

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
    autosaveSuppressed = false;
    setAutosaveStatus("");

    await renderLinks(expectedSequence);

    if (expectedSequence !== itemChangeSequence) {
      return;
    }

    await renderBacklinks(
      currentMailNotesId || messageId,
      expectedSequence
    );

    await refreshCurrentNoteTags();
    await refreshCurrentNotePersons();
    void synchronizeMailNotesCategoryForCurrentItem();

  } catch (error) {
    if (expectedSequence !== itemChangeSequence) {
      return;
    }

    const errorMessage =
      error instanceof Error
        ? error.message
        : String(error);

    setEditorText(
      "note-content",
      "MailNotesAgent-Fehler: " + errorMessage
    );

    setEditorText(
      "note-links",
      ""
    );

    setNoteMeta("", "");
    setAutosaveStatus("Laden fehlgeschlagen");

    await renderLinks(expectedSequence);
    clearBacklinks();

    fileLog("loadNote error", {
      messageId,
      error: errorMessage
    });

    console.error("loadNote fehlgeschlagen:", error);
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
  if (autosaveSuppressed) {
    return false;
  }

  if (autosaveRevision <= autosaveCompletedRevision) {
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
            currentMailKnown = true;
          }

          if (typeof result.gmlActive === "boolean") {
            currentGmlActive = result.gmlActive;
            setShlStatus(repairQueueCount > 0 ? "pending" : "active");
          }

          updateNoteMetaAfterSave(result);
          currentNoteExists = true;
          invalidateNoteAutocompleteCache();
          updateFavoriteButton();
          void refreshCurrentNoteTags();
          void refreshCurrentNotePersons();
          void refreshStatisticsCounts();
          void synchronizeMailNotesCategoryForCurrentItem();
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

  let response: Response;

  try {
    response = await fetch(url, {
      targetAddressSpace: "loopback"
    } as RequestInit);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : String(error);

    throw new Error(
      "GET /note: Netzwerkfehler – " + message
    );
  }

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
      "GET /note: HTTP " +
      response.status +
      (responseText
        ? " – " + responseText
        : "")
    );
  }

  try {
    return JSON.parse(responseText);
  } catch {
    throw new Error(
      "GET /note: Ungültige JSON-Antwort – " +
      responseText
    );
  }
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
    normalizeOfficeText(item.subject);

  const senderName =
    normalizeOfficeText((item as any).from?.displayName);

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

async function deleteBacklink(
  sourceMailNotesId: string,
  targetMailIdentity: string,
  sourceLink: string
): Promise<void> {
  const url =
    AgentUrl +
    "/backlinks?sourceMailNotesId=" +
    encodeURIComponent(sourceMailNotesId) +
    "&targetMailIdentity=" +
    encodeURIComponent(targetMailIdentity) +
    "&sourceLink=" +
    encodeURIComponent(sourceLink);

  const response =
    await fetch(
      url,
      { method: "DELETE" }
    );

  if (!response.ok) {
    const responseText =
      await response.text();

    throw new Error(
      "Backlink konnte nicht gelöscht werden (HTTP " +
      response.status +
      "): " +
      responseText
    );
  }
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

      const deleteButton =
        document.createElement("button");

      deleteButton.className =
        "link-action";

      deleteButton.type =
        "button";

      deleteButton.title =
        "Backlink entfernen";

      deleteButton.innerHTML =
        '<span class="icon-delete">×</span>';

      deleteButton.onclick = async () => {
        deleteButton.disabled = true;

        try {
          await deleteBacklink(
            item.mailNotesId || "",
            mailIdentity,
            item.sourceLink || ""
          );

          if (
            expectedSequence !== itemChangeSequence ||
            getCurrentMailIdentity() !== mailIdentity
          ) {
            return;
          }

          await renderBacklinks(
            mailIdentity,
            expectedSequence
          );

          setText(
            "mail-link-status",
            "Backlink gelöscht."
          );

          window.setTimeout(() => {
            if (
              expectedSequence === itemChangeSequence &&
              getCurrentMailIdentity() === mailIdentity
            ) {
              setText(
                "mail-link-status",
                ""
              );
            }
          }, 2000);
        } catch (error) {
          deleteButton.disabled = false;
          setText(
            "mail-link-status",
            "Backlink konnte nicht gelöscht werden."
          );
          console.error(error);
        }
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

    console.info(
      "[MailNotes][Resolve]",
      {
        url,
        found: resolved.found,
        type: resolved.type,
        mailNotesId: (resolved as any).mailNotesId || "",
        mailboxAddress: (resolved as any).mailboxAddress || "",
        itemId: storedItemId,
        itemIdLength: storedItemId.length
      }
    );

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

    console.info(
      "[MailNotes][Open] Item-ID-Kandidaten",
      candidates.map((candidate, index) => ({
        index,
        itemId: candidate,
        length: candidate.length,
        isResolvedItemId: candidate === storedItemId
      }))
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
                  "[MailNotes][Open] displayMessageFormAsync fehlgeschlagen",
                  {
                    candidateIndex: index,
                    candidate,
                    candidateLength: candidate.length,
                    errorCode: result.error?.code,
                    errorName: result.error?.name,
                    errorMessage: result.error?.message,
                    error: result.error
                  }
                );

                tryCandidate(index + 1);
                return;
              }

              console.info(
                "[MailNotes][Open] displayMessageFormAsync erfolgreich",
                {
                  candidateIndex: index,
                  candidate,
                  candidateLength: candidate.length
                }
              );

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

function getRepairQueueUrl(path: "" | "/count"): string {
  const mailboxAddress = getCurrentMailSnapshot()?.mailboxAddress.trim() || "";
  const params = new URLSearchParams();

  if (mailboxAddress) {
    params.append("mailboxAddress", mailboxAddress);
  }

  return AgentUrl + "/repairqueue" + path +
    (params.toString() ? "?" + params.toString() : "");
}

async function refreshRepairQueueNotice(forceShow = false): Promise<number> {
  try {
    const response = await fetch(getRepairQueueUrl("/count"));
    if (!response.ok) return -1;

    const result = await response.json();
    const count = Number(result.count || 0);
    repairQueueCount = count;

    if (typeof result.gmlActive === "boolean") {
      currentGmlActive = result.gmlActive;
    }
    const notice = document.getElementById("repair-queue-notice");

    if (count === 0) {
      finishRepairQueue();
      return 0;
    }

    setShlStatus("pending");

    if (forceShow) {
      repairHintDismissed = false;
    }

    // Der Queue-Zaehler im SHL-Status ist global. Der grosse Hinweis darf
    // deshalb nur unmittelbar nach einer neu erkannten defekten Mail-Zuordnung
    // erscheinen, nicht beim Start oder beim Wechsel auf eine andere Mail.
    if (notice) {
      notice.hidden = !forceShow || repairHintDismissed;
    }

    setText(
      "repair-queue-notice-text",
      "Die gespeicherte Outlook-Zuordnung dieser Mail ist nicht mehr gültig. " +
      "Das kann zum Beispiel nach dem Verschieben der Mail passieren. " +
      "Die betroffene Mail suchen und in einem separaten Fenster öffnen."
    );

    return count;
  } catch {
    // Der Hinweis ist Komfortfunktion; Agentfehler werden über SHL angezeigt.
    return -1;
  }
}

function stopRepairWatch(): void {
  if (repairWatchTimer !== undefined) {
    window.clearTimeout(repairWatchTimer);
    repairWatchTimer = undefined;
  }

  repairWatchDeadline = 0;
  repairWatchItemIds.clear();
}

function scheduleRepairWatch(): void {
  if (repairWatchTimer !== undefined) {
    window.clearTimeout(repairWatchTimer);
  }

  repairWatchTimer = window.setTimeout(() => {
    repairWatchTimer = undefined;
    void pollRepairWatch();
  }, REPAIR_WATCH_INTERVAL_MS);
}

async function startRepairWatch(): Promise<void> {
  stopRepairWatch();

  try {
    const response = await fetch(getRepairQueueUrl(""));
    if (!response.ok) {
      return;
    }

    const result = await response.json();
    const items = (result.items || []) as RepairQueueItem[];

    if (items.length === 0) {
      finishRepairQueue();
      return;
    }

    repairWatchItemIds = new Set(items.map((item) => item.id));
    repairWatchDeadline = Date.now() + REPAIR_WATCH_TIMEOUT_MS;
    scheduleRepairWatch();
  } catch {
    stopRepairWatch();
  }
}

async function pollRepairWatch(): Promise<void> {
  if (repairWatchDeadline === 0) {
    return;
  }

  if (Date.now() >= repairWatchDeadline) {
    stopRepairWatch();
    return;
  }

  try {
    const response = await fetch(getRepairQueueUrl(""));
    if (!response.ok) {
      scheduleRepairWatch();
      return;
    }

    const result = await response.json();
    const items = (result.items || []) as RepairQueueItem[];
    const currentIds = new Set(items.map((item) => item.id));
    const repaired = Array.from(repairWatchItemIds).some((id) => !currentIds.has(id));

    if (repaired) {
      repairQueueCount = items.length;
      setShlStatus("repaired");

      if (items.length === 0) {
        finishRepairQueue();
        return;
      }

      // Eine erfolgreiche Reparatur ist Fortschritt: Für den nächsten
      // Queue-Eintrag beginnt das Sicherheits-Timeout erneut.
      repairWatchItemIds = currentIds;
      repairWatchDeadline = Date.now() + REPAIR_WATCH_TIMEOUT_MS;
      await showRepairQueue();
    }

    scheduleRepairWatch();
  } catch {
    scheduleRepairWatch();
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
  stopRepairWatch();
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
  const response = await fetch(getRepairQueueUrl(""));
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
    searchButton.textContent = "Mail suchen und neu zuordnen";

    const searchText = buildOutlookSearchText(item);
    if (!searchText) {
      searchButton.disabled = true;
      searchButton.title = "Für diese Mail sind weder Betreff, Absender noch Empfangsdatum gespeichert.";
      searchFeedback.textContent =
        "Für diese Mail konnte kein brauchbarer Outlook-Suchtext erzeugt werden. " +
        "Bitte anhand der angezeigten Angaben manuell suchen und die Mail anschließend in einem eigenen Fenster öffnen.";
      searchFeedback.hidden = false;
    } else {
      searchButton.onclick = async () => {
        if (!navigator.clipboard) {
          searchFeedback.textContent =
            "Der Suchtext konnte nicht automatisch kopiert werden: „" + searchText + "“. " +
            "Bitte den Suchtext manuell in Outlook eingeben und die passende Mail in einem eigenen Fenster öffnen.";
          searchFeedback.hidden = false;
          return;
        }

        await navigator.clipboard.writeText(searchText);

        searchButton.textContent = "Suchbegriff erneut kopieren";
        searchFeedback.textContent =
          "Für die Outlook-Suche in die Zwischenablage kopiert: „" + searchText + "“. " +
          "Oben im Suchfeld in Outlook einfügen, die passende Mail suchen und anschließend in einem eigenen Fenster öffnen. " +
          "MailNotes aktualisiert die Mail-Zuordnung dann automatisch.";
        searchFeedback.hidden = false;
      };
    }
    actions.appendChild(searchButton);

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.textContent = "MailNotes-Eintrag löschen …";
    deleteButton.onclick = async () => {
      const confirmed = await confirmNoteDeletion(
        "MailNotes-Eintrag wirklich löschen?",
        "MailNotes kann nicht sicher feststellen, ob die Outlook-Mail gelöscht oder nur verschoben wurde. " +
        "Wenn die Mail tatsächlich nicht mehr existiert, werden Notiz, Links, Favorit sowie Tag-/Personen-Zuordnungen entfernt."
      );
      if (!confirmed) return;

      deleteButton.disabled = true;
      searchButton.disabled = true;

      try {
        stopRepairWatch();
        const params = new URLSearchParams();
        params.append("mailNotesId", item.mailNotesId);
        if (item.messageId) params.append("messageId", item.messageId);

        const deleteResponse = await fetch(AgentUrl + "/note?" + params.toString(), {
          method: "DELETE"
        });
        if (!deleteResponse.ok) {
          throw new Error("Agent returned HTTP " + deleteResponse.status);
        }

        invalidateNoteAutocompleteCache();
        await refreshStatisticsCounts();
        await refreshSearchForCurrentState();

        if (currentMailNotesId && currentMailNotesId === item.mailNotesId) {
          await loadNote(itemChangeSequence);
        }

        await showRepairQueue();
        await refreshRepairQueueNotice(false);
      } catch (error) {
        console.error("MailNotes-Eintrag konnte nicht gelöscht werden:", error);
        window.alert("Der MailNotes-Eintrag konnte nicht gelöscht werden.");
        deleteButton.disabled = false;
        searchButton.disabled = false;
      }
    };
    actions.appendChild(deleteButton);

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
    parts.push("subject:(" + subject + ")");
  }

  if (sender) {
    parts.push("from:(" + sender + ")");
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
