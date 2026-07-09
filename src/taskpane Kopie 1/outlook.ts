/* global Office, document */

const AgentUrl = "/api";

const DEBUG = false;

function log(...args: any[]) {
  if (DEBUG) {
    console.log(...args);
  }
}

Office.onReady((info) => {
  log("Office.onReady", info);

  if (info.host === Office.HostType.Outlook) {
    runOutlook();
  }
});

async function runOutlook() {
  showMailInformation();
  setupButtons();
  await loadNote();
}

function setupButtons() {
  const btnSave = document.getElementById("btn-save");

  if (btnSave) {
    btnSave.onclick = async () => {
      log("Speichern geklickt");
      await saveNote();
      await loadNote();
    };
  }
}

function showMailInformation() {
  const item = Office.context.mailbox.item;

  setText("mail-subject", item.subject);
  setText("mail-from", (item as any).from?.displayName);
  setText("mail-date", formatDate((item as any).dateTimeCreated));

  setText("mail-message-id", (item as any).internetMessageId);
  setText("mail-conversation-id", (item as any).conversationId);
  setText("mail-item-id", (item as any).itemId);
}

async function loadNote() {
  const item = Office.context.mailbox.item;
  const messageId = (item as any).internetMessageId;

  if (!messageId) {
    setEditorText("note-content", "Keine Message-ID vorhanden.");
    setEditorText("note-links", "");
    setNoteMeta("", "");
    return;
  }

  try {
    const note = await getNote(messageId);

    if (note.found) {
      setEditorText("note-content", note.content);
      setEditorText("note-links", note.links);
      setNoteMeta(note.createdAt, note.modifiedAt);
    } else {
      setEditorText("note-content", "");
      setEditorText("note-links", "");
      setNoteMeta("", "");
    }
  } catch (error) {
    setEditorText("note-content", "MailNotesAgent nicht erreichbar.");
    setEditorText("note-links", "");
    setNoteMeta("", "");
    console.error(error);
  }
}

async function saveNote() {
  log("saveNote() gestartet");

  const item = Office.context.mailbox.item;

  const messageId = (item as any).internetMessageId;
  const conversationId = (item as any).conversationId;
  const content = getEditorText("note-content");
  const links = getEditorText("note-links");

  log("messageId:", messageId);
  log("content:", content);
  log("links:", links);

  if (!messageId) {
    log("keine messageId");
    return;
  }

  const body = new URLSearchParams();
  body.append("messageId", messageId);
  body.append("conversationId", conversationId || "");
  body.append("content", content);
  body.append("links", links);

  log("POST body:", body.toString());

  const response = await fetch(AgentUrl + "/note", {
    method: "POST",
    body: body
  });

  const responseText = await response.text();
  log("POST response:", responseText);

  if (!response.ok) {
    throw new Error("Agent returned HTTP " + response.status);
  }
}

async function getNote(messageId: string): Promise<any> {
  const url =
    AgentUrl +
    "/note?messageId=" +
    encodeURIComponent(messageId);

  const response = await fetch(url);

  if (!response.ok) {
    throw new Error("Agent returned HTTP " + response.status);
  }

  return await response.json();
}

function getEditorText(id: string): string {
  const element = document.getElementById(id) as HTMLTextAreaElement;

  if (!element) {
    return "";
  }

  return element.value;
}

function setEditorText(id: string, value: any) {
  const element = document.getElementById(id) as HTMLTextAreaElement;

  if (!element) {
    return;
  }

  element.value = value ? value.toString() : "";
}

function setNoteMeta(createdAt: any, modifiedAt: any) {
  setText("note-created", formatDate(createdAt));
  setText("note-modified", formatDate(modifiedAt));
}

function formatDate(value: any): string {
  if (!value) {
    return "–";
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return value.toString();
  }

  return date
    .toLocaleString("de-DE", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    })
    .replace(",", "");
}

function setText(id: string, value: any) {
  const element = document.getElementById(id);

  if (!element) {
    return;
  }

  element.textContent = value ? value.toString() : "–";
}