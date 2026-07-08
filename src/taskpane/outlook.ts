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
  setText("mail-date", (item as any).dateTimeCreated);

  setText("mail-message-id", (item as any).internetMessageId);
  setText("mail-conversation-id", (item as any).conversationId);
  setText("mail-item-id", (item as any).itemId);
}

async function loadNote() {
  const item = Office.context.mailbox.item;
  const messageId = (item as any).internetMessageId;

  if (!messageId) {
    setEditorText("Keine Message-ID vorhanden.");
    return;
  }

  try {
    const note = await getNote(messageId);

    if (note.found) {
      setEditorText(note.content);
    } else {
      setEditorText("");
    }
  } catch (error) {
    setEditorText("MailNotesAgent nicht erreichbar.");
    console.error(error);
  }
}

async function saveNote() {
  log("saveNote() gestartet");

  const item = Office.context.mailbox.item;

  const messageId = (item as any).internetMessageId;
  const conversationId = (item as any).conversationId;
  const content = getEditorText();

  log("messageId:", messageId);
  log("content:", content);

  if (!messageId) {
    log("keine messageId");
    return;
  }

  const body = new URLSearchParams();
  body.append("messageId", messageId);
  body.append("conversationId", conversationId || "");
  body.append("content", content);

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

function getEditorText(): string {
  const element = document.getElementById("note-content") as HTMLTextAreaElement;

  if (!element) {
    return "";
  }

  return element.value;
}

function setEditorText(value: any) {
  const element = document.getElementById("note-content") as HTMLTextAreaElement;

  if (!element) {
    return;
  }

  element.value = value ? value.toString() : "";
}

function setText(id: string, value: any) {
  const element = document.getElementById(id);

  if (!element) {
    return;
  }

  element.textContent = value ? value.toString() : "–";
}