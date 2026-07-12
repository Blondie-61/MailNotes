/* global Office, document, window, navigator */

const AgentUrl = "/api";

const DEBUG = true;

function log(...args: any[]) {
  if (DEBUG) {
    console.log(...args);
  }
}

Office.onReady((info) => {
  log("Office.onReady", info);

  if (info.host === Office.HostType.Outlook) {
    void runOutlook();
  }
});

async function runOutlook() {
  showMailInformation();
  setupButtons();

  await loadNote();
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
      addLinkFromInput();
    };
  }

  if (btnCopyMailLink) {
    btnCopyMailLink.onclick = () => {
      void copyCurrentMailLink();
    };
  }

  if (linkInput) {
    linkInput.onkeydown = (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        addLinkFromInput();
      }
    };
  }
}

function showMailInformation() {
  const item =
    Office.context.mailbox.item;

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

async function loadNote() {
  const item =
    Office.context.mailbox.item;

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

    await renderLinks();
    clearBacklinks();

    return;
  }

  try {
    const note =
      await getNote(messageId);

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

    await renderLinks();
    await renderBacklinks(messageId);

  } catch (error) {
    setEditorText(
      "note-content",
      "MailNotesAgent nicht erreichbar."
    );

    setEditorText(
      "note-links",
      ""
    );

    setNoteMeta("", "");

    await renderLinks();
    clearBacklinks();

    console.error(error);
  }
}

async function saveNote() {
  const item =
    Office.context.mailbox.item;

  const messageId =
    (item as any).internetMessageId;

  const conversationId =
    (item as any).conversationId;

  const subject =
    item.subject || "";

  const senderName =
    (item as any).from?.displayName || "";

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

  body.append(
    "messageId",
    messageId
  );

  body.append(
    "conversationId",
    conversationId || ""
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

function addLinkFromInput() {
  const input =
    document.getElementById(
      "link-input"
    ) as HTMLInputElement;

  if (!input) {
    return;
  }

  const value =
    input.value.trim();

  if (!value) {
    return;
  }

  const links =
    getLinks();

  links.push(value);

  setLinks(links);

  input.value = "";

  void renderLinks();
}

async function copyCurrentMailLink() {
  const item =
    Office.context.mailbox.item;

  const messageId =
    (item as any).internetMessageId;

  if (!messageId) {
    setText(
      "mail-link-status",
      "Keine Message-ID vorhanden."
    );

    return;
  }

  const mailLink =
    "mailnotes:" +
    encodeURIComponent(messageId);

  try {
    await navigator.clipboard.writeText(
      mailLink
    );

    setText(
      "mail-link-status",
      "Kopiert."
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
      "Kopieren fehlgeschlagen."
    );

    console.error(error);
  }
}

async function renderLinks() {
  const list =
    document.getElementById(
      "links-list"
    );

  if (!list) {
    return;
  }

  list.innerHTML = "";

  const links =
    getLinks();

  for (
    let index = 0;
    index < links.length;
    index++
  ) {
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

    deleteButton.onclick = () => {
      deleteLink(index);
    };

    actions.appendChild(
      copyButton
    );

    actions.appendChild(
      deleteButton
    );

    row.appendChild(open);
    row.appendChild(actions);

    list.appendChild(row);

    if (
      url
        .toLowerCase()
        .startsWith("mailnotes:")
    ) {
      await resolveRenderedMailLink(
        url,
        title,
        subtitle
      );
    }
  }
}

async function resolveRenderedMailLink(
  url: string,
  title: HTMLElement,
  subtitle: HTMLElement
) {
  try {
    const resolved =
      await resolveLink(url);

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
  messageId: string
): Promise<any> {

  const url =
    AgentUrl +
    "/backlinks?messageId=" +
    encodeURIComponent(messageId);

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
  messageId: string
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

  list.innerHTML = "";
  section.hidden = true;

  try {
    const result =
      await getBacklinks(messageId);

    const items =
      Array.isArray(result.items)
        ? result.items
        : [];

    if (items.length === 0) {
      return;
    }

    section.hidden = false;

    for (const item of items) {
      const sourceMessageId =
        item.messageId || "";

      const mailLink =
        "mailnotes:" +
        encodeURIComponent(
          sourceMessageId
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
        sourceMessageId;

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

      list.appendChild(row);
    }

  } catch (error) {
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
  const messageId =
    decodeMailNotesMessageId(url);

  try {
    await navigator.clipboard.writeText(
      messageId
    );

    setText(
      "mail-link-status",
      "Direktes Öffnen folgt noch – Message-ID wurde kopiert."
    );

    window.setTimeout(() => {
      setText(
        "mail-link-status",
        ""
      );
    }, 4000);

  } catch (error) {
    setText(
      "mail-link-status",
      "Direktes Öffnen folgt noch – Kopieren fehlgeschlagen."
    );

    console.error(error);
  }
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

function deleteLink(
  index: number
) {
  const links =
    getLinks();

  links.splice(
    index,
    1
  );

  setLinks(links);

  void renderLinks();
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