// ============================================================
// graph.js — Appels API Microsoft Graph
// Outlook Calendar + SharePoint Excel + Mail
// ============================================================

const Graph = (() => {

  const BASE = "https://graph.microsoft.com/v1.0";

  async function call(endpoint, options = {}) {
    const token = await Auth.getToken();
    const res = await fetch(BASE + endpoint, {
      ...options,
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {})
      }
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `HTTP ${res.status}`);
    }
    if (res.status === 204) return null;
    return res.json();
  }

  // ----------------------------------------------------------
  // TESTS DE CONNEXION
  // ----------------------------------------------------------
  async function testConnection() {
    await call("/me");
    return true;
  }

  async function testCalendarAccess() {
    await call("/me/calendars?$top=1");
    return true;
  }

  async function testSharePointAccess() {
    const siteHostname = new URL(CONFIG.sharepoint.siteUrl).hostname;
    const sitePath = new URL(CONFIG.sharepoint.siteUrl).pathname;
    await call(`/sites/${siteHostname}:${sitePath}`);
    return true;
  }

  async function testMailAccess() {
    // Juste vérifier que le scope Mail.Send est accessible
    await call("/me/mailboxSettings");
    return true;
  }

  // ----------------------------------------------------------
  // MEMBRES DU GROUPE (agents du service IT)
  // ----------------------------------------------------------
  async function getGroupMembers() {
    if (!CONFIG.agentsGroupId) {
      throw new Error("agentsGroupId non configuré");
    }
    const res = await call(`/groups/${CONFIG.agentsGroupId}/members?$select=id,displayName,mail,jobTitle`);
    return res.value || [];
  }

  // ----------------------------------------------------------
  // CALENDRIER OUTLOOK
  // ----------------------------------------------------------
  async function getOrCreateCalendar(calendarName) {
    const res = await call("/me/calendars");
    const existing = res.value.find(c => c.name === calendarName);
    if (existing) return existing.id;

    // Créer le calendrier s'il n'existe pas
    const newCal = await call("/me/calendars", {
      method: "POST",
      body: JSON.stringify({ name: calendarName })
    });
    return newCal.id;
  }

  async function createCalendarEvent({ subject, start, end, attendeeEmail, body }) {
    const calId = await getOrCreateCalendar(CONFIG.outlook.calendarName);

    // Convertir les dates en ISO (journée entière)
    const startDT = `${start}T08:00:00`;
    const endDT   = `${end}T18:00:00`;

    const event = {
      subject,
      body: { contentType: "text", content: body },
      start: { dateTime: startDT, timeZone: "Europe/Paris" },
      end:   { dateTime: endDT,   timeZone: "Europe/Paris" },
      isAllDay: false,
      showAs: "oof",   // "Out of Office" dans Outlook
      categories: ["Congés IT"],
      attendees: attendeeEmail ? [{
        emailAddress: { address: attendeeEmail },
        type: "optional"
      }] : []
    };

    return call(`/me/calendars/${calId}/events`, {
      method: "POST",
      body: JSON.stringify(event)
    });
  }

  // ----------------------------------------------------------
  // SHAREPOINT EXCEL
  // ----------------------------------------------------------
  async function getSharePointDriveItemId() {
    // Résoudre l'ID du site
    const siteHostname = new URL(CONFIG.sharepoint.siteUrl).hostname;
    const sitePath = new URL(CONFIG.sharepoint.siteUrl).pathname;
    const site = await call(`/sites/${siteHostname}:${sitePath}`);
    const siteId = site.id;

    // Trouver le fichier dans le drive du site
    const drive = await call(`/sites/${siteId}/drive/root:/${CONFIG.sharepoint.fileName}`);
    return { siteId, driveId: drive.parentReference.driveId, itemId: drive.id };
  }

  async function appendToSharePointExcel({ agent, type, debut, fin, duree, statut }) {
    const { siteId, driveId, itemId } = await getSharePointDriveItemId();

    // Récupérer les données existantes pour trouver la prochaine ligne libre
    const existingData = await call(
      `/sites/${siteId}/drives/${driveId}/items/${itemId}/workbook/worksheets('${CONFIG.sharepoint.sheetName}')/usedRange`
    ).catch(() => ({ rowCount: 1 }));

    const nextRow = (existingData.rowCount || 1) + 1;
    const range = `A${nextRow}:F${nextRow}`;

    // Écrire la nouvelle ligne
    await call(
      `/sites/${siteId}/drives/${driveId}/items/${itemId}/workbook/worksheets('${CONFIG.sharepoint.sheetName}')/range(address='${range}')`,
      {
        method: "PATCH",
        body: JSON.stringify({
          values: [[agent, type, debut, fin, duree, statut]]
        })
      }
    );

    return true;
  }

  // ----------------------------------------------------------
  // EMAIL DE NOTIFICATION
  // ----------------------------------------------------------
  async function sendNotificationEmail({ to, subject, body }) {
    const recipients = [
      { emailAddress: { address: to } }
    ];
    if (CONFIG.outlook.notifyAgent && CONFIG.outlook.managerEmail !== to) {
      recipients.push({ emailAddress: { address: CONFIG.outlook.managerEmail } });
    }

    const message = {
      subject,
      body: { contentType: "text", content: body },
      toRecipients: recipients
    };

    return call("/me/sendMail", {
      method: "POST",
      body: JSON.stringify({ message, saveToSentItems: true })
    });
  }

  return {
    testConnection,
    testCalendarAccess,
    testSharePointAccess,
    testMailAccess,
    getGroupMembers,
    createCalendarEvent,
    appendToSharePointExcel,
    sendNotificationEmail
  };
})();
