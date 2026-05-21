// ============================================================
// config.js — CongesIT
// Personnalisez ce fichier avec vos identifiants Microsoft 365
// ============================================================

const CONFIG = {

  // --- Azure AD (récupérés sur portal.azure.com) ---
  clientId:  "1b1ab90f-46bc-463d-9edb-e52c4d32c00a",         // ID d'application (client)
  tenantId:  "e7f2c162-fd33-4911-941e-5d1a32fc5309",      // ID de l'annuaire (locataire)

  // --- URI de redirection (doit correspondre à Azure AD) ---
  redirectUri: "https://lannemezan-it.github.io/congesit/",

  // --- SharePoint ---
  sharepoint: {
    // URL de votre site SharePoint IT
    siteUrl:    "https://lannemezan.sharepoint.com/sites/support",
    // Nom du fichier Excel de suivi (dans la bibliothèque Documents)
    fileName:   "Conges_IT_2026.xlsx",
    // Nom de la feuille dans le fichier Excel
    sheetName:  "Suivi_Congés",
    // Bibliothèque SharePoint
    driveId:    ""  // Laissez vide, sera résolu automatiquement
  },

  // --- Outlook ---
  outlook: {
    // Nom exact du calendrier partagé des absences
    calendarName:  "Absences IT",
    // Email du manager à notifier (cc sur les notifications)
    managerEmail:  "jacques.negrato@mairie-lannemezan.fr",
    // Envoyer une copie au demandeur
    notifyAgent:   true
  },

  // --- Groupe d'agents IT (optionnel) ---
  // Si renseigné, les agents sont chargés depuis ce groupe Azure AD
  // Sinon, ils sont saisis manuellement
  agentsGroupId: "",   // ID du groupe Azure AD "Service IT"

  // --- Options de l'application ---
  app: {
    title:       "CongesIT",
    serviceName: "Service Informatique",
    // Nombre max de jours de CP par an
    defaultSoldeCP:  25,
    // Nombre max de jours RTT par an
    defaultSoldeRTT: 10,
    // Envoyer Teams notification (si Teams webhook configuré)
    teamsWebhook: ""
  }
};
