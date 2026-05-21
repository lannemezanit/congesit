// ============================================================
// config.js — CongesIT
// Personnalisez ce fichier avec vos identifiants Microsoft 365
// ============================================================

const CONFIG = {

  // --- Azure AD (récupérés sur portal.azure.com) ---
  clientId:  "VOTRE-APP-ID",         // ID d'application (client)
  tenantId:  "VOTRE-TENANT-ID",      // ID de l'annuaire (locataire)

  // --- URI de redirection (doit correspondre à Azure AD) ---
  redirectUri: window.location.origin + window.location.pathname,

  // --- SharePoint ---
  sharepoint: {
    // URL de votre site SharePoint IT
    siteUrl:    "https://VOTRETENANT.sharepoint.com/sites/IT",
    // Nom du fichier Excel de suivi (dans la bibliothèque Documents)
    fileName:   "Conges_2025.xlsx",
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
    managerEmail:  "manager.it@votreentreprise.fr",
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
