// ============================================================
// auth.js — Authentification MSAL (Microsoft 365)
// ============================================================

const Auth = (() => {

  let msalInstance = null;
  let account = null;

  const SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Calendars.ReadWrite",
    "Calendars.Read.Shared",
    "Files.ReadWrite",
    "Sites.ReadWrite.All",
    "Mail.Send"
  ];

  function getMsal() {
    if (!msalInstance) {
      msalInstance = new msal.PublicClientApplication({
        auth: {
          clientId:    CONFIG.clientId,
          authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
          redirectUri: CONFIG.redirectUri
        },
        cache: {
          cacheLocation: "sessionStorage",
          storeAuthStateInCookie: false
        },
        system: {
          loggerOptions: {
            logLevel: msal.LogLevel.Warning,
            loggerCallback: (level, msg) => { if (level <= msal.LogLevel.Warning) console.warn('[MSAL]', msg); }
          }
        }
      });
    }
    return msalInstance;
  }

  async function login() {
    const client = getMsal();
    await client.initialize();
    const res = await client.loginPopup({ scopes: SCOPES });
    account = res.account;
    return { name: account.name, username: account.username, id: account.localAccountId };
  }

  async function silentLogin() {
    const client = getMsal();
    await client.initialize();
    const accounts = client.getAllAccounts();
    if (!accounts.length) return null;
    account = accounts[0];
    // Tentative de refresh silencieux du token
    await client.acquireTokenSilent({ scopes: SCOPES, account });
    return { name: account.name, username: account.username, id: account.localAccountId };
  }

  async function getToken() {
    const client = getMsal();
    if (!account) {
      const accounts = client.getAllAccounts();
      if (!accounts.length) throw new Error('Non connecté');
      account = accounts[0];
    }
    try {
      const res = await client.acquireTokenSilent({ scopes: SCOPES, account });
      return res.accessToken;
    } catch(e) {
      // Si le refresh silencieux échoue, on demande une interaction
      const res = await client.acquireTokenPopup({ scopes: SCOPES, account });
      return res.accessToken;
    }
  }

  async function logout() {
    const client = getMsal();
    await client.logoutPopup({ account });
    account = null;
  }

  function isAuthenticated() {
    try {
      const client = getMsal();
      const accounts = client.getAllAccounts();
      return accounts.length > 0;
    } catch(e) {
      return false;
    }
  }

  return { login, silentLogin, logout, getToken, isAuthenticated };
})();
