import { PublicClientApplication, IPublicClientApplication } from '@azure/msal-browser';

// MSAL Configuration for Microsoft login
const msalConfig = {
  auth: {
    clientId: process.env.NEXT_PUBLIC_MSAL_CLIENT_ID || '',
    authority: process.env.NEXT_PUBLIC_MSAL_AUTHORITY || 'https://login.microsoftonline.com/common',
    redirectUri: typeof window !== 'undefined' ? window.location.origin : '',
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
};

// Scopes required for calendar access
export const calendarScopes = ['Calendars.ReadBasic', 'offline_access'];

let msalInstance: IPublicClientApplication | null = null;

/**
 * Initialize MSAL instance (client-side only)
 */
export const getMsalInstance = (): IPublicClientApplication => {
  if (typeof window === 'undefined') {
    throw new Error('MSAL can only be initialized on the client side');
  }

  if (!msalInstance) {
    msalInstance = new PublicClientApplication(msalConfig);
  }

  return msalInstance;
};

/**
 * Initialize MSAL (idempotent)
 */
export const initializeMsal = async (): Promise<void> => {
  if (typeof window === 'undefined') return;

  const instance = getMsalInstance();
  try {
    await instance.initialize();
  } catch (error) {
    console.error('Failed to initialize MSAL:', error);
    throw error;
  }
};

/**
 * Get access token for Microsoft Graph
 */
export const getAccessToken = async (): Promise<string> => {
  const instance = getMsalInstance();
  const account = instance.getAllAccounts()[0];

  if (!account) {
    throw new Error('No account found. User must be logged in first.');
  }

  try {
    const response = await instance.acquireTokenSilent({
      scopes: calendarScopes,
      account,
    });
    return response.accessToken;
  } catch (error) {
    // Silent token acquisition failed, try interactive
    console.warn('Silent token acquisition failed, attempting interactive login...');
    const response = await instance.acquireTokenPopup({
      scopes: calendarScopes,
    });
    return response.accessToken;
  }
};

/**
 * Login with Microsoft
 */
export const loginWithMicrosoft = async (): Promise<void> => {
  const instance = getMsalInstance();
  try {
    await instance.loginPopup({
      scopes: calendarScopes,
    });
  } catch (error) {
    console.error('Login failed:', error);
    throw error;
  }
};

/**
 * Logout
 */
export const logout = async (): Promise<void> => {
  const instance = getMsalInstance();
  const account = instance.getAllAccounts()[0];

  try {
    await instance.logoutPopup({
      account,
    });
  } catch (error) {
    console.error('Logout failed:', error);
    throw error;
  }
};

/**
 * Get current authenticated user
 */
export const getCurrentUser = () => {
  if (typeof window === 'undefined') return null;

  try {
    const instance = getMsalInstance();
    const accounts = instance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  } catch (error) {
    console.error('Error getting current user:', error);
    return null;
  }
};
