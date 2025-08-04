import * as msal from '@azure/msal-browser';

// MSAL configuration
const msalConfig = {
    auth: {
        clientId: import.meta.env.VITE_AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID}`,
        redirectUri: import.meta.env.VITE_AZURE_REDIRECT_URI
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
    }
};

// MSAL instance
const msalInstance = new msal.PublicClientApplication(msalConfig);

// Scopes required for accessing SharePoint files
const loginRequest = {
    scopes: import.meta.env.VITE_GRAPH_SCOPES.split(',')
};

// Function to handle login
export async function login() {
    try {
        const authResult = await msalInstance.loginPopup(loginRequest);
        return authResult;
    } catch (error) {
        console.error('Login failed:', error);
        throw error;
    }
}

// Function to get access token
export async function getAccessToken() {
    try {
        const account = msalInstance.getAllAccounts()[0];
        if (!account) {
            throw new Error('No active account! Please sign in first.');
        }

        const silentRequest = {
            scopes: loginRequest.scopes,
            account: account
        };

        const response = await msalInstance.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error) {
        console.error('Failed to get access token:', error);
        throw error;
    }
}

// Function to check if user is authenticated
export function isAuthenticated() {
    const accounts = msalInstance.getAllAccounts();
    return accounts.length > 0;
}

// Function to log out
export function logout() {
    msalInstance.logout();
}
