import { initializeApp } from 'firebase/app';
import { getAuth, signInWithPopup, GoogleAuthProvider, onAuthStateChanged, User, signOut } from 'firebase/auth';
import firebaseConfig from '../firebase-applet-config.json';

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);

const provider = new GoogleAuthProvider();
// Request Google Sheets scopes
provider.addScope('https://www.googleapis.com/auth/spreadsheets');

// Helper to get and validate stored access token
export const getStoredAccessToken = (): string | null => {
  if (typeof window === 'undefined') return null;
  try {
    const token = localStorage.getItem('google_access_token');
    const expiresAtStr = localStorage.getItem('google_access_token_expires_at');
    if (!token) return null;

    // If expiresAt is missing (from older sessions) or passed, treat as expired
    if (!expiresAtStr) {
      clearAuthToken();
      return null;
    }
    const expiresAt = parseInt(expiresAtStr, 10);
    if (isNaN(expiresAt) || Date.now() >= expiresAt) {
      clearAuthToken();
      return null;
    }
    return token;
  } catch {
    return null;
  }
};

export const clearAuthToken = () => {
  cachedAccessToken = null;
  if (typeof window !== 'undefined') {
    try {
      localStorage.removeItem('google_access_token');
      localStorage.removeItem('google_access_token_expires_at');
    } catch {
      // ignore
    }
  }
};

// Flag to indicate if we are in the middle of a sign-in flow.
let isSigningIn = false;
// Cache the access token in memory and local storage.
let cachedAccessToken: string | null = getStoredAccessToken();

// Initialize auth state listener.
export const initAuth = (
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthFailure?: () => void
) => {
  return onAuthStateChanged(auth, async (user: User | null) => {
    if (user) {
      const validToken = getStoredAccessToken();
      if (validToken) {
        cachedAccessToken = validToken;
        if (onAuthSuccess) onAuthSuccess(user, validToken);
      } else {
        // User session exists in Firebase, but Google OAuth access token has expired
        cachedAccessToken = null;
        if (onAuthFailure) onAuthFailure();
      }
    } else {
      clearAuthToken();
      if (onAuthFailure) onAuthFailure();
    }
  });
};

// Sign in via Google popup
export const googleSignIn = async (): Promise<{ user: User; accessToken: string } | null> => {
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Failed to get access token from Firebase Auth');
    }

    cachedAccessToken = credential.accessToken;
    if (typeof window !== 'undefined') {
      const expiresIn = 3500 * 1000; // ~1 hour validity window with buffer
      const expiresAt = Date.now() + expiresIn;
      localStorage.setItem('google_access_token', cachedAccessToken);
      localStorage.setItem('google_access_token_expires_at', expiresAt.toString());
    }
    return { user: result.user, accessToken: cachedAccessToken };
  } catch (error: any) {
    console.error('Sign in error:', error);
    throw error;
  } finally {
    isSigningIn = false;
  }
};

export const getAccessToken = async (): Promise<string | null> => {
  return getStoredAccessToken();
};

export const logout = async () => {
  try {
    await signOut(auth);
  } catch (e) {
    console.warn('SignOut error:', e);
  }
  clearAuthToken();
};
