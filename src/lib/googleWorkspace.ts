import { getApps, initializeApp } from 'firebase/app';
import {
  getAuth,
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  type User,
} from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';

export const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.file',
] as const;

const app = getApps().length > 0 ? getApps()[0] : initializeApp(firebaseConfig);
const auth = getAuth(app);

const provider = new GoogleAuthProvider();
for (const scope of SCOPES) {
  provider.addScope(scope);
}
provider.setCustomParameters({ prompt: 'select_account' });

let isSigningIn = false;
// Keep Google OAuth access token strictly in JS memory — never localStorage or sessionStorage
let cachedAccessToken: string | null = null;
let cachedGoogleUser: User | null = null;

type AuthListener = (state: { user: User | null; accessToken: string | null }) => void;
const authListeners = new Set<AuthListener>();

function emitAuthState(): void {
  const snapshot = { user: cachedGoogleUser, accessToken: cachedAccessToken };
  for (const listener of authListeners) {
    listener(snapshot);
  }
}

export function subscribeGoogleAuth(listener: AuthListener): () => void {
  authListeners.add(listener);
  listener({ user: cachedGoogleUser, accessToken: cachedAccessToken });
  return () => {
    authListeners.delete(listener);
  };
}

export function initGoogleAuth(
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthFailure?: () => void,
): () => void {
  return onAuthStateChanged(auth, (user: User | null) => {
    cachedGoogleUser = user;
    if (user) {
      if (cachedAccessToken) {
        onAuthSuccess?.(user, cachedAccessToken);
      } else if (!isSigningIn) {
        cachedAccessToken = null;
        onAuthFailure?.();
      }
    } else {
      cachedAccessToken = null;
      onAuthFailure?.();
    }
    emitAuthState();
  });
}

export async function googleSignIn(): Promise<{ user: User; accessToken: string }> {
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Could not obtain Google Sheets access permission. Please try signing in again.');
    }
    cachedAccessToken = credential.accessToken;
    cachedGoogleUser = result.user;
    emitAuthState();
    return { user: result.user, accessToken: cachedAccessToken };
  } finally {
    isSigningIn = false;
  }
}

export function getGoogleAccessToken(): string | null {
  return cachedAccessToken;
}

export function getGoogleUser(): User | null {
  return cachedGoogleUser;
}

export function clearGoogleAccessToken(): void {
  cachedAccessToken = null;
  emitAuthState();
}

export async function googleSignOut(): Promise<void> {
  await auth.signOut().catch(() => {});
  cachedAccessToken = null;
  cachedGoogleUser = null;
  emitAuthState();
}
