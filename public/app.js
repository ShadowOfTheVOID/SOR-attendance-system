// Shared sign-in and API helpers for every page.
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import {
  getAuth, GoogleAuthProvider, onAuthStateChanged, signInWithPopup, signOut
} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';

const cfg = window.APP_CONFIG;
const auth = getAuth(initializeApp(cfg.firebase));

export const $ = (id) => document.getElementById(id);

export function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

export function setMsg(el, kind, text) {
  el.className = 'msg ' + kind;
  el.textContent = text;
}

/** Calls the Apps Script API as the signed-in user. Throws Error(message) on failure. */
export async function api(action, params = {}) {
  const user = auth.currentUser;
  if (!user) throw new Error('Please sign in.');
  const idToken = await user.getIdToken();
  // text/plain keeps this a "simple" request, which Apps Script accepts cross-origin.
  const res = await fetch(cfg.scriptUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, idToken, ...params })
  });
  if (!res.ok) throw new Error('Server error (' + res.status + '). Try again.');
  const body = await res.json();
  if (!body.ok) throw new Error(body.error);
  return body.data;
}

/**
 * Renders a sign-in button into #auth and calls onSignedIn(user) once signed in.
 * Shows the signed-in email with a sign-out link.
 */
export function requireSignIn(onSignedIn) {
  const box = $('auth');
  onAuthStateChanged(auth, (user) => {
    if (!user) {
      box.innerHTML = '<p>Sign in with the Google account on the roster.</p>' +
        '<button class="btn" id="signin">Sign in with Google</button>';
      $('signin').onclick = async () => {
        try {
          await signInWithPopup(auth, new GoogleAuthProvider());
        } catch (e) {
          if (e.code !== 'auth/popup-closed-by-user') alert(e.message);
        }
      };
      document.querySelectorAll('.signed-in').forEach((el) => el.classList.add('hidden'));
      return;
    }
    box.innerHTML = '<p class="who">Signed in as <b>' + esc(user.email) +
      '</b> · <a href="#" id="signout">Sign out</a></p>';
    $('signout').onclick = (e) => { e.preventDefault(); signOut(auth); };
    document.querySelectorAll('.signed-in').forEach((el) => el.classList.remove('hidden'));
    onSignedIn(user);
  });
}
