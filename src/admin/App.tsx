import React, { useEffect, useState } from 'react';
import FolderBrowser from './components/FolderBrowser';
import MailFolderBrowser from './components/MailFolderBrowser';
import DenyListManager from './components/DenyListManager';
import UserSettings from './components/UserSettings';
import ConnectedServices from './components/ConnectedServices';
import AllowedSites from './components/AllowedSites';
import CollapsibleSection from './components/CollapsibleSection';
import UserManagement from './components/UserManagement';
import EmailOutputPolicy from './components/EmailOutputPolicy';
import OutboundPolicy from './components/OutboundPolicy';
import AuditLog from './components/AuditLog';
import { readRuntimeConfig } from './runtimeConfig';

// Instance display name is injected by the Azure Function serving index.html.
const { instanceName } = readRuntimeConfig();

// Guards against an infinite login bounce if the OAuth round-trip fails to set
// the mcp_session cookie (e.g. a Secure cookie over a plain-http origin).
const REDIRECT_GUARD_KEY = 'm365_auth_redirected';

// /api/manage/* needs a console session: a short-lived, browser-only cookie
// minted at sign-in and renewed by /api/auth/me. Calling /me on this interval
// while the tab is visible keeps it alive; once it lapses (idle tab, or the
// 8-hour cap) the next check sends the user back through sign-in.
const CONSOLE_KEEPALIVE_MS = 5 * 60 * 1000;

type View = 'admin' | 'settings';
type AuthState = 'checking' | 'authed' | 'error';

interface MeResponse {
  authenticated: boolean;
  userId: string;
  displayName?: string;
  email?: string;
  isGlobalAdmin?: boolean;
  consoleSession?: boolean;
}

function startLogin() {
  window.location.href = '/api/auth/login';
}

// Logout is a POST so another page cannot sign the user out with a link.
// A form submission carries the Origin header the server checks.
function signOut() {
  const form = document.createElement('form');
  form.method = 'POST';
  form.action = '/api/auth/logout';
  document.body.appendChild(form);
  form.submit();
}

function AppContent() {
  const [authState, setAuthState] = useState<AuthState>('checking');
  const [view, setView] = useState<View>('settings');
  const [userId, setUserId] = useState<string>('');
  const [accountName, setAccountName] = useState<string>('');
  const [isGlobalAdmin, setIsGlobalAdmin] = useState<boolean>(false);
  const [enabledServices, setEnabledServices] = useState<string[]>(['mail', 'sharepoint']);
  const [allowedSites, setAllowedSites] = useState<Array<{ id: string; name: string }> | undefined>(undefined);
  const [adminRefreshKey, setAdminRefreshKey] = useState(0);

  // Establish the server session. The mcp_session cookie and, for
  // /api/manage/*, the mcp_console cookie (both set by the OAuth callback)
  // authorize every /api/* call automatically on this origin, so all we need
  // here is to confirm they exist and learn who we are. No client Graph token,
  // no MSAL: identity + admin status come from the session.
  useEffect(() => {
    let cancelled = false;
    fetch('/api/auth/me', { credentials: 'same-origin' })
      .then(async (r) => {
        // A session without a console session (expired, or from before
        // console sessions existed) needs a fresh sign-in to use the admin API.
        const data = r.ok ? ((await r.json()) as MeResponse) : null;
        if (data?.consoleSession) {
          if (cancelled) return;
          sessionStorage.removeItem(REDIRECT_GUARD_KEY);
          setUserId(data.userId);
          setAccountName(data.displayName ?? data.email ?? '');
          setIsGlobalAdmin(Boolean(data.isGlobalAdmin));
          setAuthState('authed');
          return;
        }
        if (r.status === 401 || data) {
          // No valid session — start the server OAuth flow, unless we already
          // tried once (in which case something is wrong; show an error).
          if (sessionStorage.getItem(REDIRECT_GUARD_KEY)) {
            if (!cancelled) setAuthState('error');
            return;
          }
          sessionStorage.setItem(REDIRECT_GUARD_KEY, '1');
          startLogin();
          return;
        }
        if (!cancelled) setAuthState('error');
      })
      .catch(() => {
        if (!cancelled) setAuthState('error');
      });
    return () => { cancelled = true; };
  }, []);

  // Keep the console session alive while the tab is in use; re-sign-in when it
  // has lapsed. Checked on an interval and whenever the tab becomes visible.
  useEffect(() => {
    if (authState !== 'authed') return;
    const check = () => {
      if (document.visibilityState !== 'visible') return;
      fetch('/api/auth/me', { credentials: 'same-origin' })
        .then(async (r) => {
          const data = r.ok ? ((await r.json()) as MeResponse) : null;
          if (r.status === 401 || (data && !data.consoleSession)) startLogin();
        })
        .catch(() => {});
    };
    const timer = window.setInterval(check, CONSOLE_KEEPALIVE_MS);
    document.addEventListener('visibilitychange', check);
    return () => {
      window.clearInterval(timer);
      document.removeEventListener('visibilitychange', check);
    };
  }, [authState]);

  // Fetch org-wide enabled services and allowed sites (scoped by tenant via the
  // session cookie).
  useEffect(() => {
    if (authState !== 'authed') return;
    fetch('/api/manage/services', { credentials: 'same-origin' })
      .then((r) => (r.ok ? r.json() : Promise.reject(r.status)))
      .then((data: { enabledServices?: string[] }) => {
        if (Array.isArray(data.enabledServices)) setEnabledServices(data.enabledServices);
      })
      .catch(() => {});
    fetch('/api/manage/allowed-sites', { credentials: 'same-origin' })
      .then((r) => (r.ok ? r.json() : Promise.reject(r.status)))
      .then((data: { allowedSites?: Array<{ id: string; name: string }> }) => {
        if (Array.isArray(data.allowedSites)) setAllowedSites(data.allowedSites);
      })
      .catch(() => {});
  }, [authState]);

  if (authState === 'checking') {
    return <div className="loading">Signing in to Microsoft 365…</div>;
  }

  if (authState === 'error') {
    return (
      <div className="loading">
        <p>Sign-in could not be completed.</p>
        <button
          onClick={() => {
            sessionStorage.removeItem(REDIRECT_GUARD_KEY);
            startLogin();
          }}
        >
          Try again
        </button>
      </div>
    );
  }

  return (
    <div className="app">
      <nav className="navbar">
        <h1>{instanceName} — Access Management</h1>
        <div className="nav-links">
          {isGlobalAdmin && (
            <button
              className={view === 'admin' ? 'active' : ''}
              onClick={() => setView('admin')}
            >
              Admin
            </button>
          )}
          <button
            className={view === 'settings' ? 'active' : ''}
            onClick={() => setView('settings')}
          >
            My Settings
          </button>
          <span className="user-info">{accountName}</span>
          <button onClick={signOut}>Sign out</button>
        </div>
      </nav>

      <main className="main-content">
        {view === 'admin' ? (
          <div className="admin-view">
            <CollapsibleSection storageKey="connected-services" title="Connected Services">
              <ConnectedServices
                enabledServices={enabledServices}
                setEnabledServices={setEnabledServices}
                userId={userId}
              />
            </CollapsibleSection>
            <CollapsibleSection storageKey="email-output-policy" title="Email Output Policy">
              <EmailOutputPolicy userId={userId} />
            </CollapsibleSection>
            <CollapsibleSection storageKey="outbound-policy" title="Outbound Policy">
              <OutboundPolicy userId={userId} enabledServices={enabledServices} />
            </CollapsibleSection>
            <CollapsibleSection storageKey="user-management" title="User Management">
              <UserManagement
                userId={userId}
                enabledServices={enabledServices}
              />
            </CollapsibleSection>
            <CollapsibleSection storageKey="allowed-sites" title="Allowed Sites">
              <AllowedSites
                userId={userId}
                isGlobalAdmin={isGlobalAdmin}
              />
            </CollapsibleSection>
            <CollapsibleSection storageKey="global-deny-list" title="Global Deny List" defaultCollapsed>
              <DenyListManager key={adminRefreshKey} scope="global" userId={userId} readOnly={!isGlobalAdmin} addedByName={accountName} bare />
            </CollapsibleSection>
            <CollapsibleSection storageKey="audit-log" title="Audit Log" defaultCollapsed={false}>
              <AuditLog />
            </CollapsibleSection>
            <div className="browsers">
              <FolderBrowser userId={userId} addedByName={accountName} onDenyChange={() => setAdminRefreshKey((k) => k + 1)} />
              <MailFolderBrowser userId={userId} />
            </div>
          </div>
        ) : (
          <UserSettings userId={userId} accountName={accountName} enabledServices={enabledServices} allowedSites={allowedSites} instanceName={instanceName} />
        )}
      </main>
    </div>
  );
}

export default function App() {
  return <AppContent />;
}
