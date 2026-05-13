import { createContext, useContext, useState, useCallback, useEffect, type ReactNode } from 'react';
import type { User } from './api';
import { firebaseGetMe, ensureDefaultAdmin } from './firebaseAuth';

const TOKEN_KEY = 'recordbook_token';
const USER_KEY = 'recordbook_user';

interface AuthContextType {
  user: User | null;
  token: string | null;
  isLoading: boolean;
  login: (token: string, user: User) => void;
  logout: () => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

/**
 * Session-based auth: uses sessionStorage so closing the tab/browser logs the user out.
 * Also stores in localStorage for tab-duplication support, but clears on beforeunload.
 */
function loadSession(): { token: string | null; user: User | null } {
  // Try sessionStorage first (primary)
  let savedToken = sessionStorage.getItem(TOKEN_KEY);
  let savedUser = sessionStorage.getItem(USER_KEY);

  // Fallback to localStorage for page-refresh during the same session
  if (!savedToken) {
    savedToken = localStorage.getItem(TOKEN_KEY);
    savedUser = localStorage.getItem(USER_KEY);
    // If found in localStorage, copy to sessionStorage
    if (savedToken && savedUser) {
      sessionStorage.setItem(TOKEN_KEY, savedToken);
      sessionStorage.setItem(USER_KEY, savedUser);
    }
  }

  if (savedToken && savedUser) {
    try {
      const parsed = JSON.parse(savedUser);
      if (parsed && (parsed.email || parsed.name)) {
        return { token: savedToken, user: parsed };
      }
    } catch {
      // corrupted
    }
  }

  // Clean up
  sessionStorage.removeItem(TOKEN_KEY);
  sessionStorage.removeItem(USER_KEY);
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
  return { token: null, user: null };
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const [session] = useState(loadSession);
  const [token, setToken] = useState<string | null>(session.token);
  const [user, setUser] = useState<User | null>(session.user);
  const [isLoading, setIsLoading] = useState(!!session.token);

  const login = useCallback((newToken: string, newUser: User) => {
    // Store in both sessionStorage (primary) and localStorage (refresh support)
    sessionStorage.setItem(TOKEN_KEY, newToken);
    sessionStorage.setItem(USER_KEY, JSON.stringify(newUser));
    localStorage.setItem(TOKEN_KEY, newToken);
    localStorage.setItem(USER_KEY, JSON.stringify(newUser));
    setToken(newToken);
    setUser(newUser);
  }, []);

  const logout = useCallback(() => {
    sessionStorage.removeItem(TOKEN_KEY);
    sessionStorage.removeItem(USER_KEY);
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(USER_KEY);
    setToken(null);
    setUser(null);
  }, []);

  // Auto-logout when browser/tab is closed (beforeunload clears localStorage)
  useEffect(() => {
    const handleBeforeUnload = () => {
      // Clear localStorage so next fresh open requires re-login.
      // sessionStorage is automatically cleared when the tab closes.
      localStorage.removeItem(TOKEN_KEY);
      localStorage.removeItem(USER_KEY);
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, []);

  // Validate session with Firebase on mount — refreshes permissions and catches invalid tokens
  useEffect(() => {
    // Always ensure default admin exists
    ensureDefaultAdmin().catch(() => {});

    if (!token) {
      setIsLoading(false);
      return;
    }

    let cancelled = false;

    firebaseGetMe(token)
      .then((data) => {
        if (cancelled) return;
        const serverUser = data.user;
        if (serverUser) {
          if (serverUser.status && serverUser.status !== 'active') {
            logout();
            return;
          }
          sessionStorage.setItem(USER_KEY, JSON.stringify(serverUser));
          localStorage.setItem(USER_KEY, JSON.stringify(serverUser));
          setUser(serverUser as User);
        }
      })
      .catch(() => {
        if (cancelled) return;
        // Token is invalid — force logout
        logout();
      })
      .finally(() => {
        if (!cancelled) setIsLoading(false);
      });

    return () => { cancelled = true; };
  }, [token, logout]);

  return (
    <AuthContext.Provider value={{ user, token, isLoading, login, logout }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (context === undefined) throw new Error('useAuth must be used within an AuthProvider');
  return context;
}
