/**
 * Session Manager for Excel Add-in
 * Handles token storage and auto-refresh using Office.context.document.settings
 */

import { createClient, SupabaseClient, Session } from '@supabase/supabase-js';

/* global Office */

export class SessionManager {
  private static instance: SessionManager;
  private supabase: SupabaseClient;
  private refreshTimer: number | null = null;
  private onSessionChange: ((session: Session | null) => void) | null = null;

  private constructor(supabaseUrl: string, supabaseAnonKey: string) {
    this.supabase = createClient(supabaseUrl, supabaseAnonKey);
  }

  static getInstance(supabaseUrl: string, supabaseAnonKey: string): SessionManager {
    if (!SessionManager.instance) {
      SessionManager.instance = new SessionManager(supabaseUrl, supabaseAnonKey);
    }
    return SessionManager.instance;
  }

  /**
   * Initialize session manager and set up auto-refresh
   */
  async initialize(onSessionChange?: (session: Session | null) => void): Promise<Session | null> {
    if (onSessionChange) {
      this.onSessionChange = onSessionChange;
    }

    // Try to restore session from Office settings
    const storedSession = await this.getStoredSession();
    if (storedSession) {
      // Validate and refresh if needed
      const validSession = await this.validateAndRefreshSession(storedSession);
      if (validSession) {
        this.setupAutoRefresh(validSession);
        return validSession;
      }
    }

    // Check current Supabase session
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      await this.storeSession(session);
      this.setupAutoRefresh(session);
    }

    // Listen for auth changes
    this.supabase.auth.onAuthStateChange(async (event, session) => {
      if (session) {
        await this.storeSession(session);
        this.setupAutoRefresh(session);
      } else {
        await this.clearStoredSession();
        this.clearAutoRefresh();
      }
      
      if (this.onSessionChange) {
        this.onSessionChange(session);
      }
    });

    return session;
  }

  /**
   * Store session in Office document settings
   */
  private async storeSession(session: Session): Promise<void> {
    return new Promise((resolve) => {
      Office.context.document.settings.set('supabase_session', JSON.stringify({
        access_token: session.access_token,
        refresh_token: session.refresh_token,
        expires_at: session.expires_at,
        expires_in: session.expires_in,
        user: session.user
      }));
      
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Session stored successfully');
        } else {
          console.error('Failed to store session:', result.error);
        }
        resolve();
      });
    });
  }

  /**
   * Get stored session from Office document settings
   */
  private async getStoredSession(): Promise<Session | null> {
    return new Promise((resolve) => {
      const storedData = Office.context.document.settings.get('supabase_session');
      if (!storedData) {
        resolve(null);
        return;
      }

      try {
        const sessionData = JSON.parse(storedData);
        resolve(sessionData as Session);
      } catch (error) {
        console.error('Failed to parse stored session:', error);
        resolve(null);
      }
    });
  }

  /**
   * Clear stored session
   */
  private async clearStoredSession(): Promise<void> {
    return new Promise((resolve) => {
      Office.context.document.settings.remove('supabase_session');
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Session cleared successfully');
        } else {
          console.error('Failed to clear session:', result.error);
        }
        resolve();
      });
    });
  }

  /**
   * Validate session and refresh if needed
   */
  private async validateAndRefreshSession(session: Session): Promise<Session | null> {
    try {
      // Check if session is expired or about to expire (within 5 minutes)
      const expiresAt = session.expires_at || 0;
      const now = Math.floor(Date.now() / 1000);
      const timeUntilExpiry = expiresAt - now;

      if (timeUntilExpiry < 300) { // Less than 5 minutes
        // Refresh the session
        const { data, error } = await this.supabase.auth.refreshSession({
          refresh_token: session.refresh_token
        });

        if (error) {
          console.error('Failed to refresh session:', error);
          return null;
        }

        return data.session;
      }

      // Session is still valid
      await this.supabase.auth.setSession({
        access_token: session.access_token,
        refresh_token: session.refresh_token
      });

      return session;
    } catch (error) {
      console.error('Session validation error:', error);
      return null;
    }
  }

  /**
   * Set up auto-refresh timer
   */
  private setupAutoRefresh(session: Session): void {
    this.clearAutoRefresh();

    // Calculate when to refresh (5 minutes before expiry)
    const expiresAt = session.expires_at || 0;
    const now = Math.floor(Date.now() / 1000);
    const timeUntilRefresh = (expiresAt - now - 300) * 1000; // Convert to milliseconds

    if (timeUntilRefresh > 0) {
      this.refreshTimer = window.setTimeout(async () => {
        const refreshedSession = await this.refreshSession();
        if (refreshedSession) {
          this.setupAutoRefresh(refreshedSession);
        }
      }, timeUntilRefresh);
    }
  }

  /**
   * Clear auto-refresh timer
   */
  private clearAutoRefresh(): void {
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }
  }

  /**
   * Manually refresh the session
   */
  async refreshSession(): Promise<Session | null> {
    try {
      const { data, error } = await this.supabase.auth.refreshSession();
      
      if (error) {
        console.error('Failed to refresh session:', error);
        return null;
      }

      if (data.session) {
        await this.storeSession(data.session);
        return data.session;
      }

      return null;
    } catch (error) {
      console.error('Session refresh error:', error);
      return null;
    }
  }

  /**
   * Get current session
   */
  async getSession(): Promise<Session | null> {
    const { data: { session } } = await this.supabase.auth.getSession();
    return session;
  }

  /**
   * Sign out and clear session
   */
  async signOut(): Promise<void> {
    await this.supabase.auth.signOut();
    await this.clearStoredSession();
    this.clearAutoRefresh();
  }

  /**
   * Get Supabase client instance
   */
  getSupabaseClient(): SupabaseClient {
    return this.supabase;
  }
}

// Export singleton getter
export function getSessionManager(supabaseUrl: string, supabaseAnonKey: string): SessionManager {
  return SessionManager.getInstance(supabaseUrl, supabaseAnonKey);
}