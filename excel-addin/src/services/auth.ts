import { supabaseClient } from './supabase'

class AuthService {
  async signIn(email: string): Promise<void> {
    const { error } = await supabaseClient.auth.signInWithOtp({
      email,
      options: {
        emailRedirectTo: `${window.location.origin}/auth/callback`
      }
    })

    if (error) {
      throw error
    }
  }

  async signOut(): Promise<void> {
    const { error } = await supabaseClient.auth.signOut()
    if (error) {
      throw error
    }
  }

  getSession() {
    return supabaseClient.auth.getSession()
  }

  onAuthStateChange(callback: (event: string, session: any) => void) {
    return supabaseClient.auth.onAuthStateChange(callback)
  }
}

export const authService = new AuthService()