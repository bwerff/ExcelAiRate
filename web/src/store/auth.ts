import { create } from 'zustand'
import { supabaseClient } from '../../lib/supabase'

interface User {
  id: string
  email: string
  plan: 'free' | 'pro' | 'team'
  queries_used: number
  queries_limit: number
}

interface AuthStore {
  user: User | null
  loading: boolean
  error: string | null
  
  signIn: (email: string) => Promise<void>
  signOut: () => Promise<void>
  checkSession: () => Promise<void>
  updateUser: (updates: Partial<User>) => void
}

export const useAuthStore = create<AuthStore>((set) => ({
  user: null,
  loading: false,
  error: null,

  signIn: async (email: string) => {
    set({ loading: true, error: null })
    try {
      const { error } = await supabaseClient.auth.signInWithOtp({
        email,
        options: {
          emailRedirectTo: `${window.location.origin}/auth/callback`
        }
      })
      
      if (error) throw error
      
      set({ loading: false })
    } catch (error: any) {
      set({ error: error.message, loading: false })
      throw error
    }
  },

  signOut: async () => {
    set({ loading: true })
    try {
      const { error } = await supabaseClient.auth.signOut()
      if (error) throw error
      
      set({ user: null, loading: false })
    } catch (error: any) {
      set({ error: error.message, loading: false })
      throw error
    }
  },

  checkSession: async () => {
    set({ loading: true })
    try {
      const { data: { session } } = await supabaseClient.auth.getSession()
      
      if (session) {
        const { data: profile } = await supabaseClient
          .from('profiles')
          .select('*')
          .eq('id', session.user.id)
          .single()
        
        if (profile) {
          set({ 
            user: {
              id: session.user.id,
              email: session.user.email!,
              ...profile
            },
            loading: false 
          })
        }
      } else {
        set({ user: null, loading: false })
      }
    } catch (error: any) {
      set({ error: error.message, loading: false })
    }
  },

  updateUser: (updates: Partial<User>) => {
    set((state) => ({
      user: state.user ? { ...state.user, ...updates } : null
    }))
  }
}))