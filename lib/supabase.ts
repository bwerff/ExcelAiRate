import { createClient } from '@supabase/supabase-js'

// Simplified Supabase client factory
export function createSupabaseClient() {
  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL!
  const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
  
  return createClient(supabaseUrl, supabaseAnonKey)
}

// Export singleton instance for compatibility
export const supabaseClient = createSupabaseClient()

// Auth helpers
export const auth = {
  // Magic link sign in
  async signIn(email: string) {
    const supabase = createSupabaseClient()
    
    const { error } = await supabase.auth.signInWithOtp({
      email,
      options: {
        emailRedirectTo: `${window.location.origin}/auth/callback`,
      }
    })
    
    if (error) throw error
    
    return { success: true, message: 'Check your email for the login link!' }
  },

  // Sign out
  async signOut() {
    const supabase = createSupabaseClient()
    const { error } = await supabase.auth.signOut()
    if (error) throw error
  },

  // Get current user
  async getUser() {
    const supabase = createSupabaseClient()
    const { data: { user } } = await supabase.auth.getUser()
    return user
  },

  // Get session
  async getSession() {
    const supabase = createSupabaseClient()
    const { data: { session } } = await supabase.auth.getSession()
    return session
  }
}

// API helpers
export const api = {
  // Call AI function
  async callAI(type: 'analyze' | 'generate' | 'explain', prompt: string, data?: string) {
    const supabase = createSupabaseClient()
    const session = await auth.getSession()
    
    if (!session) throw new Error('Not authenticated')
    
    const { data: result, error } = await supabase.functions.invoke('ai', {
      body: { type, prompt, data }
    })
    
    if (error) throw error
    return result
  },

  // Create checkout session
  async createCheckout(plan: 'pro' | 'team', interval: 'monthly' | 'yearly' = 'monthly') {
    const supabase = createSupabaseClient()
    const session = await auth.getSession()
    
    if (!session) throw new Error('Not authenticated')
    
    const { data, error } = await supabase.functions.invoke('stripe-checkout', {
      body: { plan, interval }
    })
    
    if (error) throw error
    return data
  },

  // Open billing portal
  async openBillingPortal() {
    const supabase = createSupabaseClient()
    const session = await auth.getSession()
    
    if (!session) throw new Error('Not authenticated')
    
    const { data, error } = await supabase.functions.invoke('stripe-portal')
    
    if (error) throw error
    return data
  },

  // Get user profile with usage
  async getProfile() {
    const supabase = createSupabaseClient()
    const user = await auth.getUser()
    
    if (!user) throw new Error('Not authenticated')
    
    const { data: profile, error } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', user.id)
      .single()
    
    if (error) throw error
    
    // Get usage stats
    const { data: usage } = await supabase
      .rpc('get_usage_stats', { user_id: user.id })
    
    return { ...profile, usage }
  },

  // Get templates
  async getTemplates(category?: string) {
    const supabase = createSupabaseClient()
    const user = await auth.getUser()
    
    let query = supabase
      .from('templates')
      .select('*')
      .or(`is_public.eq.true${user ? `,user_id.eq.${user.id}` : ''}`)
      .order('usage_count', { ascending: false })
    
    if (category) {
      query = query.eq('category', category)
    }
    
    const { data, error } = await query
    
    if (error) throw error
    return data
  },

  // Save template
  async saveTemplate(template: {
    name: string
    description?: string
    prompt_template: string
    category?: string
    is_public?: boolean
  }) {
    const supabase = createSupabaseClient()
    const user = await auth.getUser()
    
    if (!user) throw new Error('Not authenticated')
    
    const { data, error } = await supabase
      .from('templates')
      .insert({
        ...template,
        user_id: user.id
      })
      .select()
      .single()
    
    if (error) throw error
    return data
  }
}

// Types
export interface Profile {
  id: string
  email: string
  full_name?: string
  plan: 'free' | 'pro' | 'team'
  subscription_status: string
  queries_used: number
  queries_limit: number
  usage?: {
    used: number
    limit: number
    plan: string
    period_start: string
    period_end: string
  }
}

export interface Template {
  id: string
  name: string
  description?: string
  prompt_template: string
  category: string
  is_public: boolean
  usage_count: number
}