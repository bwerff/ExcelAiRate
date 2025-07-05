import { createSupabaseClient } from './supabase'
import { auth } from './supabase'

export const stripe = {
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
  }
}

// Re-export for compatibility
export const createCheckout = stripe.createCheckout
export const openBillingPortal = stripe.openBillingPortal
