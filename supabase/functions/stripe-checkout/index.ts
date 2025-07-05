import { serve } from 'https://deno.land/std@0.168.0/http/server.ts'
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import Stripe from 'https://esm.sh/stripe@14.0.0'
import { getCorsHeaders, corsResponse } from '../_shared/cors.ts'

// Simplified pricing - hardcoded for easy maintenance
const PRICES = {
  pro: {
    monthly: Deno.env.get('STRIPE_PRICE_PRO_MONTHLY')!,
    yearly: Deno.env.get('STRIPE_PRICE_PRO_YEARLY')!,
  },
  team: {
    monthly: Deno.env.get('STRIPE_PRICE_TEAM_MONTHLY')!,
    yearly: Deno.env.get('STRIPE_PRICE_TEAM_YEARLY')!,
  }
}

const PLAN_LIMITS = {
  free: { queries: 10 },
  pro: { queries: 500 },
  team: { queries: 5000 }
}

serve(async (req) => {
  const origin = req.headers.get('origin')
  
  if (req.method === 'OPTIONS') {
    return corsResponse(origin)
  }
  
  const corsHeaders = getCorsHeaders(origin)

  try {
    // Initialize services
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const stripeKey = Deno.env.get('STRIPE_SECRET_KEY')!
    
    const supabase = createClient(supabaseUrl, supabaseServiceKey)
    const stripe = new Stripe(stripeKey, { apiVersion: '2023-10-16' })

    // Get user
    const authHeader = req.headers.get('Authorization')!
    const token = authHeader.replace('Bearer ', '')
    const { data: { user }, error: authError } = await supabase.auth.getUser(token)
    
    if (authError || !user) {
      throw new Error('Unauthorized')
    }

    // Parse request
    const { plan, interval = 'monthly' } = await req.json()
    
    if (!['pro', 'team'].includes(plan)) {
      throw new Error('Invalid plan')
    }

    // Get or create Stripe customer
    const { data: profile } = await supabase
      .from('profiles')
      .select('stripe_customer_id')
      .eq('id', user.id)
      .single()

    let customerId = profile?.stripe_customer_id

    if (!customerId) {
      const customer = await stripe.customers.create({
        email: user.email,
        metadata: { supabase_user_id: user.id }
      })
      customerId = customer.id

      // Save customer ID
      await supabase
        .from('profiles')
        .update({ stripe_customer_id: customerId })
        .eq('id', user.id)
    }

    // Create checkout session
    const session = await stripe.checkout.sessions.create({
      customer: customerId,
      payment_method_types: ['card'],
      mode: 'subscription',
      line_items: [{
        price: PRICES[plan][interval],
        quantity: 1,
      }],
      success_url: `${Deno.env.get('NEXT_PUBLIC_APP_URL')}/dashboard?success=true`,
      cancel_url: `${Deno.env.get('NEXT_PUBLIC_APP_URL')}/pricing`,
      metadata: {
        user_id: user.id,
        plan,
        interval
      },
      subscription_data: {
        metadata: {
          user_id: user.id,
          plan
        }
      }
    })

    return new Response(
      JSON.stringify({ 
        url: session.url,
        session_id: session.id 
      }),
      { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )

  } catch (error) {
    console.error('Checkout error:', error)
    return new Response(
      JSON.stringify({ error: error.message }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  }
})