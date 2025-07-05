// Simplified 3-tier pricing configuration
export const PRICING_PLANS = {
  free: {
    name: 'Free',
    price: 0,
    description: 'Perfect for trying out ExcelAIRate',
    features: [
      '10 AI queries per month',
      'Basic Excel analysis',
      'Standard support',
      'GPT-3.5 model'
    ],
    limits: {
      queries: 10,
      gpt4_queries: 0
    },
    cta: 'Get Started',
    popular: false
  },
  pro: {
    name: 'Pro',
    price: 29,
    yearlyPrice: 290, // 2 months free
    description: 'For professionals who need more power',
    features: [
      '500 AI queries per month',
      'Advanced analysis & insights',
      'Content generation',
      'GPT-4 Turbo access',
      'Priority support',
      'Save custom templates'
    ],
    limits: {
      queries: 500,
      gpt4_queries: 100
    },
    cta: 'Start Free Trial',
    popular: true
  },
  team: {
    name: 'Team',
    price: 99,
    yearlyPrice: 990, // 2 months free
    description: 'For teams and heavy users',
    features: [
      '5,000 AI queries per month',
      'Unlimited GPT-4 Turbo',
      'Team collaboration',
      'Advanced templates',
      'Phone support',
      'Usage analytics',
      'API access (coming soon)'
    ],
    limits: {
      queries: 5000,
      gpt4_queries: 5000
    },
    cta: 'Contact Sales',
    popular: false
  }
} as const

// Helper to format prices
export function formatPrice(price: number, interval: 'monthly' | 'yearly' = 'monthly') {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(price)
}

// Get savings for yearly plans
export function getYearlySavings(monthlyPrice: number, yearlyPrice: number) {
  const yearlyCost = monthlyPrice * 12
  const savings = yearlyCost - yearlyPrice
  const percentSaved = Math.round((savings / yearlyCost) * 100)
  return { amount: savings, percent: percentSaved }
}