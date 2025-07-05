// Shared type definitions for ExcelAIRate

// Database types
export interface Profile {
  id: string
  email: string
  name?: string
  plan: 'free' | 'pro' | 'team'
  subscription_status: 'active' | 'trialing' | 'past_due' | 'canceled' | 'incomplete' | null
  subscription_id?: string
  customer_id?: string
  subscription_current_period_end?: string
  queries_used: number
  queries_limit: number
  created_at: string
  updated_at: string
}

export interface UsageLog {
  id: string
  user_id: string
  action_type: 'analyze' | 'generate' | 'explain'
  prompt: string
  response: Record<string, unknown>
  model: string
  tokens_used: number
  response_time: number
  cached: boolean
  created_at: string
}

export interface Template {
  id: string
  name: string
  description?: string
  prompt: string
  category: string
  is_public: boolean
  created_by?: string
  created_at: string
  updated_at: string
}

// AI types
export interface AIRequest {
  type: 'analyze' | 'generate' | 'explain'
  prompt: string
  data?: string
  model?: string
  temperature?: number
}

export interface AIAnalyzeResponse {
  summary: string
  insights?: string[]
  recommendations?: string[]
  key_metrics?: Record<string, number | string>
}

export interface AIGenerateResponse {
  content: string
  format?: string
}

export interface AIExplainResponse {
  explanation: string
  examples?: string[]
  tips?: string[]
}

export type AIResponse = AIAnalyzeResponse | AIGenerateResponse | AIExplainResponse

export interface AIServiceResponse {
  success: true
  result: AIResponse
  usage?: {
    prompt_tokens: number
    completion_tokens: number
    total_tokens: number
  }
  cached?: boolean
}

export interface AIServiceError {
  success: false
  error: string
  code?: string
}

export type AIServiceResult = AIServiceResponse | AIServiceError

// Supabase types
export interface SupabaseClient {
  auth: {
    signInWithOtp: (credentials: { email: string }) => Promise<{ data: unknown; error: Error | null }>
    signOut: () => Promise<{ error: Error | null }>
    getSession: () => Promise<{ data: { session: Session | null }; error: Error | null }>
    onAuthStateChange: (callback: (event: string, session: Session | null) => void) => { data: { subscription: { unsubscribe: () => void } } }
  }
  from: (table: string) => {
    select: (columns?: string) => any // This needs to be more specific based on usage
    insert: (data: any) => any
    update: (data: any) => any
    delete: () => any
  }
  functions: {
    invoke: (name: string, options?: { body?: unknown }) => Promise<{ data: unknown; error: Error | null }>
  }
}

export interface Session {
  user: {
    id: string
    email?: string
    user_metadata?: Record<string, unknown>
  }
  access_token: string
  refresh_token: string
  expires_at?: number
}

// Error types
export interface AppError extends Error {
  code?: string
  statusCode?: number
  details?: unknown
}

// Stripe types
export interface StripeSubscription {
  id: string
  status: string
  metadata: {
    user_id?: string
    plan?: string
  }
  items: {
    data: Array<{
      price: {
        id: string
      }
    }>
  }
}

export interface StripeEvent {
  id: string
  type: string
  data: {
    object: StripeSubscription | unknown
  }
}

// Component prop types
export interface AuthResponse {
  data?: unknown
  error?: Error | null
}