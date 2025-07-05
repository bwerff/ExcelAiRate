import { supabaseClient } from './supabase'

interface AIResponse {
  success: boolean
  result: any
  cached?: boolean
  usage?: {
    tokens: number
    response_time_ms: number
  }
  error?: string
}

class AIService {
  async analyze(prompt: string, data: string): Promise<any> {
    const response = await this.makeRequest('analyze', prompt, data)
    return response.result
  }

  async generate(prompt: string): Promise<string> {
    const response = await this.makeRequest('generate', prompt)
    return response.result
  }

  async explain(prompt: string, data: string): Promise<string> {
    const response = await this.makeRequest('explain', prompt, data)
    return response.result
  }

  private async makeRequest(type: string, prompt: string, data?: string): Promise<AIResponse> {
    const { data: { session } } = await supabaseClient.auth.getSession()
    
    if (!session) {
      throw new Error('Not authenticated')
    }

    const response = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/ai`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${session.access_token}`
      },
      body: JSON.stringify({
        type,
        prompt,
        data
      })
    })

    const result = await response.json()

    if (!response.ok) {
      if (result.code === 'USAGE_LIMIT_EXCEEDED') {
        throw new Error('You have reached your usage limit. Please upgrade your plan.')
      }
      throw new Error(result.error || 'AI request failed')
    }

    return result
  }
}

export const aiService = new AIService()