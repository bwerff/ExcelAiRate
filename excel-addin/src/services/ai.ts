import { supabaseClient } from './supabase'
import type { AIServiceResponse, AIAnalyzeResponse, AIGenerateResponse, AIExplainResponse } from '../../../shared/types'

class AIService {
  async analyze(prompt: string, data: string): Promise<AIAnalyzeResponse> {
    const response = await this.makeRequest('analyze', prompt, data)
    return response.result as AIAnalyzeResponse
  }

  async generate(prompt: string): Promise<AIGenerateResponse> {
    const response = await this.makeRequest('generate', prompt)
    return response.result as AIGenerateResponse
  }

  async explain(prompt: string, data: string): Promise<AIExplainResponse> {
    const response = await this.makeRequest('explain', prompt, data)
    return response.result as AIExplainResponse
  }

  private async makeRequest(type: string, prompt: string, data?: string): Promise<AIServiceResponse> {
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