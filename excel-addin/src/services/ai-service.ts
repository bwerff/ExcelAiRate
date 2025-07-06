/**
 * AI Service Layer - Core AI operations with OpenAI integration
 */

import { createClient, SupabaseClient } from '@supabase/supabase-js';
import { APIError, NetworkError, AuthenticationError, getErrorMessage } from '../types/errors';

// Types
export interface AIResponse {
  content: string;
  usage?: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
  cached?: boolean;
  model?: string;
}

export interface AIContext {
  conversationHistory: Array<{ role: string; content: string }>;
  workbookContext?: string;
  userPreferences?: {
    defaultModel?: string;
    analysisDepth?: 'shallow' | 'medium' | 'deep';
    outputFormat?: 'concise' | 'detailed' | 'technical';
  };
}

// Configuration
const CONFIG = {
  SUPABASE_URL: import.meta.env.VITE_SUPABASE_URL || '',
  SUPABASE_ANON_KEY: import.meta.env.VITE_SUPABASE_ANON_KEY || '',
  MODEL_PRIMARY: 'gpt-4o-mini',
  MODEL_FALLBACK: 'gpt-3.5-turbo',
  MAX_TOKENS: 500,
  TEMPERATURE: 0.3,
  CACHE_DURATION: 300000, // 5 minutes
  MAX_RETRIES: 3,
  RATE_LIMIT_DELAY: 1000,
  MAX_CACHE_SIZE: 50, // Maximum cache entries
};

// Initialize Supabase client
const supabase: SupabaseClient = createClient(CONFIG.SUPABASE_URL, CONFIG.SUPABASE_ANON_KEY);

// Cache implementation
class ResponseCache {
  private cache: Map<string, { response: AIResponse; timestamp: number }> = new Map();

  generateKey(prompt: string, model: string, params: Record<string, unknown> = {}): string {
    return btoa(JSON.stringify({ prompt, model, ...params }));
  }

  get(key: string): AIResponse | null {
    const cached = this.cache.get(key);
    if (cached && Date.now() - cached.timestamp < CONFIG.CACHE_DURATION) {
      return { ...cached.response, cached: true };
    }
    this.cache.delete(key);
    return null;
  }

  set(key: string, response: AIResponse): void {
    this.cache.set(key, { response, timestamp: Date.now() });
    
    // Clean old entries if cache is too large
    if (this.cache.size > CONFIG.MAX_CACHE_SIZE) {
      // Remove oldest entries (first 10)
      const keysToDelete = Array.from(this.cache.keys()).slice(0, 10);
      keysToDelete.forEach(key => this.cache.delete(key));
    }
  }

  clear(): void {
    this.cache.clear();
  }
}

// Shared state management
export class SharedState {
  private static instance: SharedState;
  
  conversationHistory: Array<{ role: string; content: string }> = [];
  statistics = {
    totalCalls: 0,
    tokensUsed: 0,
    costEstimate: 0,
  };
  currentWorkbook: string | null = null;
  userPreferences = {
    defaultModel: CONFIG.MODEL_PRIMARY,
    analysisDepth: 'medium' as const,
    outputFormat: 'concise' as const,
  };
  lastError: string | null = null;

  static getInstance(): SharedState {
    if (!SharedState.instance) {
      SharedState.instance = new SharedState();
    }
    return SharedState.instance;
  }

  addToHistory(role: string, content: string): void {
    this.conversationHistory.push({ role, content });
    // Keep only last 20 messages
    if (this.conversationHistory.length > 20) {
      this.conversationHistory = this.conversationHistory.slice(-20);
    }
  }

  updateStatistics(tokens: number, model: string): void {
    this.statistics.totalCalls++;
    this.statistics.tokensUsed += tokens;
    this.statistics.costEstimate += this.estimateCost(tokens, model);
  }

  private estimateCost(tokens: number, model: string): number {
    const rates: Record<string, number> = {
      'gpt-4o-mini': 0.00015 / 1000,
      'gpt-3.5-turbo': 0.0005 / 1000,
    };
    return tokens * (rates[model] || rates['gpt-4o-mini']);
  }
}

// Main AI Service
export class AIService {
  private static instance: AIService;
  private cache = new ResponseCache();
  private sharedState = SharedState.getInstance();
  private lastApiCall = 0;

  static getInstance(): AIService {
    if (!AIService.instance) {
      AIService.instance = new AIService();
    }
    return AIService.instance;
  }

  /**
   * Rate limiting for API calls
   */
  private async rateLimitedCall<T>(fn: () => Promise<T>): Promise<T> {
    const now = Date.now();
    const timeSinceLastCall = now - this.lastApiCall;
    
    if (timeSinceLastCall < CONFIG.RATE_LIMIT_DELAY) {
      await new Promise(resolve => 
        setTimeout(resolve, CONFIG.RATE_LIMIT_DELAY - timeSinceLastCall)
      );
    }
    
    this.lastApiCall = Date.now();
    return await fn();
  }

  /**
   * Main AI call method using Supabase Edge Function
   */
  async callAI(
    prompt: string,
    options: {
      model?: string;
      maxTokens?: number;
      temperature?: number;
      systemPrompt?: string;
      type?: 'analyze' | 'generate' | 'explain';
      includeHistory?: boolean;
    } = {}
  ): Promise<AIResponse> {
    const {
      model = CONFIG.MODEL_PRIMARY,
      maxTokens = CONFIG.MAX_TOKENS,
      temperature = CONFIG.TEMPERATURE,
      systemPrompt = '',
      type = 'generate',
      includeHistory = false,
    } = options;

    // Check cache
    const cacheKey = this.cache.generateKey(prompt, model, { systemPrompt, maxTokens, temperature });
    const cachedResponse = this.cache.get(cacheKey);
    if (cachedResponse) {
      return cachedResponse;
    }

    try {
      // Get current session
      const { data: { session } } = await supabase.auth.getSession();
      if (!session) {
        throw new Error('Not authenticated. Please sign in first.');
      }

      // Build messages array
      const messages = [];
      
      if (systemPrompt) {
        messages.push({ role: 'system', content: systemPrompt });
      }
      
      if (includeHistory && this.sharedState.conversationHistory.length > 0) {
        // Include last 10 messages for context
        messages.push(...this.sharedState.conversationHistory.slice(-10));
      }
      
      messages.push({ role: 'user', content: prompt });

      // Call Supabase Edge Function
      const response = await this.rateLimitedCall(async () => {
        const { data, error } = await supabase.functions.invoke('ai', {
          body: {
            type,
            prompt,
            messages,
            model,
            max_tokens: maxTokens,
            temperature,
          },
        });

        if (error) throw error;
        return data;
      });

      // Process response
      const aiResponse: AIResponse = {
        content: response.result || response.content || '',
        usage: response.usage,
        model: response.model || model,
      };

      // Update statistics
      if (aiResponse.usage) {
        this.sharedState.updateStatistics(aiResponse.usage.total_tokens, model);
      }

      // Cache response
      this.cache.set(cacheKey, aiResponse);

      // Add to history if requested
      if (includeHistory) {
        this.sharedState.addToHistory('assistant', aiResponse.content);
      }

      return aiResponse;

    } catch (error) {
      console.error('AI Service Error:', error);
      const errorMessage = getErrorMessage(error);
      
      // Retry with fallback model if primary fails
      if (model === CONFIG.MODEL_PRIMARY && options.model !== CONFIG.MODEL_FALLBACK) {
        console.log('Retrying with fallback model...');
        return this.callAI(prompt, { ...options, model: CONFIG.MODEL_FALLBACK });
      }
      
      this.sharedState.lastError = errorMessage;
      
      if (error instanceof Response) {
        throw new APIError(`AI service error: ${errorMessage}`, error.status);
      } else if (errorMessage.includes('Not authenticated')) {
        throw new AuthenticationError(errorMessage);
      } else if (errorMessage.includes('network') || errorMessage.includes('fetch')) {
        throw new NetworkError(errorMessage);
      }
      
      throw new APIError(`AI service error: ${errorMessage}`);
    }
  }

  /**
   * Preprocess Excel data for AI analysis
   */
  preprocessData(data: unknown[][]): string {
    if (!data || data.length === 0) return '';
    
    // Convert 2D array to CSV-like format
    const headers = data[0];
    const rows = data.slice(1);
    
    let result = headers.join('\t') + '\n';
    result += rows.map(row => row.join('\t')).join('\n');
    
    // Limit to 1000 characters for API efficiency
    const MAX_PREVIEW_LENGTH = 1000;
    if (result.length > MAX_PREVIEW_LENGTH) {
      result = result.substring(0, MAX_PREVIEW_LENGTH) + '...(truncated)';
    }
    
    return result;
  }

  /**
   * Build context for AI operations
   */
  buildContext(
    data: unknown,
    operation: string,
    additionalContext: string = ''
  ): string {
    const dataPreview = Array.isArray(data) 
      ? this.preprocessData(data)
      : String(data);
    
    const dataSize = Array.isArray(data) 
      ? `${data.length} rows, ${data[0]?.length || 0} columns`
      : '1 value';
    
    let context = `Operation: ${operation}\n`;
    context += `Data size: ${dataSize}\n`;
    
    if (additionalContext) {
      context += `Context: ${additionalContext}\n`;
    }
    
    if (this.sharedState.currentWorkbook) {
      context += `Workbook: ${this.sharedState.currentWorkbook}\n`;
    }
    
    context += `\nData preview:\n${dataPreview}`;
    
    return context;
  }

  /**
   * Specialized method for data analysis
   */
  async analyzeData(
    data: unknown[][],
    analysisType: string = 'summary',
    options: Record<string, unknown> = {}
  ): Promise<string> {
    const context = this.buildContext(data, 'analysis', `Analysis type: ${analysisType}`);
    
    const systemPrompt = `You are an expert data analyst. Provide concise, actionable insights based on the data provided. 
    Focus on: trends, patterns, anomalies, and key statistics. 
    Keep responses under 200 words and include specific numbers when relevant.`;
    
    let prompt = `Analyze this data with focus on ${analysisType}:\n${context}`;
    
    // Customize prompt based on analysis type
    switch (analysisType.toLowerCase()) {
      case 'trend':
        prompt += '\n\nFocus on trends, growth patterns, and directional changes.';
        break;
      case 'correlation':
        prompt += '\n\nFocus on relationships and correlations between variables.';
        break;
      case 'insights':
        prompt += '\n\nFocus on key insights, opportunities, and recommendations.';
        break;
      case 'summary':
      default:
        prompt += '\n\nProvide a comprehensive summary with key statistics and main findings.';
    }

    const response = await this.callAI(prompt, {
      systemPrompt,
      maxTokens: 300,
      type: 'analyze',
      ...options,
    });

    return response.content;
  }

  /**
   * Generate Excel formulas
   */
  async generateFormula(
    description: string,
    dataRange?: string,
    options: Record<string, unknown> = {}
  ): Promise<string> {
    const systemPrompt = `You are an Excel formula expert. Generate ONLY the Excel formula (no explanations) 
    based on the user's description. Use proper Excel syntax and functions. 
    If referencing data, use the provided range or create generic references like A1:A10.`;
    
    let prompt = `Generate an Excel formula for: ${description}`;
    if (dataRange) {
      prompt += `\nData range: ${dataRange}`;
    }
    prompt += '\n\nReturn ONLY the formula starting with =';

    const response = await this.callAI(prompt, {
      systemPrompt,
      maxTokens: 150,
      ...options,
    });
    
    // Ensure formula starts with =
    const formula = response.content.trim();
    return formula.startsWith('=') ? formula : `=${formula}`;
  }

  /**
   * Clear cache
   */
  clearCache(): void {
    this.cache.clear();
  }

  /**
   * Get current statistics
   */
  getStatistics() {
    return this.sharedState.statistics;
  }

  /**
   * Get conversation history
   */
  getHistory() {
    return this.sharedState.conversationHistory;
  }

  /**
   * Set current workbook context
   */
  setWorkbookContext(workbookName: string): void {
    this.sharedState.currentWorkbook = workbookName;
  }
}

// Export singleton instance
export const aiService = AIService.getInstance();
export const sharedState = SharedState.getInstance();