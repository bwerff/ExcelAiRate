import { serve } from 'https://deno.land/std@0.168.0/http/server.ts'
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { Configuration, OpenAIApi } from 'https://esm.sh/openai@4.24.0'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

interface AIRequest {
  type: 'analyze' | 'generate' | 'explain'
  prompt: string
  data?: string
  options?: {
    model?: string
    temperature?: number
  }
}

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    // Initialize Supabase
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const supabaseServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    const supabase = createClient(supabaseUrl, supabaseServiceKey)

    // Get user from auth header
    const authHeader = req.headers.get('Authorization')!
    const token = authHeader.replace('Bearer ', '')
    const { data: { user }, error: authError } = await supabase.auth.getUser(token)
    
    if (authError || !user) {
      throw new Error('Unauthorized')
    }

    // Check if user can query
    const { data: canQuery } = await supabase
      .rpc('can_user_query', { user_id: user.id })

    if (!canQuery) {
      return new Response(
        JSON.stringify({ 
          error: 'Usage limit exceeded',
          code: 'USAGE_LIMIT_EXCEEDED'
        }),
        { status: 429, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      )
    }

    // Parse request
    const { type, prompt, data, options = {} } = await req.json() as AIRequest
    const model = options.model || 'gpt-3.5-turbo'
    const temperature = options.temperature || 0.7

    // Create cache key
    const cacheKey = await createHash(`${type}:${prompt}:${data || ''}:${model}`)

    // Check cache
    const { data: cached } = await supabase
      .from('ai_cache')
      .select('response')
      .eq('prompt_hash', cacheKey)
      .single()

    if (cached) {
      // Update cache hit count
      await supabase
        .from('ai_cache')
        .update({ 
          hit_count: cached.hit_count + 1,
          last_accessed_at: new Date().toISOString()
        })
        .eq('prompt_hash', cacheKey)

      // Log usage
      await logUsage(supabase, user.id, type, prompt, cached.response, model, 0, 0, true)

      return new Response(
        JSON.stringify({ 
          success: true,
          result: cached.response,
          cached: true 
        }),
        { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      )
    }

    // Call OpenAI
    const openaiKey = Deno.env.get('OPENAI_API_KEY')!
    const openai = new OpenAIApi(new Configuration({ apiKey: openaiKey }))

    const systemPrompt = getSystemPrompt(type)
    const userPrompt = getUserPrompt(type, prompt, data)

    const startTime = Date.now()
    const completion = await openai.createChatCompletion({
      model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature,
      max_tokens: type === 'explain' ? 500 : 2000,
      ...(type === 'analyze' && { response_format: { type: 'json_object' } })
    })

    const result = completion.data.choices[0]?.message?.content
    if (!result) throw new Error('No response from AI')

    const responseTime = Date.now() - startTime
    const tokens = completion.data.usage?.total_tokens || 0

    // Parse result for analysis type
    let parsedResult = result
    if (type === 'analyze') {
      try {
        parsedResult = JSON.parse(result)
      } catch {
        parsedResult = { response: result }
      }
    }

    // Store in cache
    await supabase
      .from('ai_cache')
      .insert({
        prompt_hash: cacheKey,
        prompt: `${type}:${prompt}`,
        response: parsedResult,
        model_used: model
      })

    // Log usage
    await logUsage(supabase, user.id, type, prompt, parsedResult, model, tokens, responseTime, false)

    // Increment usage counter
    await supabase.rpc('increment_usage', { user_id: user.id })

    return new Response(
      JSON.stringify({ 
        success: true,
        result: parsedResult,
        cached: false,
        usage: {
          tokens,
          response_time_ms: responseTime
        }
      }),
      { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )

  } catch (error) {
    console.error('AI function error:', error)
    return new Response(
      JSON.stringify({ 
        error: error.message || 'Internal server error',
        code: 'INTERNAL_ERROR' 
      }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  }
})

// Helper functions
async function createHash(input: string): Promise<string> {
  const encoder = new TextEncoder()
  const data = encoder.encode(input)
  const hashBuffer = await crypto.subtle.digest('SHA-256', data)
  const hashArray = Array.from(new Uint8Array(hashBuffer))
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('')
}

function getSystemPrompt(type: string): string {
  switch (type) {
    case 'analyze':
      return `You are an AI assistant specialized in Excel data analysis. 
      Analyze the provided data and return a structured JSON response with:
      {
        "summary": "Brief overview",
        "insights": ["Key insight 1", "Key insight 2", ...],
        "trends": ["Trend 1", "Trend 2", ...],
        "recommendations": ["Action 1", "Action 2", ...]
      }`
    
    case 'generate':
      return `You are an AI assistant that generates Excel-compatible data.
      Create structured data based on the user's request.
      Format the output as tab-separated values (TSV) for easy Excel paste.`
    
    case 'explain':
      return `You are an AI assistant that explains Excel data clearly and concisely.
      Provide a brief, easy-to-understand explanation of the data patterns.`
    
    default:
      return 'You are a helpful AI assistant for Excel users.'
  }
}

function getUserPrompt(type: string, prompt: string, data?: string): string {
  switch (type) {
    case 'analyze':
      return `Analyze this data: ${data}\n\nUser question: ${prompt}`
    
    case 'generate':
      return `Generate data based on this request: ${prompt}`
    
    case 'explain':
      return `Explain this data in simple terms: ${data}\n\nFocus on: ${prompt}`
    
    default:
      return prompt
  }
}

async function logUsage(
  supabase: any,
  userId: string,
  actionType: string,
  prompt: string,
  response: any,
  model: string,
  tokens: number,
  responseTime: number,
  fromCache: boolean
) {
  await supabase
    .from('usage_logs')
    .insert({
      user_id: userId,
      action_type: actionType,
      prompt,
      response,
      model_used: model,
      tokens_used: tokens,
      response_time_ms: responseTime,
      from_cache: fromCache
    })
}