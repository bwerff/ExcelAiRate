// Shared CORS utility for Edge Functions

// Get allowed origins from environment
const allowedOrigins = (Deno.env.get('ALLOWED_ORIGINS') || 'http://localhost:3000,http://localhost:5173').split(',')

export function getCorsHeaders(origin: string | null) {
  // Check if origin is allowed
  const allowedOrigin = origin && allowedOrigins.includes(origin) ? origin : allowedOrigins[0]
  
  return {
    'Access-Control-Allow-Origin': allowedOrigin,
    'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Max-Age': '86400', // 24 hours
  }
}

export function corsResponse(origin: string | null) {
  return new Response('ok', { headers: getCorsHeaders(origin) })
}