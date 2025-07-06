/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_SUPABASE_URL: string
  readonly VITE_SUPABASE_ANON_KEY: string
  // Note: API keys should not be in frontend code
  // OpenAI API calls should go through Supabase Edge Functions
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}

// Office.js global declarations
declare global {
  interface Window {
    Office: typeof Office;
  }
}

// Process environment for webpack builds
declare namespace NodeJS {
  interface ProcessEnv {
    VITE_SUPABASE_URL?: string;
    VITE_SUPABASE_ANON_KEY?: string;
    NODE_ENV?: 'development' | 'production' | 'test';
  }
}