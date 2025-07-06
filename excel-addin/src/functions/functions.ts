/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global clearInterval, console, CustomFunctions, setInterval, Office */

import { createClient, SupabaseClient } from '@supabase/supabase-js';
import { getSessionManager, SessionManager } from '../services/session-manager';

// Initialize Supabase client for custom functions
declare const process: any;
const supabaseUrl = process.env.VITE_SUPABASE_URL || 'https://your-project.supabase.co';
const supabaseAnonKey = process.env.VITE_SUPABASE_ANON_KEY || 'your-anon-key';
let sessionManager: SessionManager;
let supabase: SupabaseClient;
let initialized = false;

// Initialize on first use
async function initSupabase() {
  if (!initialized) {
    // Wait for Office to be ready
    await new Promise<void>((resolve) => {
      if (Office.context) {
        resolve();
      } else {
        Office.onReady(() => resolve());
      }
    });
    
    sessionManager = getSessionManager(supabaseUrl, supabaseAnonKey);
    supabase = sessionManager.getSupabaseClient();
    await sessionManager.initialize();
    initialized = true;
  }
}

/**
 * Analyze Excel data using AI
 * @customfunction
 * @param {any[][]} data The data range to analyze
 * @param {string} [prompt] Optional analysis prompt
 * @returns {Promise<string>} The AI analysis result
 */
export async function AIANALYZE(data: any[][], prompt?: string): Promise<string> {
  try {
    await initSupabase();
    
    // Check if user is authenticated
    const session = await sessionManager.getSession();
    if (!session) {
      return "ERROR: Please sign in using the ExcelAiRate task pane";
    }
    
    // Call AI endpoint
    const { data: result, error } = await supabase.functions.invoke('ai', {
      body: {
        type: 'analyze',
        data: data,
        prompt: prompt
      }
    });
    
    if (error) {
      return `ERROR: ${error.message}`;
    }
    
    return result.result || "No analysis available";
    
  } catch (error: any) {
    return `ERROR: ${error.message || 'Analysis failed'}`;
  }
}

/**
 * Generate content using AI
 * @customfunction
 * @param {string} prompt The generation prompt
 * @param {string} [context] Optional context data
 * @returns {Promise<string>} The generated content
 */
export async function AIGENERATE(prompt: string, context?: string): Promise<string> {
  try {
    await initSupabase();
    
    // Check if user is authenticated
    const session = await sessionManager.getSession();
    if (!session) {
      return "ERROR: Please sign in using the ExcelAiRate task pane";
    }
    
    // Call AI endpoint
    const { data: result, error } = await supabase.functions.invoke('ai', {
      body: {
        type: 'generate',
        prompt: prompt,
        context: context
      }
    });
    
    if (error) {
      return `ERROR: ${error.message}`;
    }
    
    return result.result || "No content generated";
    
  } catch (error: any) {
    return `ERROR: ${error.message || 'Generation failed'}`;
  }
}

/**
 * Explain a formula or concept
 * @customfunction
 * @param {string} formula The formula or concept to explain
 * @returns {Promise<string>} The explanation
 */
export async function AIEXPLAIN(formula: string): Promise<string> {
  try {
    await initSupabase();
    
    // Check if user is authenticated
    const session = await sessionManager.getSession();
    if (!session) {
      return "ERROR: Please sign in using the ExcelAiRate task pane";
    }
    
    // Call AI endpoint
    const { data: result, error } = await supabase.functions.invoke('ai', {
      body: {
        type: 'explain',
        prompt: formula
      }
    });
    
    if (error) {
      return `ERROR: ${error.message}`;
    }
    
    return result.result || "No explanation available";
    
  } catch (error: any) {
    return `ERROR: ${error.message || 'Explanation failed'}`;
  }
}

/**
 * Summarize data using AI
 * @customfunction
 * @param {any[][]} data The data range to summarize
 * @param {number} [maxLength] Maximum length of summary
 * @returns {Promise<string>} The summary
 */
export async function AISUMMARIZE(data: any[][], maxLength?: number): Promise<string> {
  try {
    await initSupabase();
    
    // Check if user is authenticated
    const session = await sessionManager.getSession();
    if (!session) {
      return "ERROR: Please sign in using the ExcelAiRate task pane";
    }
    
    // Call AI endpoint
    const { data: result, error } = await supabase.functions.invoke('ai', {
      body: {
        type: 'analyze',
        data: data,
        prompt: `Summarize this data concisely${maxLength ? ` in about ${maxLength} words` : ''}`
      }
    });
    
    if (error) {
      return `ERROR: ${error.message}`;
    }
    
    return result.result || "No summary available";
    
  } catch (error: any) {
    return `ERROR: ${error.message || 'Summary failed'}`;
  }
}

/**
 * Get usage information for the current user
 * @customfunction
 * @returns {Promise<string>} Usage count and limit
 */
export async function AIUSAGE(): Promise<string> {
  try {
    await initSupabase();
    
    // Check if user is authenticated
    const session = await sessionManager.getSession();
    if (!session) {
      return "ERROR: Please sign in using the ExcelAiRate task pane";
    }
    
    // Get user profile
    const { data: profile, error } = await supabase
      .from('profiles')
      .select('usage_count, usage_limit')
      .eq('id', session.user.id)
      .single();
    
    if (error) {
      return `ERROR: ${error.message}`;
    }
    
    return `${profile.usage_count || 0} / ${profile.usage_limit || 10}`;
    
  } catch (error: any) {
    return `ERROR: ${error.message || 'Failed to get usage'}`;
  }
}