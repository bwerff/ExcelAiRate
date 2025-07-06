/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { createClient, SupabaseClient } from '@supabase/supabase-js';
import { getSessionManager, SessionManager } from '../services/session-manager';

// Initialize Supabase client
declare const process: any;
const supabaseUrl = process.env.VITE_SUPABASE_URL || 'https://your-project.supabase.co';
const supabaseAnonKey = process.env.VITE_SUPABASE_ANON_KEY || 'your-anon-key';
let supabase: SupabaseClient;
let sessionManager: SessionManager;

// Global state
let currentUser: any = null;
let lastResult: string = '';

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // Initialize session manager
    sessionManager = getSessionManager(supabaseUrl, supabaseAnonKey);
    supabase = sessionManager.getSupabaseClient();
    
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    
    // Event listeners
    document.getElementById("sign-in")!.onclick = signIn;
    document.getElementById("sign-out")!.onclick = signOut;
    document.getElementById("analyze-selection")!.onclick = analyzeSelection;
    document.getElementById("generate-content")!.onclick = generateContent;
    document.getElementById("insert-result")!.onclick = insertResult;
    
    // Initialize session manager with callback
    await sessionManager.initialize((session) => {
      if (session) {
        currentUser = session.user;
        showMainSection();
        loadUserData();
      } else {
        currentUser = null;
        showAuthSection();
      }
    });
    
    // Check initial auth state
    checkAuthState();
  }
});

async function checkAuthState() {
  try {
    const session = await sessionManager.getSession();
    if (session) {
      currentUser = session.user;
      showMainSection();
      await loadUserData();
    } else {
      showAuthSection();
    }
  } catch (error) {
    console.error('Auth check error:', error);
    showAuthSection();
  }
}

async function signIn() {
  const emailInput = document.getElementById("email-input") as HTMLInputElement;
  const email = emailInput.value.trim();
  
  if (!email) {
    showMessage('auth-message', 'Please enter your email address', 'error');
    return;
  }
  
  try {
    showMessage('auth-message', 'Sending magic link...', 'info');
    
    const { error } = await supabase.auth.signInWithOtp({
      email: email,
      options: {
        emailRedirectTo: window.location.origin
      }
    });
    
    if (error) throw error;
    
    showMessage('auth-message', 'Check your email for the magic link!', 'success');
    emailInput.value = '';
    
    // Session manager already handles auth state changes via the callback
    
  } catch (error: any) {
    showMessage('auth-message', error.message || 'Sign in failed', 'error');
  }
}

async function signOut() {
  try {
    await sessionManager.signOut();
    currentUser = null;
    showAuthSection();
  } catch (error) {
    console.error('Sign out error:', error);
  }
}

async function loadUserData() {
  if (!currentUser) return;
  
  // Update UI with user email
  document.getElementById("user-email")!.textContent = currentUser.email || '';
  
  try {
    // Get user profile
    const { data: profile } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', currentUser.id)
      .single();
    
    if (profile) {
      document.getElementById("usage-count")!.textContent = profile.usage_count || '0';
      document.getElementById("usage-limit")!.textContent = profile.usage_limit || '10';
    }
  } catch (error) {
    console.error('Load user data error:', error);
  }
}

async function analyzeSelection() {
  await Excel.run(async (context) => {
    try {
      showMessage('message', 'Analyzing selection...', 'info');
      
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load(['values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      if (!range.values || range.rowCount === 0 || range.columnCount === 0) {
        showMessage('message', 'Please select some data to analyze', 'error');
        return;
      }
      
      // Call AI endpoint
      const response = await callAIEndpoint({
        type: 'analyze',
        data: range.values
      });
      
      if (response.result) {
        lastResult = response.result;
        showResult(response.result);
        showMessage('message', 'Analysis complete!', 'success');
        await loadUserData(); // Refresh usage count
      }
      
    } catch (error: any) {
      showMessage('message', error.message || 'Analysis failed', 'error');
    }
  });
}

async function generateContent() {
  const promptInput = document.getElementById("generate-prompt") as HTMLInputElement;
  const prompt = promptInput.value.trim();
  
  if (!prompt) {
    showMessage('message', 'Please enter a prompt', 'error');
    return;
  }
  
  try {
    showMessage('message', 'Generating content...', 'info');
    
    // Call AI endpoint
    const response = await callAIEndpoint({
      type: 'generate',
      prompt: prompt
    });
    
    if (response.result) {
      lastResult = response.result;
      showResult(response.result);
      showMessage('message', 'Content generated!', 'success');
      promptInput.value = '';
      await loadUserData(); // Refresh usage count
    }
    
  } catch (error: any) {
    showMessage('message', error.message || 'Generation failed', 'error');
  }
}

async function insertResult() {
  if (!lastResult) {
    showMessage('message', 'No result to insert', 'error');
    return;
  }
  
  await Excel.run(async (context) => {
    try {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      
      // Split result into lines for multi-cell insertion
      const lines = lastResult.split('\n').filter(line => line.trim());
      const values = lines.map(line => [line]);
      
      // Set values
      range.getResizedRange(values.length - 1, 0).values = values;
      
      await context.sync();
      showMessage('message', 'Result inserted!', 'success');
      
    } catch (error: any) {
      showMessage('message', 'Failed to insert result', 'error');
    }
  });
}

async function callAIEndpoint(payload: any) {
  const { data, error } = await supabase.functions.invoke('ai', {
    body: payload
  });
  
  if (error) {
    throw error;
  }
  
  return data;
}

function showMessage(elementId: string, message: string, type: 'error' | 'success' | 'info') {
  const element = document.getElementById(elementId)!;
  element.textContent = message;
  element.className = `message ${type}`;
  element.style.display = 'block';
  
  // Auto-hide after 5 seconds for success/info messages
  if (type !== 'error') {
    setTimeout(() => {
      element.style.display = 'none';
    }, 5000);
  }
}

function showResult(result: string) {
  document.getElementById("result-content")!.textContent = result;
  document.getElementById("result-section")!.style.display = 'block';
}

function showAuthSection() {
  document.getElementById("auth-section")!.style.display = 'block';
  document.getElementById("main-section")!.style.display = 'none';
}

function showMainSection() {
  document.getElementById("auth-section")!.style.display = 'none';
  document.getElementById("main-section")!.style.display = 'block';
}