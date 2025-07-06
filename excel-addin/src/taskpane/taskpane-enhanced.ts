/**
 * Enhanced Task Pane for Excel AI Assistant
 * Complete implementation with all advanced features
 */

import { createClient } from '@supabase/supabase-js';
import { aiService, sharedState } from '../services/ai-service';
import { ExcelHelpers } from '../utils/excel-helpers';
import { DashboardBuilder } from '../utils/dashboard-builder';
import { DialogManager } from '../utils/dialog-manager';

/* global Excel, Office */

// Initialize Supabase client
const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL || '',
  import.meta.env.VITE_SUPABASE_ANON_KEY || ''
);

// UI Elements
interface UIElements {
  signInSection: HTMLElement;
  mainSection: HTMLElement;
  emailInput: HTMLInputElement;
  signInButton: HTMLButtonElement;
  userEmail: HTMLElement;
  signOutButton: HTMLButtonElement;
  messageDiv: HTMLElement;
  chatContainer: HTMLElement;
  userInput: HTMLTextAreaElement;
  sendButton: HTMLButtonElement;
  clearButton: HTMLButtonElement;
  totalCalls: HTMLElement;
  totalTokens: HTMLElement;
  totalCost: HTMLElement;
  quickButtons: NodeListOf<HTMLButtonElement>;
  advancedButtons: NodeListOf<HTMLButtonElement>;
  mlButtons: NodeListOf<HTMLButtonElement>;
  autoButtons: NodeListOf<HTMLButtonElement>;
  finButtons: NodeListOf<HTMLButtonElement>;
}

let ui: UIElements;
let currentUser: any = null;
let dialogManager: DialogManager;
let dashboardBuilder: DashboardBuilder;
let excelHelpers: ExcelHelpers;

/**
 * Initialize the task pane when Office is ready
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.addEventListener("DOMContentLoaded", initializeTaskPane);
  }
});

/**
 * Initialize task pane
 */
async function initializeTaskPane() {
  // Initialize UI elements
  ui = {
    signInSection: document.getElementById('sign-in-section')!,
    mainSection: document.getElementById('main-section')!,
    emailInput: document.getElementById('email-input') as HTMLInputElement,
    signInButton: document.getElementById('sign-in-button') as HTMLButtonElement,
    userEmail: document.getElementById('user-email')!,
    signOutButton: document.getElementById('sign-out-button') as HTMLButtonElement,
    messageDiv: document.getElementById('message')!,
    chatContainer: document.getElementById('chat-container')!,
    userInput: document.getElementById('user-input') as HTMLTextAreaElement,
    sendButton: document.getElementById('send-btn') as HTMLButtonElement,
    clearButton: document.getElementById('clear-btn') as HTMLButtonElement,
    totalCalls: document.getElementById('total-calls')!,
    totalTokens: document.getElementById('total-tokens')!,
    totalCost: document.getElementById('total-cost')!,
    quickButtons: document.querySelectorAll('.quick-btn') as NodeListOf<HTMLButtonElement>,
    advancedButtons: document.querySelectorAll('.advanced-btn') as NodeListOf<HTMLButtonElement>,
    mlButtons: document.querySelectorAll('.ml-btn') as NodeListOf<HTMLButtonElement>,
    autoButtons: document.querySelectorAll('.auto-btn') as NodeListOf<HTMLButtonElement>,
    finButtons: document.querySelectorAll('.fin-btn') as NodeListOf<HTMLButtonElement>,
  };

  // Initialize helpers
  dialogManager = new DialogManager();
  dashboardBuilder = new DashboardBuilder();
  excelHelpers = new ExcelHelpers();

  // Set up event listeners
  setupEventListeners();

  // Check authentication status
  await checkAuthStatus();

  // Set workbook context
  await setWorkbookContext();

  // Start statistics update interval
  setInterval(updateStatisticsDisplay, 5000);
}

/**
 * Set up all event listeners
 */
function setupEventListeners() {
  // Authentication
  ui.signInButton.addEventListener('click', handleSignIn);
  ui.signOutButton.addEventListener('click', handleSignOut);
  ui.emailInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') handleSignIn();
  });

  // Chat interface
  ui.sendButton.addEventListener('click', sendMessage);
  ui.clearButton.addEventListener('click', clearChat);
  ui.userInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter' && e.ctrlKey) {
      sendMessage();
    }
  });

  // Quick action buttons
  ui.quickButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      const action = (e.target as HTMLElement).dataset.action;
      if (action) handleQuickAction(action);
    });
  });

  // Advanced BI buttons
  ui.advancedButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      const action = (e.target as HTMLElement).dataset.action;
      if (action) handleAdvancedAction(action);
    });
  });

  // ML buttons
  ui.mlButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      const action = (e.target as HTMLElement).dataset.action;
      if (action) handleMLAction(action);
    });
  });

  // Automation buttons
  ui.autoButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      const action = (e.target as HTMLElement).dataset.action;
      if (action) handleAutomationAction(action);
    });
  });

  // Financial modeling buttons
  ui.finButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      const action = (e.target as HTMLElement).dataset.action;
      if (action) handleFinancialAction(action);
    });
  });
}

/**
 * Check authentication status
 */
async function checkAuthStatus() {
  const { data: { session } } = await supabase.auth.getSession();
  
  if (session) {
    await handleAuthSuccess(session);
  } else {
    ui.signInSection.style.display = 'block';
    ui.mainSection.style.display = 'none';
  }
}

/**
 * Handle sign in
 */
async function handleSignIn() {
  const email = ui.emailInput.value.trim();
  
  if (!email) {
    showMessage('Please enter your email address', 'error');
    return;
  }

  ui.signInButton.disabled = true;
  ui.signInButton.textContent = 'Sending...';

  try {
    const { error } = await supabase.auth.signInWithOtp({ email });
    
    if (error) throw error;
    
    showMessage('Check your email for the login link!', 'success');
    ui.emailInput.value = '';
    
    // Listen for auth state changes
    supabase.auth.onAuthStateChange(async (event, session) => {
      if (event === 'SIGNED_IN' && session) {
        await handleAuthSuccess(session);
      }
    });
  } catch (error: any) {
    showMessage(error.message || 'Sign in failed', 'error');
  } finally {
    ui.signInButton.disabled = false;
    ui.signInButton.textContent = 'Send Magic Link';
  }
}

/**
 * Handle successful authentication
 */
async function handleAuthSuccess(session: any) {
  currentUser = session.user;
  
  // Update UI
  ui.signInSection.style.display = 'none';
  ui.mainSection.style.display = 'block';
  ui.userEmail.textContent = currentUser.email;
  
  // Load user profile and usage
  await loadUserProfile();
  
  showMessage('Welcome to ExcelAiRate!', 'success');
}

/**
 * Handle sign out
 */
async function handleSignOut() {
  try {
    await supabase.auth.signOut();
    currentUser = null;
    
    ui.signInSection.style.display = 'block';
    ui.mainSection.style.display = 'none';
    
    // Clear conversation history
    clearChat();
    
    showMessage('Signed out successfully', 'info');
  } catch (error: any) {
    showMessage(error.message || 'Sign out failed', 'error');
  }
}

/**
 * Load user profile and usage information
 */
async function loadUserProfile() {
  try {
    const { data, error } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', currentUser.id)
      .single();
    
    if (error) throw error;
    
    // Update display with usage information
    if (data) {
      const usagePercent = (data.usage_count / data.usage_limit) * 100;
      showMessage(`Usage: ${data.usage_count} / ${data.usage_limit} (${usagePercent.toFixed(0)}%)`, 'info');
    }
  } catch (error: any) {
    console.error('Error loading profile:', error);
  }
}

/**
 * Set workbook context
 */
async function setWorkbookContext() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load('name');
      await context.sync();
      
      aiService.setWorkbookContext(workbook.name);
    });
  } catch (error) {
    console.error('Error setting workbook context:', error);
  }
}

/**
 * Send message to AI
 */
async function sendMessage() {
  const message = ui.userInput.value.trim();
  
  if (!message) return;
  
  ui.userInput.value = '';
  addMessageToChat('user', message);
  
  try {
    // Add loading indicator
    const loadingId = addMessageToChat('assistant', '...', true);
    
    // Check if message is asking about current selection
    let enhancedMessage = message;
    if (message.toLowerCase().includes('selection') || message.toLowerCase().includes('selected')) {
      const selectionData = await excelHelpers.getSelectedData();
      if (selectionData) {
        enhancedMessage += `\n\nCurrent selection data: ${aiService.preprocessData(selectionData)}`;
      }
    }
    
    // Call AI service
    const response = await aiService.callAI(enhancedMessage, {
      systemPrompt: getSystemPrompt(),
      maxTokens: 400,
      includeHistory: true,
    });
    
    // Update loading message with response
    updateChatMessage(loadingId, response.content);
    
    // Update statistics
    updateStatisticsDisplay();
    
  } catch (error: any) {
    addMessageToChat('error', `Error: ${error.message}`);
  }
}

/**
 * Get system prompt for chat
 */
function getSystemPrompt(): string {
  return `You are an advanced Excel AI assistant with business intelligence capabilities. Help users with:
  - Excel tasks, data analysis, formula creation, and spreadsheet optimization
  - Dashboard design and data visualization
  - PowerQuery data transformations and M language code generation  
  - PowerPivot data modeling and DAX formula creation
  - Business intelligence and advanced analytics
  - Financial modeling and scenario analysis
  - Machine learning and predictive analytics
  
  Be concise but comprehensive. If users ask about their data, provide specific insights and actionable recommendations.
  
  Available custom functions that users can use in Excel:
  =AI.ANALYZE, =AI.FORMULA, =AI.CLEAN, =AI.TRANSLATE, =AI.INSIGHTS, =AI.CATEGORIZE, =AI.GENERATE, 
  =AI.DASHBOARD, =AI.POWERQUERY, =AI.DAX, =AI.FORECAST, =AI.ANALYZE_ML, =AI.AUTOMATE, 
  =AI.SQL, =AI.FINMODEL, =AI.VALIDATE, =AI.VISUALIZE, =AI.OPTIMIZE, =AI.CONNECT
  
  For dashboard requests, consider the data type and business context to recommend appropriate visualizations, KPIs, and layout.
  For PowerQuery requests, generate efficient M code with proper error handling and documentation.
  For DAX requests, create optimized formulas following best practices for performance and accuracy.`;
}

/**
 * Handle quick actions
 */
async function handleQuickAction(action: string) {
  switch (action) {
    case 'analyze-selection':
      await analyzeSelection();
      break;
    case 'summarize-sheet':
      ui.userInput.value = 'Summarize the data in this worksheet';
      sendMessage();
      break;
    case 'find-patterns':
      ui.userInput.value = 'Find patterns and trends in my data';
      sendMessage();
      break;
    case 'generate-formula':
      ui.userInput.value = 'Help me create a formula for...';
      break;
  }
}

/**
 * Handle advanced BI actions
 */
async function handleAdvancedAction(action: string) {
  switch (action) {
    case 'create-dashboard':
      await createDashboard();
      break;
    case 'powerquery-wizard':
      await showPowerQueryWizard();
      break;
    case 'dax-helper':
      await showDAXHelper();
      break;
    case 'data-model':
      ui.userInput.value = 'Help me design a data model for my current data. Suggest relationships, measures, and optimization strategies.';
      sendMessage();
      break;
  }
}

/**
 * Handle ML actions
 */
async function handleMLAction(action: string) {
  switch (action) {
    case 'forecasting':
      await showForecastingDialog();
      break;
    case 'ml-analysis':
      await showMLAnalysisDialog();
      break;
    case 'advanced-viz':
      await showVisualizationDialog();
      break;
    case 'data-quality':
      ui.userInput.value = 'Analyze the quality of my selected data. Check for missing values, outliers, duplicates, and provide a quality score with remediation recommendations.';
      sendMessage();
      break;
  }
}

/**
 * Handle automation actions
 */
async function handleAutomationAction(action: string) {
  switch (action) {
    case 'vba-generator':
      await showVBADialog();
      break;
    case 'sql-generator':
      await showSQLDialog();
      break;
    case 'api-connect':
      await showAPIDialog();
      break;
    case 'optimize':
      ui.userInput.value = 'Analyze my current worksheet for performance optimization opportunities. Focus on formula efficiency, file size, and calculation speed.';
      sendMessage();
      break;
  }
}

/**
 * Handle financial actions
 */
async function handleFinancialAction(action: string) {
  switch (action) {
    case 'dcf-model':
      ui.userInput.value = 'Create a DCF (Discounted Cash Flow) valuation model. Include cash flow projections, discount rate analysis, terminal value calculation, and sensitivity analysis.';
      sendMessage();
      break;
    case 'scenario-analysis':
      ui.userInput.value = 'Build a scenario analysis model with base case, optimistic, and pessimistic scenarios. Include key assumptions and variance analysis.';
      sendMessage();
      break;
    case 'monte-carlo':
      ui.userInput.value = 'Create a Monte Carlo simulation model for risk analysis. Include probability distributions, random sampling, and statistical output analysis.';
      sendMessage();
      break;
    case 'budget-model':
      ui.userInput.value = 'Design a comprehensive budget model with monthly/quarterly breakdowns, variance analysis, and forecasting capabilities.';
      sendMessage();
      break;
  }
}

/**
 * Analyze current selection
 */
async function analyzeSelection() {
  try {
    const data = await excelHelpers.getSelectedData();
    if (!data) {
      showMessage('Please select some data first', 'error');
      return;
    }
    
    const analysis = await aiService.analyzeData(data, 'comprehensive');
    addMessageToChat('assistant', analysis);
    
  } catch (error: any) {
    showMessage(`Analysis error: ${error.message}`, 'error');
  }
}

/**
 * Create dashboard
 */
async function createDashboard() {
  try {
    const range = await excelHelpers.getSelectedRange();
    if (!range) {
      showMessage('Please select a data range first', 'error');
      return;
    }
    
    const dashboardConfig = await dialogManager.showDashboardDialog();
    if (!dashboardConfig.confirmed) return;
    
    // Generate dashboard design
    const design = await aiService.callAI(
      `Design a ${dashboardConfig.type} dashboard for Excel data. Provide specific chart types, KPIs, and layout recommendations.`,
      { maxTokens: 600 }
    );
    
    // Create dashboard in Excel
    await dashboardBuilder.createDashboard(range, dashboardConfig.type, design.content);
    
    showMessage('Dashboard created successfully!', 'success');
    addMessageToChat('assistant', `Dashboard created! ${design.content}`);
    
  } catch (error: any) {
    showMessage(`Dashboard error: ${error.message}`, 'error');
  }
}

/**
 * Show PowerQuery wizard
 */
async function showPowerQueryWizard() {
  const config = await dialogManager.showPowerQueryDialog();
  if (config.confirmed) {
    ui.userInput.value = `Generate PowerQuery M code to: ${config.transformation}. Source: ${config.source || 'Current worksheet'}`;
    sendMessage();
  }
}

/**
 * Show DAX helper
 */
async function showDAXHelper() {
  const config = await dialogManager.showDAXDialog();
  if (config.confirmed) {
    ui.userInput.value = `Create a DAX formula for: ${config.measure}. Table context: ${config.context || 'Auto-detect'}`;
    sendMessage();
  }
}

/**
 * Show forecasting dialog
 */
async function showForecastingDialog() {
  try {
    const selection = await excelHelpers.getSelectedData();
    if (!selection) {
      showMessage('Please select historical data first', 'error');
      return;
    }
    
    const config = await dialogManager.showForecastDialog();
    if (config.confirmed) {
      const forecast = await aiService.callAI(
        `Create a ${config.type} forecast for ${config.periods} periods using the selected data. Include confidence intervals and methodology.`,
        { maxTokens: 600 }
      );
      
      addMessageToChat('assistant', forecast.content);
    }
  } catch (error: any) {
    showMessage(`Forecast error: ${error.message}`, 'error');
  }
}

/**
 * Show ML analysis dialog
 */
async function showMLAnalysisDialog() {
  const config = await dialogManager.showMLDialog();
  if (config.confirmed) {
    ui.userInput.value = `Perform ${config.analysisType} analysis on my data. ${config.targetVariable ? 'Target variable: ' + config.targetVariable : ''} Provide statistical insights and business recommendations.`;
    sendMessage();
  }
}

/**
 * Show visualization dialog
 */
async function showVisualizationDialog() {
  const config = await dialogManager.showVisualizationDialog();
  if (config.confirmed) {
    ui.userInput.value = `Create a ${config.vizType} visualization with ${config.styleOptions}. Make it interactive and business-focused.`;
    sendMessage();
  }
}

/**
 * Show VBA dialog
 */
async function showVBADialog() {
  const config = await dialogManager.showVBADialog();
  if (config.confirmed) {
    ui.userInput.value = `Generate VBA code to: ${config.task}. Include error handling and user-friendly features.`;
    sendMessage();
  }
}

/**
 * Show SQL dialog
 */
async function showSQLDialog() {
  const config = await dialogManager.showSQLDialog();
  if (config.confirmed) {
    ui.userInput.value = `Generate a ${config.dialect} SQL query to: ${config.query}. ${config.schema ? 'Schema: ' + config.schema : ''} Include optimization tips.`;
    sendMessage();
  }
}

/**
 * Show API dialog
 */
async function showAPIDialog() {
  const config = await dialogManager.showAPIDialog();
  if (config.confirmed) {
    ui.userInput.value = `Help me connect to ${config.source} via ${config.method}. Provide step-by-step integration instructions and PowerQuery code.`;
    sendMessage();
  }
}

/**
 * Add message to chat
 */
function addMessageToChat(sender: string, message: string, isLoading: boolean = false): string {
  const messageId = `msg-${Date.now()}`;
  const messageDiv = document.createElement('div');
  messageDiv.id = messageId;
  messageDiv.className = `message ${sender}-message`;
  messageDiv.style.marginBottom = '15px';
  
  const senderColor: Record<string, string> = {
    user: '#0078d4',
    assistant: '#107c10',
    error: '#d13438'
  };
  
  messageDiv.innerHTML = `
    <div style="font-weight: bold; color: ${senderColor[sender]}; margin-bottom: 5px;">
      ${sender === 'user' ? 'You' : sender === 'error' ? 'Error' : 'AI Assistant'}
    </div>
    <div style="color: #333; white-space: pre-wrap; line-height: 1.4;">
      ${isLoading ? '<span class="loading">Thinking...</span>' : message}
    </div>
  `;
  
  ui.chatContainer.appendChild(messageDiv);
  ui.chatContainer.scrollTop = ui.chatContainer.scrollHeight;
  
  return messageId;
}

/**
 * Update chat message
 */
function updateChatMessage(messageId: string, newContent: string) {
  const messageDiv = document.getElementById(messageId);
  if (messageDiv) {
    const contentDiv = messageDiv.querySelector('div:last-child');
    if (contentDiv) {
      contentDiv.innerHTML = newContent;
    }
  }
}

/**
 * Clear chat
 */
function clearChat() {
  ui.chatContainer.innerHTML = '<div id="welcome-message" style="color: #666; font-style: italic;">Chat cleared. How can I help you?</div>';
  aiService.clearCache();
}

/**
 * Update statistics display
 */
function updateStatisticsDisplay() {
  const stats = aiService.getStatistics();
  ui.totalCalls.textContent = stats.totalCalls.toString();
  ui.totalTokens.textContent = stats.tokensUsed.toString();
  ui.totalCost.textContent = stats.costEstimate.toFixed(4);
}

/**
 * Show message
 */
function showMessage(text: string, type: 'info' | 'success' | 'error') {
  ui.messageDiv.textContent = text;
  ui.messageDiv.className = `message ${type}`;
  ui.messageDiv.style.display = 'block';
  
  setTimeout(() => {
    ui.messageDiv.style.display = 'none';
  }, 5000);
}