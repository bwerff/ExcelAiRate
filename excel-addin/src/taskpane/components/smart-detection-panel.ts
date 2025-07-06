/**
 * Smart Detection Panel Component
 * UI for displaying smart range detection results and context
 */

import { smartDetection, SmartContext, ContextSuggestion } from '../../services/smart-detection';

/* global document */

export class SmartDetectionPanel {
  private container: HTMLElement;
  private currentContext: SmartContext | null = null;

  constructor(containerId: string) {
    this.container = document.getElementById(containerId) || document.createElement('div');
    this.initialize();
  }

  /**
   * Initialize the panel
   */
  private initialize(): void {
    this.container.innerHTML = `
      <div class="smart-detection-panel">
        <h3>Smart Detection</h3>
        <button id="analyze-data-btn" class="ms-Button ms-Button--primary">
          <span class="ms-Button-label">Analyze Selection</span>
        </button>
        
        <div id="detection-results" style="display: none;">
          <div class="detection-section">
            <h4>Data Information</h4>
            <div id="data-info-content"></div>
          </div>
          
          <div class="detection-section">
            <h4>Detected Patterns</h4>
            <div id="patterns-content"></div>
          </div>
          
          <div class="detection-section">
            <h4>Suggestions</h4>
            <div id="suggestions-content"></div>
          </div>
          
          <div class="detection-section">
            <h4>Related Data</h4>
            <div id="related-data-content"></div>
          </div>
        </div>
        
        <div id="detection-loading" style="display: none;">
          <div class="spinner"></div>
          <p>Analyzing data...</p>
        </div>
      </div>
    `;

    this.attachEventListeners();
    this.addStyles();
  }

  /**
   * Attach event listeners
   */
  private attachEventListeners(): void {
    const analyzeBtn = document.getElementById('analyze-data-btn');
    if (analyzeBtn) {
      analyzeBtn.onclick = () => this.analyzeSelection();
    }
  }

  /**
   * Analyze current selection
   */
  private async analyzeSelection(): Promise<void> {
    try {
      this.showLoading();
      
      const context = await smartDetection.analyzeSelection();
      this.currentContext = context;
      
      this.displayResults(context);
      this.hideLoading();
      
    } catch (error) {
      console.error('Analysis failed:', error);
      this.showError('Failed to analyze selection. Please try again.');
      this.hideLoading();
    }
  }

  /**
   * Display analysis results
   */
  private displayResults(context: SmartContext): void {
    const resultsDiv = document.getElementById('detection-results');
    if (!resultsDiv) return;
    
    resultsDiv.style.display = 'block';
    
    // Display data information
    this.displayDataInfo(context.primaryData);
    
    // Display patterns
    this.displayPatterns(context.primaryData.patterns || []);
    
    // Display suggestions
    this.displaySuggestions(context.suggestions);
    
    // Display related data
    this.displayRelatedData(context.relatedData);
  }

  /**
   * Display data information
   */
  private displayDataInfo(dataInfo: any): void {
    const content = document.getElementById('data-info-content');
    if (!content) return;

    const formatDataType = (type: string) => {
      const icons: Record<string, string> = {
        'numeric': '🔢',
        'text': '📝',
        'date': '📅',
        'currency': '💰',
        'percentage': '📊',
        'boolean': '✓✗',
        'mixed': '🔀',
        'empty': '⚪'
      };
      return `${icons[type] || '❓'} ${type}`;
    };

    content.innerHTML = `
      <div class="info-grid">
        <div class="info-item">
          <label>Range:</label>
          <span>${dataInfo.range}</span>
        </div>
        <div class="info-item">
          <label>Data Type:</label>
          <span>${formatDataType(dataInfo.dataType)}</span>
        </div>
        <div class="info-item">
          <label>Has Headers:</label>
          <span>${dataInfo.hasHeaders ? '✓ Yes' : '✗ No'}</span>
        </div>
        ${dataInfo.headers.length > 0 ? `
          <div class="info-item full-width">
            <label>Headers:</label>
            <div class="header-list">
              ${dataInfo.headers.map((h: string) => `<span class="header-tag">${h}</span>`).join('')}
            </div>
          </div>
        ` : ''}
        ${dataInfo.statistics ? `
          <div class="info-item full-width">
            <label>Statistics:</label>
            <div class="stats-grid">
              <span>Count: ${dataInfo.statistics.count}</span>
              <span>Average: ${dataInfo.statistics.average?.toFixed(2) || 'N/A'}</span>
              <span>Min: ${dataInfo.statistics.min || 'N/A'}</span>
              <span>Max: ${dataInfo.statistics.max || 'N/A'}</span>
            </div>
          </div>
        ` : ''}
        <div class="info-item">
          <label>Null Values:</label>
          <span>${dataInfo.nullCount}</span>
        </div>
        <div class="info-item">
          <label>Unique Values:</label>
          <span>${dataInfo.uniqueValues || 'N/A'}</span>
        </div>
      </div>
    `;
  }

  /**
   * Display detected patterns
   */
  private displayPatterns(patterns: any[]): void {
    const content = document.getElementById('patterns-content');
    if (!content) return;

    if (patterns.length === 0) {
      content.innerHTML = '<p class="no-data">No patterns detected</p>';
      return;
    }

    const patternIcons: Record<string, string> = {
      'email': '✉️',
      'phone': '📞',
      'url': '🔗',
      'id': '🏷️',
      'postal': '📮'
    };

    content.innerHTML = patterns.map(pattern => `
      <div class="pattern-item">
        <div class="pattern-header">
          <span class="pattern-icon">${patternIcons[pattern.type] || '🔍'}</span>
          <span class="pattern-type">${pattern.type}</span>
          <span class="pattern-confidence">${Math.round(pattern.confidence * 100)}% confident</span>
        </div>
        ${pattern.examples.length > 0 ? `
          <div class="pattern-examples">
            Examples: ${pattern.examples.map(ex => `<code>${ex}</code>`).join(', ')}
          </div>
        ` : ''}
      </div>
    `).join('');
  }

  /**
   * Display suggestions
   */
  private displaySuggestions(suggestions: ContextSuggestion[]): void {
    const content = document.getElementById('suggestions-content');
    if (!content) return;

    if (suggestions.length === 0) {
      content.innerHTML = '<p class="no-data">No suggestions available</p>';
      return;
    }

    const suggestionIcons: Record<string, string> = {
      'analysis': '📊',
      'formatting': '🎨',
      'validation': '✅',
      'relationship': '🔗'
    };

    content.innerHTML = suggestions.map((suggestion, index) => `
      <div class="suggestion-item">
        <div class="suggestion-header">
          <span class="suggestion-icon">${suggestionIcons[suggestion.type] || '💡'}</span>
          <span class="suggestion-description">${suggestion.description}</span>
        </div>
        <div class="suggestion-footer">
          <span class="confidence-badge">${Math.round(suggestion.confidence * 100)}%</span>
          ${suggestion.action ? `
            <button class="ms-Button ms-Button--small apply-suggestion-btn" data-index="${index}">
              Apply
            </button>
          ` : ''}
        </div>
      </div>
    `).join('');

    // Attach event listeners to apply buttons
    const applyButtons = content.querySelectorAll('.apply-suggestion-btn');
    applyButtons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt((e.target as HTMLElement).getAttribute('data-index') || '0');
        this.applySuggestion(suggestions[index]);
      });
    });
  }

  /**
   * Display related data
   */
  private displayRelatedData(relatedData: any[]): void {
    const content = document.getElementById('related-data-content');
    if (!content) return;

    if (relatedData.length === 0) {
      content.innerHTML = '<p class="no-data">No related data found</p>';
      return;
    }

    content.innerHTML = relatedData.map(data => `
      <div class="related-item">
        <div class="related-header">
          <span class="related-range">${data.range}</span>
          <span class="related-type">${data.dataType}</span>
        </div>
        ${data.headers.length > 0 ? `
          <div class="related-headers">
            ${data.headers.slice(0, 3).join(', ')}${data.headers.length > 3 ? '...' : ''}
          </div>
        ` : ''}
      </div>
    `).join('');
  }

  /**
   * Apply a suggestion
   */
  private async applySuggestion(suggestion: ContextSuggestion): Promise<void> {
    try {
      if (suggestion.action) {
        await suggestion.action();
        this.showSuccess(`Applied: ${suggestion.description}`);
        // Refresh analysis
        await this.analyzeSelection();
      }
    } catch (error) {
      console.error('Failed to apply suggestion:', error);
      this.showError('Failed to apply suggestion');
    }
  }

  /**
   * Show loading state
   */
  private showLoading(): void {
    const loading = document.getElementById('detection-loading');
    const results = document.getElementById('detection-results');
    
    if (loading) loading.style.display = 'block';
    if (results) results.style.display = 'none';
  }

  /**
   * Hide loading state
   */
  private hideLoading(): void {
    const loading = document.getElementById('detection-loading');
    if (loading) loading.style.display = 'none';
  }

  /**
   * Show error message
   */
  private showError(message: string): void {
    const content = document.getElementById('detection-results');
    if (content) {
      content.innerHTML = `<div class="error-message">${message}</div>`;
      content.style.display = 'block';
    }
  }

  /**
   * Show success message
   */
  private showSuccess(message: string): void {
    // Create temporary success notification
    const notification = document.createElement('div');
    notification.className = 'success-notification';
    notification.textContent = message;
    this.container.appendChild(notification);
    
    setTimeout(() => {
      notification.remove();
    }, 3000);
  }

  /**
   * Add custom styles
   */
  private addStyles(): void {
    const style = document.createElement('style');
    style.textContent = `
      .smart-detection-panel {
        padding: 15px;
      }
      
      .detection-section {
        margin-top: 20px;
        padding: 10px;
        background: #f5f5f5;
        border-radius: 4px;
      }
      
      .detection-section h4 {
        margin: 0 0 10px 0;
        color: #333;
      }
      
      .info-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
      }
      
      .info-item {
        display: flex;
        flex-direction: column;
      }
      
      .info-item.full-width {
        grid-column: 1 / -1;
      }
      
      .info-item label {
        font-weight: 600;
        color: #666;
        font-size: 12px;
      }
      
      .header-list {
        display: flex;
        flex-wrap: wrap;
        gap: 5px;
        margin-top: 5px;
      }
      
      .header-tag {
        background: #e1f5fe;
        color: #0277bd;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 12px;
      }
      
      .stats-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 5px;
        font-size: 12px;
      }
      
      .pattern-item {
        background: white;
        padding: 10px;
        margin-bottom: 8px;
        border-radius: 4px;
        border: 1px solid #e0e0e0;
      }
      
      .pattern-header {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      
      .pattern-icon {
        font-size: 20px;
      }
      
      .pattern-type {
        font-weight: 600;
        text-transform: capitalize;
      }
      
      .pattern-confidence {
        margin-left: auto;
        font-size: 12px;
        color: #666;
      }
      
      .pattern-examples {
        margin-top: 5px;
        font-size: 12px;
        color: #666;
      }
      
      .pattern-examples code {
        background: #f5f5f5;
        padding: 2px 4px;
        border-radius: 2px;
      }
      
      .suggestion-item {
        background: white;
        padding: 12px;
        margin-bottom: 8px;
        border-radius: 4px;
        border: 1px solid #e0e0e0;
        transition: box-shadow 0.2s;
      }
      
      .suggestion-item:hover {
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      }
      
      .suggestion-header {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      
      .suggestion-icon {
        font-size: 20px;
      }
      
      .suggestion-description {
        flex: 1;
      }
      
      .suggestion-footer {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-top: 8px;
      }
      
      .confidence-badge {
        background: #4CAF50;
        color: white;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
      }
      
      .apply-suggestion-btn {
        min-width: 60px;
      }
      
      .related-item {
        background: white;
        padding: 10px;
        margin-bottom: 8px;
        border-radius: 4px;
        border: 1px solid #e0e0e0;
      }
      
      .related-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      
      .related-range {
        font-weight: 600;
        color: #0078d4;
      }
      
      .related-type {
        font-size: 12px;
        color: #666;
      }
      
      .related-headers {
        margin-top: 5px;
        font-size: 12px;
        color: #666;
      }
      
      .no-data {
        color: #999;
        text-align: center;
        padding: 20px;
      }
      
      .error-message {
        background: #ffebee;
        color: #c62828;
        padding: 10px;
        border-radius: 4px;
        text-align: center;
      }
      
      .success-notification {
        position: fixed;
        top: 20px;
        right: 20px;
        background: #4CAF50;
        color: white;
        padding: 12px 20px;
        border-radius: 4px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        animation: slideIn 0.3s ease-out;
      }
      
      @keyframes slideIn {
        from {
          transform: translateX(100%);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
      
      .spinner {
        border: 3px solid #f3f3f3;
        border-top: 3px solid #0078d4;
        border-radius: 50%;
        width: 40px;
        height: 40px;
        animation: spin 1s linear infinite;
        margin: 20px auto;
      }
      
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    `;
    document.head.appendChild(style);
  }

  /**
   * Refresh the panel
   */
  public refresh(): void {
    if (this.currentContext) {
      this.displayResults(this.currentContext);
    }
  }

  /**
   * Clear the panel
   */
  public clear(): void {
    const results = document.getElementById('detection-results');
    if (results) {
      results.style.display = 'none';
    }
    this.currentContext = null;
  }
}

// Export for use in main taskpane
export function initializeSmartDetectionPanel(containerId: string): SmartDetectionPanel {
  return new SmartDetectionPanel(containerId);
}