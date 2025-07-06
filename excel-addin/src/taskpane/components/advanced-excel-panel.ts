/**
 * Advanced Excel Panel Component
 * UI for AI-powered PivotTables, formula dependencies, and smart formatting
 */

import { 
  advancedExcel, 
  PivotRecommendation,
  DependencyGraph,
  SmartFormattingRule,
  TableEnhancement
} from '../../services/advanced-excel';

/* global document, Office */

export class AdvancedExcelPanel {
  private container: HTMLElement;
  private currentPivotRecommendations: PivotRecommendation[] = [];
  private currentDependencyGraph: DependencyGraph | null = null;
  private currentFormattingRules: SmartFormattingRule[] = [];

  constructor(containerId: string) {
    this.container = document.getElementById(containerId) || document.createElement('div');
    this.initialize();
  }

  /**
   * Initialize the panel
   */
  private initialize(): void {
    this.container.innerHTML = `
      <div class="advanced-excel-panel">
        <div class="panel-tabs">
          <button class="tab-button active" data-tab="pivot">PivotTable AI</button>
          <button class="tab-button" data-tab="dependencies">Formula Map</button>
          <button class="tab-button" data-tab="formatting">Smart Format</button>
          <button class="tab-button" data-tab="tables">Table AI</button>
        </div>
        
        <div class="tab-content">
          <div id="pivot-tab" class="tab-panel active">
            ${this.renderPivotTab()}
          </div>
          
          <div id="dependencies-tab" class="tab-panel">
            ${this.renderDependenciesTab()}
          </div>
          
          <div id="formatting-tab" class="tab-panel">
            ${this.renderFormattingTab()}
          </div>
          
          <div id="tables-tab" class="tab-panel">
            ${this.renderTablesTab()}
          </div>
        </div>
      </div>
    `;

    this.attachEventListeners();
    this.addStyles();
  }

  /**
   * Render PivotTable tab
   */
  private renderPivotTab(): string {
    return `
      <div class="pivot-section">
        <h3>AI-Powered PivotTable Generator</h3>
        <p class="description">Let AI analyze your data and suggest the best PivotTable configuration.</p>
        
        <div class="action-group">
          <button id="analyze-pivot-btn" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">Analyze Selected Data</span>
          </button>
        </div>
        
        <div id="pivot-recommendations" class="recommendations-container" style="display: none;">
          <h4>Recommended PivotTables</h4>
          <div id="pivot-list"></div>
        </div>
        
        <div id="pivot-preview" class="preview-container" style="display: none;">
          <h4>PivotTable Preview</h4>
          <div id="pivot-config"></div>
          <button id="create-pivot-btn" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">Create PivotTable</span>
          </button>
        </div>
        
        <div id="pivot-loading" class="loading-container" style="display: none;">
          <div class="spinner"></div>
          <p>Analyzing data structure...</p>
        </div>
      </div>
    `;
  }

  /**
   * Render Dependencies tab
   */
  private renderDependenciesTab(): string {
    return `
      <div class="dependencies-section">
        <h3>Formula Dependency Mapper</h3>
        <p class="description">Visualize and analyze formula relationships in your worksheet.</p>
        
        <div class="action-group">
          <button id="map-dependencies-btn" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">Map Dependencies</span>
          </button>
          <button id="analyze-worksheet-btn" class="ms-Button">
            <span class="ms-Button-label">Analyze Entire Sheet</span>
          </button>
        </div>
        
        <div id="dependency-stats" class="stats-container" style="display: none;">
          <h4>Dependency Statistics</h4>
          <div class="stats-grid">
            <div class="stat-item">
              <span class="stat-value" id="total-formulas">0</span>
              <span class="stat-label">Formulas</span>
            </div>
            <div class="stat-item">
              <span class="stat-value" id="max-depth">0</span>
              <span class="stat-label">Max Depth</span>
            </div>
            <div class="stat-item">
              <span class="stat-value" id="circular-refs">0</span>
              <span class="stat-label">Circular Refs</span>
            </div>
            <div class="stat-item">
              <span class="stat-value" id="volatile-funcs">0</span>
              <span class="stat-label">Volatile</span>
            </div>
          </div>
        </div>
        
        <div id="dependency-viz" class="visualization-container" style="display: none;">
          <h4>Dependency Graph</h4>
          <div id="graph-container"></div>
          <div class="graph-controls">
            <button id="zoom-in-btn" class="control-btn">🔍+</button>
            <button id="zoom-out-btn" class="control-btn">🔍-</button>
            <button id="reset-view-btn" class="control-btn">⟲</button>
            <button id="export-graph-btn" class="control-btn">📥</button>
          </div>
        </div>
        
        <div id="critical-path" class="path-container" style="display: none;">
          <h4>Critical Path</h4>
          <div id="path-list"></div>
        </div>
      </div>
    `;
  }

  /**
   * Render Formatting tab
   */
  private renderFormattingTab(): string {
    return `
      <div class="formatting-section">
        <h3>Smart Conditional Formatting</h3>
        <p class="description">AI-powered formatting rules based on your data patterns.</p>
        
        <div class="action-group">
          <button id="suggest-formatting-btn" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">Suggest Formatting</span>
          </button>
          <button id="auto-format-btn" class="ms-Button">
            <span class="ms-Button-label">Auto-Format Selection</span>
          </button>
        </div>
        
        <div id="formatting-rules" class="rules-container" style="display: none;">
          <h4>Suggested Rules</h4>
          <div id="rules-list"></div>
          <div class="rule-actions">
            <button id="apply-all-rules-btn" class="ms-Button ms-Button--primary">
              <span class="ms-Button-label">Apply All</span>
            </button>
            <button id="clear-rules-btn" class="ms-Button">
              <span class="ms-Button-label">Clear</span>
            </button>
          </div>
        </div>
        
        <div id="format-preview" class="preview-container" style="display: none;">
          <h4>Format Preview</h4>
          <div id="format-sample"></div>
        </div>
      </div>
    `;
  }

  /**
   * Render Tables tab
   */
  private renderTablesTab(): string {
    return `
      <div class="tables-section">
        <h3>Table Intelligence</h3>
        <p class="description">Enhance Excel tables with AI-powered features.</p>
        
        <div class="action-group">
          <label>Select Table:</label>
          <select id="table-select" class="ms-Dropdown">
            <option value="">Loading tables...</option>
          </select>
          <button id="enhance-table-btn" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">Enhance Table</span>
          </button>
        </div>
        
        <div id="table-enhancements" class="enhancements-container" style="display: none;">
          <div class="enhancement-section">
            <h4>Calculated Columns</h4>
            <div id="calculated-columns"></div>
          </div>
          
          <div class="enhancement-section">
            <h4>Total Row Suggestions</h4>
            <div id="total-suggestions"></div>
          </div>
          
          <div class="enhancement-section">
            <h4>Recommended Slicers</h4>
            <div id="slicer-suggestions"></div>
          </div>
          
          <button id="apply-enhancements-btn" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">Apply Enhancements</span>
          </button>
        </div>
      </div>
    `;
  }

  /**
   * Attach event listeners
   */
  private attachEventListeners(): void {
    // Tab switching
    this.container.querySelectorAll('.tab-button').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const tab = (e.target as HTMLElement).dataset.tab;
        if (tab) this.switchTab(tab);
      });
    });

    // PivotTable actions
    const analyzePivotBtn = document.getElementById('analyze-pivot-btn');
    if (analyzePivotBtn) {
      analyzePivotBtn.onclick = () => this.analyzePivotData();
    }

    // Dependencies actions
    const mapDepsBtn = document.getElementById('map-dependencies-btn');
    if (mapDepsBtn) {
      mapDepsBtn.onclick = () => this.mapDependencies(false);
    }

    const analyzeSheetBtn = document.getElementById('analyze-worksheet-btn');
    if (analyzeSheetBtn) {
      analyzeSheetBtn.onclick = () => this.mapDependencies(true);
    }

    // Formatting actions
    const suggestFormatBtn = document.getElementById('suggest-formatting-btn');
    if (suggestFormatBtn) {
      suggestFormatBtn.onclick = () => this.suggestFormatting();
    }

    const autoFormatBtn = document.getElementById('auto-format-btn');
    if (autoFormatBtn) {
      autoFormatBtn.onclick = () => this.autoFormat();
    }

    // Table actions
    const enhanceTableBtn = document.getElementById('enhance-table-btn');
    if (enhanceTableBtn) {
      enhanceTableBtn.onclick = () => this.enhanceTable();
    }

    // Load tables on panel init
    this.loadTables();
  }

  /**
   * Switch between tabs
   */
  private switchTab(tabName: string): void {
    // Update tab buttons
    this.container.querySelectorAll('.tab-button').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.tab === tabName);
    });

    // Update tab panels
    this.container.querySelectorAll('.tab-panel').forEach(panel => {
      panel.classList.toggle('active', panel.id === `${tabName}-tab`);
    });
  }

  /**
   * Analyze data for PivotTable recommendations
   */
  private async analyzePivotData(): Promise<void> {
    try {
      this.showPivotLoading(true);
      
      // Get selected range
      const range = await this.getSelectedRange();
      if (!range) {
        this.showError('Please select a data range');
        this.showPivotLoading(false);
        return;
      }

      // Get recommendations
      const recommendations = await advancedExcel.analyzeDateForPivot(range);
      this.currentPivotRecommendations = recommendations;
      
      this.displayPivotRecommendations(recommendations);
      this.showPivotLoading(false);
      
    } catch (error) {
      console.error('PivotTable analysis failed:', error);
      this.showError('Failed to analyze data for PivotTable');
      this.showPivotLoading(false);
    }
  }

  /**
   * Display PivotTable recommendations
   */
  private displayPivotRecommendations(recommendations: PivotRecommendation[]): void {
    const container = document.getElementById('pivot-recommendations');
    const list = document.getElementById('pivot-list');
    
    if (!container || !list) return;
    
    if (recommendations.length === 0) {
      list.innerHTML = '<p class="no-data">No recommendations available for this data</p>';
      container.style.display = 'block';
      return;
    }

    list.innerHTML = recommendations.map((rec, index) => `
      <div class="recommendation-card" data-index="${index}">
        <div class="rec-header">
          <span class="confidence-score">${Math.round(rec.confidence * 100)}%</span>
          <h5>Option ${index + 1}</h5>
        </div>
        <p class="rec-reasoning">${rec.reasoning}</p>
        <div class="rec-config">
          <div class="config-item">
            <strong>Rows:</strong> ${rec.config.rows.map(r => r.name).join(', ') || 'None'}
          </div>
          <div class="config-item">
            <strong>Columns:</strong> ${rec.config.columns.map(c => c.name).join(', ') || 'None'}
          </div>
          <div class="config-item">
            <strong>Values:</strong> ${rec.config.values.map(v => `${v.operation}(${v.field})`).join(', ') || 'None'}
          </div>
        </div>
        <div class="rec-insights">
          <strong>Insights:</strong>
          <ul>
            ${rec.insights.map(insight => `<li>${insight}</li>`).join('')}
          </ul>
        </div>
        <button class="ms-Button ms-Button--small preview-pivot-btn" data-index="${index}">
          Preview Configuration
        </button>
      </div>
    `).join('');

    container.style.display = 'block';

    // Attach preview button listeners
    list.querySelectorAll('.preview-pivot-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt((e.target as HTMLElement).dataset.index || '0');
        this.previewPivotConfig(recommendations[index]);
      });
    });
  }

  /**
   * Preview PivotTable configuration
   */
  private previewPivotConfig(recommendation: PivotRecommendation): void {
    const preview = document.getElementById('pivot-preview');
    const config = document.getElementById('pivot-config');
    
    if (!preview || !config) return;
    
    config.innerHTML = `
      <div class="pivot-preview-content">
        <h5>PivotTable Configuration</h5>
        <div class="preview-grid">
          <div class="preview-section">
            <h6>Row Fields</h6>
            ${recommendation.config.rows.length > 0 ? 
              recommendation.config.rows.map(r => `
                <div class="field-item">
                  <span class="field-name">${r.name}</span>
                  ${r.sortOrder ? `<span class="field-sort">${r.sortOrder}</span>` : ''}
                </div>
              `).join('') : '<p class="empty">None</p>'
            }
          </div>
          
          <div class="preview-section">
            <h6>Column Fields</h6>
            ${recommendation.config.columns.length > 0 ? 
              recommendation.config.columns.map(c => `
                <div class="field-item">
                  <span class="field-name">${c.name}</span>
                </div>
              `).join('') : '<p class="empty">None</p>'
            }
          </div>
          
          <div class="preview-section">
            <h6>Value Fields</h6>
            ${recommendation.config.values.map(v => `
              <div class="field-item">
                <span class="field-name">${v.field}</span>
                <span class="field-operation">${v.operation}</span>
              </div>
            `).join('')}
          </div>
          
          <div class="preview-section">
            <h6>Filters</h6>
            ${recommendation.config.filters.length > 0 ? 
              recommendation.config.filters.map(f => `
                <div class="field-item">
                  <span class="field-name">${f.field}</span>
                </div>
              `).join('') : '<p class="empty">None</p>'
            }
          </div>
        </div>
        
        <div class="expected-value">
          <strong>What you'll learn:</strong> ${recommendation.expectedValue}
        </div>
      </div>
    `;

    preview.style.display = 'block';

    // Update create button
    const createBtn = document.getElementById('create-pivot-btn');
    if (createBtn) {
      createBtn.onclick = () => this.createPivotTable(recommendation);
    }
  }

  /**
   * Create PivotTable
   */
  private async createPivotTable(recommendation: PivotRecommendation): Promise<void> {
    try {
      const targetCell = prompt('Enter target cell for PivotTable (e.g., Sheet2!A1):', 'Sheet2!A1');
      if (!targetCell) return;

      const instructions = await advancedExcel.generatePivotTable(
        recommendation.config,
        targetCell
      );

      // Show instructions in a modal
      this.showPivotInstructions(instructions);
      
    } catch (error) {
      console.error('Failed to create PivotTable:', error);
      this.showError('Failed to generate PivotTable configuration');
    }
  }

  /**
   * Map formula dependencies
   */
  private async mapDependencies(entireSheet: boolean): Promise<void> {
    try {
      this.showDependencyLoading(true);
      
      const graph = await advancedExcel.mapFormulaDependencies(
        entireSheet ? undefined : await this.getActiveWorksheet()
      );
      
      this.currentDependencyGraph = graph;
      this.displayDependencyStats(graph);
      this.displayDependencyGraph(graph);
      this.displayCriticalPath(graph);
      
      this.showDependencyLoading(false);
      
    } catch (error) {
      console.error('Dependency mapping failed:', error);
      this.showError('Failed to map formula dependencies');
      this.showDependencyLoading(false);
    }
  }

  /**
   * Display dependency statistics
   */
  private displayDependencyStats(graph: DependencyGraph): void {
    const container = document.getElementById('dependency-stats');
    if (!container) return;
    
    document.getElementById('total-formulas')!.textContent = graph.stats.formulaCells.toString();
    document.getElementById('max-depth')!.textContent = graph.stats.maxDepth.toString();
    document.getElementById('circular-refs')!.textContent = graph.stats.circularCount.toString();
    document.getElementById('volatile-funcs')!.textContent = graph.stats.volatileFunctions.toString();
    
    container.style.display = 'block';
  }

  /**
   * Display dependency graph
   */
  private displayDependencyGraph(graph: DependencyGraph): void {
    const container = document.getElementById('dependency-viz');
    const graphDiv = document.getElementById('graph-container');
    
    if (!container || !graphDiv) return;
    
    // Create simple text representation (in real app, use D3.js)
    graphDiv.innerHTML = `
      <div class="graph-placeholder">
        <p>Dependency Graph Visualization</p>
        <p>Nodes: ${graph.nodes.length}</p>
        <p>Edges: ${graph.edges.length}</p>
        <p class="note">Full visualization requires D3.js integration</p>
      </div>
    `;
    
    container.style.display = 'block';
    
    // Attach control listeners
    this.attachGraphControls();
  }

  /**
   * Display critical path
   */
  private displayCriticalPath(graph: DependencyGraph): void {
    const container = document.getElementById('critical-path');
    const list = document.getElementById('path-list');
    
    if (!container || !list) return;
    
    if (graph.criticalPath.length === 0) {
      list.innerHTML = '<p class="no-data">No critical path found</p>';
    } else {
      list.innerHTML = `
        <div class="path-flow">
          ${graph.criticalPath.map((nodeId, index) => {
            const node = graph.nodes.find(n => n.id === nodeId);
            return `
              <div class="path-node">
                <span class="node-address">${node?.address || nodeId}</span>
                ${node?.formula ? `<code class="node-formula">${node.formula}</code>` : ''}
                ${index < graph.criticalPath.length - 1 ? '<span class="path-arrow">→</span>' : ''}
              </div>
            `;
          }).join('')}
        </div>
      `;
    }
    
    container.style.display = 'block';
  }

  /**
   * Suggest formatting
   */
  private async suggestFormatting(): Promise<void> {
    try {
      const range = await this.getSelectedRange();
      if (!range) {
        this.showError('Please select a data range');
        return;
      }

      this.showFormattingLoading(true);
      
      const rules = await advancedExcel.suggestSmartFormatting(range);
      this.currentFormattingRules = rules;
      
      this.displayFormattingRules(rules);
      this.showFormattingLoading(false);
      
    } catch (error) {
      console.error('Formatting suggestion failed:', error);
      this.showError('Failed to suggest formatting rules');
      this.showFormattingLoading(false);
    }
  }

  /**
   * Display formatting rules
   */
  private displayFormattingRules(rules: SmartFormattingRule[]): void {
    const container = document.getElementById('formatting-rules');
    const list = document.getElementById('rules-list');
    
    if (!container || !list) return;
    
    if (rules.length === 0) {
      list.innerHTML = '<p class="no-data">No formatting suggestions for this data</p>';
      container.style.display = 'block';
      return;
    }

    list.innerHTML = rules.map((rule, index) => `
      <div class="rule-card ${rule.aiGenerated ? 'ai-rule' : ''}" data-index="${index}">
        <div class="rule-header">
          <input type="checkbox" id="rule-${index}" checked />
          <label for="rule-${index}">
            <span class="rule-type">${this.formatRuleType(rule.type)}</span>
            ${rule.aiGenerated ? '<span class="ai-badge">AI</span>' : ''}
          </label>
        </div>
        <p class="rule-reasoning">${rule.reasoning}</p>
        ${rule.condition ? `<code class="rule-condition">${rule.condition}</code>` : ''}
        <button class="preview-rule-btn" data-index="${index}">Preview</button>
      </div>
    `).join('');

    container.style.display = 'block';

    // Attach preview listeners
    list.querySelectorAll('.preview-rule-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt((e.target as HTMLElement).dataset.index || '0');
        this.previewFormattingRule(rules[index]);
      });
    });

    // Apply all button
    const applyAllBtn = document.getElementById('apply-all-rules-btn');
    if (applyAllBtn) {
      applyAllBtn.onclick = () => this.applySelectedRules();
    }
  }

  /**
   * Auto format selection
   */
  private async autoFormat(): Promise<void> {
    try {
      await this.suggestFormatting();
      
      if (this.currentFormattingRules.length > 0) {
        // Apply top 3 rules automatically
        const topRules = this.currentFormattingRules.slice(0, 3);
        await advancedExcel.applySmartFormatting(topRules);
        this.showSuccess('Auto-formatting applied!');
      }
    } catch (error) {
      console.error('Auto-format failed:', error);
      this.showError('Failed to apply auto-formatting');
    }
  }

  /**
   * Apply selected formatting rules
   */
  private async applySelectedRules(): Promise<void> {
    const selectedRules: SmartFormattingRule[] = [];
    
    this.currentFormattingRules.forEach((rule, index) => {
      const checkbox = document.getElementById(`rule-${index}`) as HTMLInputElement;
      if (checkbox && checkbox.checked) {
        selectedRules.push(rule);
      }
    });

    if (selectedRules.length === 0) {
      this.showError('No rules selected');
      return;
    }

    try {
      await advancedExcel.applySmartFormatting(selectedRules);
      this.showSuccess(`Applied ${selectedRules.length} formatting rules`);
    } catch (error) {
      console.error('Failed to apply rules:', error);
      this.showError('Failed to apply formatting rules');
    }
  }

  /**
   * Load available tables
   */
  private async loadTables(): Promise<void> {
    try {
      await Office.context.sync(async (context) => {
        const tables = context.workbook.tables;
        tables.load('items');
        await context.sync();
        
        const select = document.getElementById('table-select') as HTMLSelectElement;
        if (!select) return;
        
        if (tables.items.length === 0) {
          select.innerHTML = '<option value="">No tables found</option>';
        } else {
          select.innerHTML = tables.items.map(table => 
            `<option value="${table.name}">${table.name}</option>`
          ).join('');
        }
      });
    } catch (error) {
      console.error('Failed to load tables:', error);
    }
  }

  /**
   * Enhance selected table
   */
  private async enhanceTable(): Promise<void> {
    const select = document.getElementById('table-select') as HTMLSelectElement;
    const tableName = select.value;
    
    if (!tableName) {
      this.showError('Please select a table');
      return;
    }

    try {
      this.showTableLoading(true);
      
      const enhancements = await advancedExcel.enhanceTable(tableName);
      this.displayTableEnhancements(tableName, enhancements);
      
      this.showTableLoading(false);
      
    } catch (error) {
      console.error('Table enhancement failed:', error);
      this.showError('Failed to enhance table');
      this.showTableLoading(false);
    }
  }

  /**
   * Display table enhancements
   */
  private displayTableEnhancements(tableName: string, enhancements: TableEnhancement): void {
    const container = document.getElementById('table-enhancements');
    if (!container) return;
    
    // Calculated columns
    const calcColumns = document.getElementById('calculated-columns');
    if (calcColumns) {
      calcColumns.innerHTML = enhancements.calculatedColumns.length > 0 ?
        enhancements.calculatedColumns.map(col => `
          <div class="calc-column">
            <input type="checkbox" id="calc-${col.name}" checked />
            <label for="calc-${col.name}">
              <strong>${col.name}</strong>: ${col.description}
              <code>${col.formula}</code>
            </label>
          </div>
        `).join('') : '<p class="no-data">No calculated columns suggested</p>';
    }

    // Total suggestions
    const totalSuggestions = document.getElementById('total-suggestions');
    if (totalSuggestions) {
      totalSuggestions.innerHTML = enhancements.suggestedTotals.length > 0 ?
        enhancements.suggestedTotals.map(total => `
          <div class="total-item">
            <span class="total-column">${total.column}</span>
            <span class="total-function">${total.function}</span>
          </div>
        `).join('') : '<p class="no-data">No total row suggestions</p>';
    }

    // Slicer suggestions
    const slicerSuggestions = document.getElementById('slicer-suggestions');
    if (slicerSuggestions) {
      slicerSuggestions.innerHTML = enhancements.slicers.length > 0 ?
        enhancements.slicers.map(slicer => `
          <div class="slicer-item">
            <input type="checkbox" id="slicer-${slicer.field}" checked />
            <label for="slicer-${slicer.field}">${slicer.field}</label>
          </div>
        `).join('') : '<p class="no-data">No slicer suggestions</p>';
    }

    container.style.display = 'block';

    // Apply button
    const applyBtn = document.getElementById('apply-enhancements-btn');
    if (applyBtn) {
      applyBtn.onclick = () => this.applyTableEnhancements(tableName, enhancements);
    }
  }

  /**
   * Apply table enhancements
   */
  private async applyTableEnhancements(tableName: string, enhancements: TableEnhancement): Promise<void> {
    try {
      // Filter selected enhancements
      const selectedEnhancements = { ...enhancements };
      
      // Filter calculated columns
      selectedEnhancements.calculatedColumns = enhancements.calculatedColumns.filter(col => {
        const checkbox = document.getElementById(`calc-${col.name}`) as HTMLInputElement;
        return checkbox && checkbox.checked;
      });

      // Filter slicers
      selectedEnhancements.slicers = enhancements.slicers.filter(slicer => {
        const checkbox = document.getElementById(`slicer-${slicer.field}`) as HTMLInputElement;
        return checkbox && checkbox.checked;
      });

      await advancedExcel.applyTableEnhancements(tableName, selectedEnhancements);
      this.showSuccess('Table enhancements applied!');
      
    } catch (error) {
      console.error('Failed to apply enhancements:', error);
      this.showError('Failed to apply table enhancements');
    }
  }

  // Helper methods

  /**
   * Get selected range address
   */
  private async getSelectedRange(): Promise<string | null> {
    try {
      return await Office.context.sync(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address');
        await context.sync();
        return range.address;
      });
    } catch (error) {
      console.error('Failed to get selected range:', error);
      return null;
    }
  }

  /**
   * Get active worksheet
   */
  private async getActiveWorksheet(): Promise<any> {
    return await Office.context.sync(async (context) => {
      return context.workbook.worksheets.getActiveWorksheet();
    });
  }

  /**
   * Show PivotTable instructions
   */
  private showPivotInstructions(instructions: string): void {
    const modal = document.createElement('div');
    modal.className = 'instruction-modal';
    modal.innerHTML = `
      <div class="modal-content">
        <h3>PivotTable Creation Instructions</h3>
        <div class="instructions">
          ${instructions.split('\n').map(line => `<p>${line}</p>`).join('')}
        </div>
        <button class="ms-Button ms-Button--primary close-modal">Close</button>
      </div>
    `;
    
    document.body.appendChild(modal);
    
    modal.querySelector('.close-modal')?.addEventListener('click', () => {
      modal.remove();
    });
  }

  /**
   * Attach graph control listeners
   */
  private attachGraphControls(): void {
    // These would control the D3.js visualization in a real implementation
    const zoomInBtn = document.getElementById('zoom-in-btn');
    if (zoomInBtn) {
      zoomInBtn.onclick = () => console.log('Zoom in');
    }

    const zoomOutBtn = document.getElementById('zoom-out-btn');
    if (zoomOutBtn) {
      zoomOutBtn.onclick = () => console.log('Zoom out');
    }

    const resetBtn = document.getElementById('reset-view-btn');
    if (resetBtn) {
      resetBtn.onclick = () => console.log('Reset view');
    }

    const exportBtn = document.getElementById('export-graph-btn');
    if (exportBtn) {
      exportBtn.onclick = () => this.exportGraph();
    }
  }

  /**
   * Export dependency graph
   */
  private exportGraph(): void {
    if (!this.currentDependencyGraph) return;
    
    const data = JSON.stringify(this.currentDependencyGraph, null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'dependency-graph.json';
    a.click();
    
    URL.revokeObjectURL(url);
  }

  /**
   * Preview formatting rule
   */
  private previewFormattingRule(rule: SmartFormattingRule): void {
    const preview = document.getElementById('format-preview');
    const sample = document.getElementById('format-sample');
    
    if (!preview || !sample) return;
    
    sample.innerHTML = `
      <div class="format-preview-content">
        <h5>Rule Preview: ${this.formatRuleType(rule.type)}</h5>
        <div class="preview-details">
          <p><strong>Type:</strong> ${rule.type}</p>
          <p><strong>Reasoning:</strong> ${rule.reasoning}</p>
          ${rule.condition ? `<p><strong>Condition:</strong> <code>${rule.condition}</code></p>` : ''}
          <p><strong>Priority:</strong> ${rule.priority}</p>
        </div>
        <div class="preview-example">
          ${this.generateRulePreview(rule)}
        </div>
      </div>
    `;
    
    preview.style.display = 'block';
  }

  /**
   * Generate rule preview HTML
   */
  private generateRulePreview(rule: SmartFormattingRule): string {
    switch (rule.type) {
      case 'dataBar':
        return `
          <div class="databar-preview">
            <div class="databar-cell">
              <div class="databar" style="width: 75%; background: #2196F3;"></div>
              <span class="databar-value">75</span>
            </div>
            <div class="databar-cell">
              <div class="databar" style="width: 50%; background: #2196F3;"></div>
              <span class="databar-value">50</span>
            </div>
            <div class="databar-cell">
              <div class="databar" style="width: 25%; background: #2196F3;"></div>
              <span class="databar-value">25</span>
            </div>
          </div>
        `;
      
      case 'colorScale':
        return `
          <div class="colorscale-preview">
            <div class="color-cell" style="background: ${rule.parameters.minColor || '#F44336'};">Low</div>
            <div class="color-cell" style="background: ${rule.parameters.midColor || '#FFEB3B'};">Mid</div>
            <div class="color-cell" style="background: ${rule.parameters.maxColor || '#4CAF50'};">High</div>
          </div>
        `;
      
      case 'iconSet':
        return `
          <div class="iconset-preview">
            <div class="icon-cell">🔴 Low</div>
            <div class="icon-cell">🟡 Medium</div>
            <div class="icon-cell">🟢 High</div>
          </div>
        `;
      
      default:
        return '<p>Preview not available for this rule type</p>';
    }
  }

  /**
   * Format rule type for display
   */
  private formatRuleType(type: string): string {
    const typeMap: Record<string, string> = {
      'conditional': 'Conditional Format',
      'dataBar': 'Data Bars',
      'colorScale': 'Color Scale',
      'iconSet': 'Icon Set',
      'custom': 'Custom Rule'
    };
    return typeMap[type] || type;
  }

  /**
   * Show/hide loading states
   */
  private showPivotLoading(show: boolean): void {
    const loading = document.getElementById('pivot-loading');
    if (loading) loading.style.display = show ? 'block' : 'none';
  }

  private showDependencyLoading(show: boolean): void {
    // Would show loading for dependency mapping
  }

  private showFormattingLoading(show: boolean): void {
    // Would show loading for formatting suggestions
  }

  private showTableLoading(show: boolean): void {
    // Would show loading for table enhancements
  }

  /**
   * Show error message
   */
  private showError(message: string): void {
    const notification = document.createElement('div');
    notification.className = 'error-notification';
    notification.textContent = message;
    document.body.appendChild(notification);
    
    setTimeout(() => notification.remove(), 5000);
  }

  /**
   * Show success message
   */
  private showSuccess(message: string): void {
    const notification = document.createElement('div');
    notification.className = 'success-notification';
    notification.textContent = message;
    document.body.appendChild(notification);
    
    setTimeout(() => notification.remove(), 3000);
  }

  /**
   * Add custom styles
   */
  private addStyles(): void {
    const style = document.createElement('style');
    style.textContent = `
      .advanced-excel-panel {
        height: 100%;
        display: flex;
        flex-direction: column;
      }
      
      .panel-tabs {
        display: flex;
        border-bottom: 2px solid #e0e0e0;
        background: #f5f5f5;
      }
      
      .tab-button {
        flex: 1;
        padding: 10px;
        border: none;
        background: none;
        cursor: pointer;
        font-weight: 500;
        color: #666;
        transition: all 0.3s;
      }
      
      .tab-button:hover {
        background: #e8e8e8;
      }
      
      .tab-button.active {
        color: #0078d4;
        border-bottom: 2px solid #0078d4;
        margin-bottom: -2px;
        background: white;
      }
      
      .tab-content {
        flex: 1;
        overflow-y: auto;
        padding: 15px;
      }
      
      .tab-panel {
        display: none;
      }
      
      .tab-panel.active {
        display: block;
      }
      
      .description {
        color: #666;
        margin-bottom: 15px;
      }
      
      .action-group {
        margin-bottom: 20px;
      }
      
      .action-group label {
        display: block;
        margin-bottom: 5px;
        font-weight: 600;
      }
      
      .recommendations-container,
      .preview-container,
      .stats-container,
      .visualization-container,
      .path-container,
      .rules-container,
      .enhancements-container {
        background: #f5f5f5;
        padding: 15px;
        border-radius: 4px;
        margin-top: 15px;
      }
      
      .recommendation-card,
      .rule-card {
        background: white;
        padding: 15px;
        margin-bottom: 10px;
        border-radius: 4px;
        border: 1px solid #e0e0e0;
      }
      
      .rec-header,
      .rule-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
      }
      
      .confidence-score {
        background: #4CAF50;
        color: white;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 12px;
      }
      
      .rec-reasoning,
      .rule-reasoning {
        color: #666;
        margin-bottom: 10px;
      }
      
      .rec-config,
      .config-item {
        font-size: 13px;
        margin-bottom: 5px;
      }
      
      .rec-insights ul {
        margin: 5px 0;
        padding-left: 20px;
      }
      
      .stats-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 15px;
        text-align: center;
      }
      
      .stat-item {
        background: white;
        padding: 15px;
        border-radius: 4px;
      }
      
      .stat-value {
        display: block;
        font-size: 24px;
        font-weight: bold;
        color: #0078d4;
      }
      
      .stat-label {
        display: block;
        font-size: 12px;
        color: #666;
        margin-top: 5px;
      }
      
      .graph-placeholder {
        background: white;
        padding: 40px;
        text-align: center;
        border: 2px dashed #d0d0d0;
        border-radius: 4px;
      }
      
      .graph-controls {
        display: flex;
        gap: 10px;
        justify-content: center;
        margin-top: 10px;
      }
      
      .control-btn {
        padding: 5px 10px;
        border: 1px solid #d0d0d0;
        background: white;
        border-radius: 4px;
        cursor: pointer;
      }
      
      .control-btn:hover {
        background: #f5f5f5;
      }
      
      .path-flow {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 10px;
      }
      
      .path-node {
        background: white;
        padding: 8px 12px;
        border-radius: 4px;
        border: 1px solid #d0d0d0;
      }
      
      .node-address {
        font-weight: 600;
        color: #0078d4;
      }
      
      .node-formula {
        display: block;
        font-size: 11px;
        color: #666;
        margin-top: 3px;
      }
      
      .path-arrow {
        font-size: 20px;
        color: #666;
      }
      
      .ai-badge {
        background: #2196F3;
        color: white;
        padding: 2px 6px;
        border-radius: 10px;
        font-size: 10px;
        margin-left: 8px;
      }
      
      .preview-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 15px;
        margin-bottom: 15px;
      }
      
      .preview-section {
        background: #f5f5f5;
        padding: 10px;
        border-radius: 4px;
      }
      
      .preview-section h6 {
        margin: 0 0 10px 0;
        color: #666;
      }
      
      .field-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 5px;
        background: white;
        margin-bottom: 5px;
        border-radius: 2px;
      }
      
      .field-name {
        font-weight: 500;
      }
      
      .field-operation,
      .field-sort {
        font-size: 11px;
        color: #666;
      }
      
      .empty {
        color: #999;
        font-style: italic;
      }
      
      .instruction-modal {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.5);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 2000;
      }
      
      .modal-content {
        background: white;
        padding: 20px;
        border-radius: 8px;
        max-width: 600px;
        width: 90%;
        max-height: 80vh;
        overflow-y: auto;
      }
      
      .instructions {
        background: #f5f5f5;
        padding: 15px;
        border-radius: 4px;
        font-family: 'Courier New', monospace;
        margin: 15px 0;
      }
      
      .loading-container {
        text-align: center;
        padding: 40px;
      }
      
      .spinner {
        border: 3px solid #f3f3f3;
        border-top: 3px solid #0078d4;
        border-radius: 50%;
        width: 40px;
        height: 40px;
        animation: spin 1s linear infinite;
        margin: 0 auto 20px;
      }
      
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
      
      .no-data {
        color: #999;
        text-align: center;
        padding: 20px;
      }
      
      .success-notification,
      .error-notification {
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 4px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        z-index: 2000;
        animation: slideIn 0.3s ease-out;
      }
      
      .success-notification {
        background: #4CAF50;
        color: white;
      }
      
      .error-notification {
        background: #f44336;
        color: white;
      }
      
      .databar-preview,
      .colorscale-preview,
      .iconset-preview {
        display: flex;
        gap: 10px;
      }
      
      .databar-cell {
        position: relative;
        background: white;
        border: 1px solid #d0d0d0;
        padding: 5px;
        width: 100px;
      }
      
      .databar {
        height: 20px;
        background: #2196F3;
        position: absolute;
        left: 0;
        top: 0;
      }
      
      .databar-value {
        position: relative;
        z-index: 1;
      }
      
      .color-cell,
      .icon-cell {
        padding: 10px 20px;
        border-radius: 4px;
        text-align: center;
      }
      
      .calc-column,
      .total-item,
      .slicer-item {
        margin-bottom: 10px;
      }
      
      .calc-column code {
        display: block;
        margin-top: 5px;
        font-size: 12px;
        background: #f5f5f5;
        padding: 5px;
        border-radius: 2px;
      }
      
      .total-column {
        font-weight: 600;
      }
      
      .total-function {
        color: #666;
        margin-left: 10px;
      }
      
      .rule-condition {
        display: block;
        margin-top: 5px;
        font-size: 12px;
        background: #f5f5f5;
        padding: 5px;
        border-radius: 2px;
      }
      
      .preview-rule-btn {
        background: none;
        border: 1px solid #0078d4;
        color: #0078d4;
        padding: 4px 12px;
        border-radius: 2px;
        cursor: pointer;
        font-size: 12px;
        margin-top: 8px;
      }
      
      .preview-rule-btn:hover {
        background: #0078d4;
        color: white;
      }
    `;
    document.head.appendChild(style);
  }
}

// Export for use in main taskpane
export function initializeAdvancedExcelPanel(containerId: string): AdvancedExcelPanel {
  return new AdvancedExcelPanel(containerId);
}