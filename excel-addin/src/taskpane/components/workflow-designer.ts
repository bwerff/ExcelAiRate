/**
 * Workflow Designer Component
 * Visual interface for creating and managing workflows
 */

import { 
  workflowEngine, 
  Workflow, 
  WorkflowStep, 
  StepType,
  WORKFLOW_TEMPLATES,
  BatchOperation,
  BatchProgress,
  BatchError
} from '../../services/workflow-engine';

/* global document */

interface StepTemplate {
  type: StepType;
  name: string;
  icon: string;
  description: string;
  defaultOperation: string;
}

const STEP_TEMPLATES: StepTemplate[] = [
  {
    type: 'ai-analysis',
    name: 'AI Analysis',
    icon: '🤖',
    description: 'Analyze data with AI',
    defaultOperation: 'analyze'
  },
  {
    type: 'ai-generation',
    name: 'AI Generate',
    icon: '✨',
    description: 'Generate content with AI',
    defaultOperation: 'generate'
  },
  {
    type: 'data-transform',
    name: 'Transform Data',
    icon: '🔄',
    description: 'Transform or filter data',
    defaultOperation: 'filter'
  },
  {
    type: 'excel-operation',
    name: 'Excel Operation',
    icon: '📊',
    description: 'Excel-specific operations',
    defaultOperation: 'create-chart'
  },
  {
    type: 'conditional',
    name: 'Conditional',
    icon: '🔀',
    description: 'If/then logic',
    defaultOperation: '$variable > 0'
  },
  {
    type: 'loop',
    name: 'Loop',
    icon: '🔁',
    description: 'Repeat for each item',
    defaultOperation: 'forEach'
  },
  {
    type: 'validation',
    name: 'Validate',
    icon: '✅',
    description: 'Validate data',
    defaultOperation: 'validate-required'
  },
  {
    type: 'formatting',
    name: 'Format',
    icon: '🎨',
    description: 'Apply formatting',
    defaultOperation: 'conditional-format'
  }
];

export class WorkflowDesigner {
  private container: HTMLElement;
  private currentWorkflow: Workflow;
  private selectedStep: WorkflowStep | null = null;
  private isExecuting: boolean = false;
  private draggedStep: StepTemplate | null = null;

  constructor(containerId: string) {
    this.container = document.getElementById(containerId) || document.createElement('div');
    this.currentWorkflow = this.createEmptyWorkflow();
    this.initialize();
  }

  /**
   * Initialize the designer
   */
  private initialize(): void {
    this.container.innerHTML = `
      <div class="workflow-designer">
        <div class="designer-header">
          <h3>Workflow Designer</h3>
          <div class="header-actions">
            <button id="new-workflow-btn" class="ms-Button ms-Button--small">
              <span class="ms-Button-label">New</span>
            </button>
            <button id="load-template-btn" class="ms-Button ms-Button--small">
              <span class="ms-Button-label">Templates</span>
            </button>
            <button id="save-workflow-btn" class="ms-Button ms-Button--small">
              <span class="ms-Button-label">Save</span>
            </button>
          </div>
        </div>
        
        <div class="designer-body">
          <div class="step-palette">
            <h4>Steps</h4>
            <div id="step-templates"></div>
          </div>
          
          <div class="workflow-canvas">
            <div class="workflow-info">
              <input type="text" id="workflow-name" class="ms-TextField-field" placeholder="Workflow Name" />
              <textarea id="workflow-description" class="ms-TextField-field" placeholder="Description" rows="2"></textarea>
            </div>
            
            <div id="workflow-steps" class="workflow-steps">
              <div class="drop-zone" data-position="0">Drop step here</div>
            </div>
            
            <div class="workflow-actions">
              <button id="execute-workflow-btn" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Execute Workflow</span>
              </button>
              <button id="execute-batch-btn" class="ms-Button">
                <span class="ms-Button-label">Batch Execute</span>
              </button>
            </div>
          </div>
          
          <div class="step-properties">
            <h4>Step Properties</h4>
            <div id="step-properties-content">
              <p class="no-selection">Select a step to edit properties</p>
            </div>
          </div>
        </div>
        
        <div id="execution-results" class="execution-results" style="display: none;">
          <h4>Execution Results</h4>
          <div id="results-content"></div>
        </div>
        
        <div id="template-modal" class="modal" style="display: none;">
          <div class="modal-content">
            <h3>Workflow Templates</h3>
            <div id="template-list"></div>
            <button class="ms-Button close-modal">Close</button>
          </div>
        </div>
        
        <div id="batch-modal" class="modal" style="display: none;">
          <div class="modal-content">
            <h3>Batch Execution</h3>
            <div class="batch-config">
              <label>Ranges (one per line):</label>
              <textarea id="batch-ranges" class="ms-TextField-field" rows="4" placeholder="A1:A10\nB1:B10"></textarea>
              
              <label>Sheets (one per line):</label>
              <textarea id="batch-sheets" class="ms-TextField-field" rows="4" placeholder="Sheet1\nSheet2"></textarea>
              
              <label>
                <input type="checkbox" id="batch-parallel" />
                Execute in parallel
              </label>
            </div>
            <div class="modal-actions">
              <button id="start-batch-btn" class="ms-Button ms-Button--primary">Start Batch</button>
              <button class="ms-Button close-modal">Cancel</button>
            </div>
            <div id="batch-progress" style="display: none;">
              <div class="progress-bar">
                <div class="progress-fill" style="width: 0%"></div>
              </div>
              <p class="progress-text">0 / 0 completed</p>
            </div>
          </div>
        </div>
      </div>
    `;

    this.renderStepTemplates();
    this.renderWorkflow();
    this.attachEventListeners();
    this.addStyles();
  }

  /**
   * Create empty workflow
   */
  private createEmptyWorkflow(): Workflow {
    return {
      id: `workflow_${Date.now()}`,
      name: 'New Workflow',
      description: '',
      steps: [],
      variables: {},
      errorHandling: 'stop',
      created: new Date(),
      lastModified: new Date()
    };
  }

  /**
   * Render step templates
   */
  private renderStepTemplates(): void {
    const container = document.getElementById('step-templates');
    if (!container) return;

    container.innerHTML = STEP_TEMPLATES.map(template => `
      <div class="step-template" draggable="true" data-type="${template.type}">
        <span class="step-icon">${template.icon}</span>
        <span class="step-name">${template.name}</span>
      </div>
    `).join('');

    // Add drag event listeners
    container.querySelectorAll('.step-template').forEach(el => {
      el.addEventListener('dragstart', (e) => this.handleDragStart(e as DragEvent));
      el.addEventListener('dragend', () => this.handleDragEnd());
    });
  }

  /**
   * Render workflow steps
   */
  private renderWorkflow(): void {
    const container = document.getElementById('workflow-steps');
    if (!container) return;

    // Update workflow info
    const nameInput = document.getElementById('workflow-name') as HTMLInputElement;
    const descInput = document.getElementById('workflow-description') as HTMLTextAreaElement;
    
    if (nameInput) nameInput.value = this.currentWorkflow.name;
    if (descInput) descInput.value = this.currentWorkflow.description;

    // Render steps
    const stepsHtml = this.currentWorkflow.steps.map((step, index) => {
      const template = STEP_TEMPLATES.find(t => t.type === step.type);
      return `
        <div class="workflow-step ${this.selectedStep?.id === step.id ? 'selected' : ''}" 
             data-step-id="${step.id}" data-index="${index}">
          <div class="step-header">
            <span class="step-icon">${template?.icon || '❓'}</span>
            <span class="step-name">${step.name}</span>
            <button class="remove-step-btn" data-step-id="${step.id}">✕</button>
          </div>
          <div class="step-details">
            <span class="step-type">${step.type}</span>
            <span class="step-operation">${step.operation}</span>
          </div>
        </div>
        <div class="drop-zone" data-position="${index + 1}">Drop step here</div>
      `;
    }).join('');

    container.innerHTML = `
      <div class="drop-zone" data-position="0">Drop step here</div>
      ${stepsHtml}
    `;

    // Add event listeners
    container.querySelectorAll('.workflow-step').forEach(el => {
      el.addEventListener('click', (e) => {
        const stepId = (e.currentTarget as HTMLElement).dataset.stepId;
        if (stepId) this.selectStep(stepId);
      });
    });

    container.querySelectorAll('.remove-step-btn').forEach(el => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const stepId = (e.currentTarget as HTMLElement).dataset.stepId;
        if (stepId) this.removeStep(stepId);
      });
    });

    // Add drop zone listeners
    container.querySelectorAll('.drop-zone').forEach(el => {
      el.addEventListener('dragover', (e) => this.handleDragOver(e as DragEvent));
      el.addEventListener('drop', (e) => this.handleDrop(e as DragEvent));
      el.addEventListener('dragleave', (e) => this.handleDragLeave(e as DragEvent));
    });
  }

  /**
   * Select a step for editing
   */
  private selectStep(stepId: string): void {
    this.selectedStep = this.currentWorkflow.steps.find(s => s.id === stepId) || null;
    this.renderWorkflow();
    this.renderStepProperties();
  }

  /**
   * Render step properties panel
   */
  private renderStepProperties(): void {
    const container = document.getElementById('step-properties-content');
    if (!container) return;

    if (!this.selectedStep) {
      container.innerHTML = '<p class="no-selection">Select a step to edit properties</p>';
      return;
    }

    container.innerHTML = `
      <div class="property-group">
        <label>Step Name:</label>
        <input type="text" id="step-name-input" class="ms-TextField-field" value="${this.selectedStep.name}" />
      </div>
      
      <div class="property-group">
        <label>Operation:</label>
        <input type="text" id="step-operation-input" class="ms-TextField-field" value="${this.selectedStep.operation}" />
      </div>
      
      <div class="property-group">
        <label>Inputs:</label>
        <div id="step-inputs">
          ${this.renderStepInputs()}
        </div>
        <button id="add-input-btn" class="ms-Button ms-Button--small">Add Input</button>
      </div>
      
      <div class="property-group">
        <label>Outputs:</label>
        <div id="step-outputs">
          ${this.renderStepOutputs()}
        </div>
        <button id="add-output-btn" class="ms-Button ms-Button--small">Add Output</button>
      </div>
      
      <div class="property-group">
        <label>Error Handling:</label>
        <select id="step-error-handling" class="ms-Dropdown">
          <option value="default">Use Workflow Default</option>
          <option value="retry" ${this.selectedStep.retryPolicy ? 'selected' : ''}>Retry on Error</option>
        </select>
      </div>
      
      ${this.selectedStep.retryPolicy ? `
        <div class="property-group">
          <label>Max Retries:</label>
          <input type="number" id="max-retries" class="ms-TextField-field" 
                 value="${this.selectedStep.retryPolicy.maxAttempts}" min="1" max="5" />
        </div>
      ` : ''}
      
      <button id="update-step-btn" class="ms-Button ms-Button--primary">Update Step</button>
    `;

    // Attach property event listeners
    this.attachPropertyListeners();
  }

  /**
   * Render step inputs
   */
  private renderStepInputs(): string {
    if (!this.selectedStep) return '';
    
    return this.selectedStep.inputs.map((input, index) => `
      <div class="input-item">
        <input type="text" placeholder="Name" value="${input.name}" 
               data-index="${index}" data-field="name" class="input-field" />
        <select data-index="${index}" data-field="type" class="input-field">
          <option value="range" ${input.type === 'range' ? 'selected' : ''}>Range</option>
          <option value="value" ${input.type === 'value' ? 'selected' : ''}>Value</option>
          <option value="variable" ${input.type === 'variable' ? 'selected' : ''}>Variable</option>
          <option value="previous-output" ${input.type === 'previous-output' ? 'selected' : ''}>Previous Output</option>
        </select>
        <input type="text" placeholder="Source" value="${input.source}" 
               data-index="${index}" data-field="source" class="input-field" />
        <button class="remove-input-btn" data-index="${index}">✕</button>
      </div>
    `).join('');
  }

  /**
   * Render step outputs
   */
  private renderStepOutputs(): string {
    if (!this.selectedStep) return '';
    
    return this.selectedStep.outputs.map((output, index) => `
      <div class="output-item">
        <input type="text" placeholder="Name" value="${output.name}" 
               data-index="${index}" data-field="name" class="output-field" />
        <select data-index="${index}" data-field="type" class="output-field">
          <option value="range" ${output.type === 'range' ? 'selected' : ''}>Range</option>
          <option value="value" ${output.type === 'value' ? 'selected' : ''}>Value</option>
          <option value="variable" ${output.type === 'variable' ? 'selected' : ''}>Variable</option>
        </select>
        <input type="text" placeholder="Target" value="${output.target}" 
               data-index="${index}" data-field="target" class="output-field" />
        <button class="remove-output-btn" data-index="${index}">✕</button>
      </div>
    `).join('');
  }

  /**
   * Attach property panel event listeners
   */
  private attachPropertyListeners(): void {
    // Update step button
    const updateBtn = document.getElementById('update-step-btn');
    if (updateBtn) {
      updateBtn.onclick = () => this.updateSelectedStep();
    }

    // Add input/output buttons
    const addInputBtn = document.getElementById('add-input-btn');
    if (addInputBtn) {
      addInputBtn.onclick = () => this.addStepInput();
    }

    const addOutputBtn = document.getElementById('add-output-btn');
    if (addOutputBtn) {
      addOutputBtn.onclick = () => this.addStepOutput();
    }

    // Input/output field changes
    document.querySelectorAll('.input-field').forEach(el => {
      el.addEventListener('change', (e) => this.updateInputField(e as Event));
    });

    document.querySelectorAll('.output-field').forEach(el => {
      el.addEventListener('change', (e) => this.updateOutputField(e as Event));
    });

    // Remove buttons
    document.querySelectorAll('.remove-input-btn').forEach(el => {
      el.addEventListener('click', (e) => {
        const index = parseInt((e.target as HTMLElement).dataset.index || '0');
        this.removeStepInput(index);
      });
    });

    document.querySelectorAll('.remove-output-btn').forEach(el => {
      el.addEventListener('click', (e) => {
        const index = parseInt((e.target as HTMLElement).dataset.index || '0');
        this.removeStepOutput(index);
      });
    });
  }

  /**
   * Update selected step
   */
  private updateSelectedStep(): void {
    if (!this.selectedStep) return;

    const nameInput = document.getElementById('step-name-input') as HTMLInputElement;
    const operationInput = document.getElementById('step-operation-input') as HTMLInputElement;
    const errorHandling = document.getElementById('step-error-handling') as HTMLSelectElement;

    if (nameInput) this.selectedStep.name = nameInput.value;
    if (operationInput) this.selectedStep.operation = operationInput.value;
    
    if (errorHandling && errorHandling.value === 'retry') {
      this.selectedStep.retryPolicy = {
        maxAttempts: 3,
        delayMs: 1000,
        backoffMultiplier: 2
      };
    } else {
      this.selectedStep.retryPolicy = undefined;
    }

    this.currentWorkflow.lastModified = new Date();
    this.renderWorkflow();
    this.showSuccess('Step updated');
  }

  /**
   * Add step to workflow
   */
  private addStep(type: StepType, position: number): void {
    const template = STEP_TEMPLATES.find(t => t.type === type);
    if (!template) return;

    const newStep: WorkflowStep = {
      id: `step_${Date.now()}`,
      name: template.name,
      type: type,
      operation: template.defaultOperation,
      inputs: [],
      outputs: [],
      conditions: []
    };

    this.currentWorkflow.steps.splice(position, 0, newStep);
    this.currentWorkflow.lastModified = new Date();
    this.renderWorkflow();
    this.selectStep(newStep.id);
  }

  /**
   * Remove step from workflow
   */
  private removeStep(stepId: string): void {
    const index = this.currentWorkflow.steps.findIndex(s => s.id === stepId);
    if (index > -1) {
      this.currentWorkflow.steps.splice(index, 1);
      this.currentWorkflow.lastModified = new Date();
      this.selectedStep = null;
      this.renderWorkflow();
      this.renderStepProperties();
    }
  }

  /**
   * Add input to selected step
   */
  private addStepInput(): void {
    if (!this.selectedStep) return;
    
    this.selectedStep.inputs.push({
      name: '',
      type: 'value',
      source: '',
      required: true
    });
    
    this.renderStepProperties();
  }

  /**
   * Add output to selected step
   */
  private addStepOutput(): void {
    if (!this.selectedStep) return;
    
    this.selectedStep.outputs.push({
      name: '',
      type: 'variable',
      target: ''
    });
    
    this.renderStepProperties();
  }

  /**
   * Update input field
   */
  private updateInputField(event: Event): void {
    if (!this.selectedStep) return;
    
    const target = event.target as HTMLElement;
    const index = parseInt(target.dataset.index || '0');
    const field = target.dataset.field as keyof typeof this.selectedStep.inputs[0];
    const value = (target as HTMLInputElement).value;
    
    if (this.selectedStep.inputs[index]) {
      (this.selectedStep.inputs[index] as any)[field] = value;
    }
  }

  /**
   * Update output field
   */
  private updateOutputField(event: Event): void {
    if (!this.selectedStep) return;
    
    const target = event.target as HTMLElement;
    const index = parseInt(target.dataset.index || '0');
    const field = target.dataset.field as keyof typeof this.selectedStep.outputs[0];
    const value = (target as HTMLInputElement).value;
    
    if (this.selectedStep.outputs[index]) {
      (this.selectedStep.outputs[index] as any)[field] = value;
    }
  }

  /**
   * Remove input from step
   */
  private removeStepInput(index: number): void {
    if (!this.selectedStep) return;
    this.selectedStep.inputs.splice(index, 1);
    this.renderStepProperties();
  }

  /**
   * Remove output from step
   */
  private removeStepOutput(index: number): void {
    if (!this.selectedStep) return;
    this.selectedStep.outputs.splice(index, 1);
    this.renderStepProperties();
  }

  /**
   * Execute workflow
   */
  private async executeWorkflow(): Promise<void> {
    if (this.isExecuting) return;
    
    try {
      this.isExecuting = true;
      this.showExecutionStatus('Executing workflow...');
      
      // Update workflow name and description
      const nameInput = document.getElementById('workflow-name') as HTMLInputElement;
      const descInput = document.getElementById('workflow-description') as HTMLTextAreaElement;
      
      if (nameInput) this.currentWorkflow.name = nameInput.value;
      if (descInput) this.currentWorkflow.description = descInput.value;
      
      const result = await workflowEngine.executeWorkflow(this.currentWorkflow);
      this.showExecutionResults(result);
      
    } catch (error) {
      this.showError('Workflow execution failed: ' + (error as Error).message);
    } finally {
      this.isExecuting = false;
    }
  }

  /**
   * Show batch execution modal
   */
  private showBatchModal(): void {
    const modal = document.getElementById('batch-modal');
    if (modal) {
      modal.style.display = 'flex';
    }
  }

  /**
   * Execute batch operation
   */
  private async executeBatch(): Promise<void> {
    const rangesInput = document.getElementById('batch-ranges') as HTMLTextAreaElement;
    const sheetsInput = document.getElementById('batch-sheets') as HTMLTextAreaElement;
    const parallelInput = document.getElementById('batch-parallel') as HTMLInputElement;
    
    const ranges = rangesInput.value.split('\n').filter(r => r.trim());
    const sheets = sheetsInput.value.split('\n').filter(s => s.trim());
    const parallel = parallelInput.checked;
    
    if (ranges.length === 0 && sheets.length === 0) {
      this.showError('Please specify ranges or sheets');
      return;
    }
    
    const progressDiv = document.getElementById('batch-progress');
    if (progressDiv) progressDiv.style.display = 'block';
    
    const batch: BatchOperation = {
      ranges,
      sheets,
      operation: this.currentWorkflow,
      parallelExecution: parallel,
      progressCallback: (progress: BatchProgress) => {
        this.updateBatchProgress(progress);
      },
      errorCallback: (error: BatchError) => {
        console.error(`Batch error for ${error.item}:`, error.error);
      }
    };
    
    try {
      const results = await workflowEngine.executeBatch(batch);
      this.showBatchResults(results);
    } catch (error) {
      this.showError('Batch execution failed: ' + (error as Error).message);
    }
  }

  /**
   * Update batch progress
   */
  private updateBatchProgress(progress: BatchProgress): void {
    const fill = document.querySelector('.progress-fill') as HTMLElement;
    const text = document.querySelector('.progress-text') as HTMLElement;
    
    if (fill) fill.style.width = `${progress.percentage}%`;
    if (text) text.textContent = `${progress.completed} / ${progress.total} completed`;
  }

  /**
   * Show batch results
   */
  private showBatchResults(results: any[]): void {
    const modal = document.getElementById('batch-modal');
    if (modal) modal.style.display = 'none';
    
    const successful = results.filter(r => r.success).length;
    const failed = results.length - successful;
    
    this.showExecutionStatus(
      `Batch complete: ${successful} successful, ${failed} failed`
    );
  }

  /**
   * Show execution results
   */
  private showExecutionResults(result: any): void {
    const resultsDiv = document.getElementById('execution-results');
    const content = document.getElementById('results-content');
    
    if (!resultsDiv || !content) return;
    
    resultsDiv.style.display = 'block';
    
    content.innerHTML = `
      <div class="result-summary ${result.success ? 'success' : 'error'}">
        <h5>Execution ${result.success ? 'Successful' : 'Failed'}</h5>
        <p>Duration: ${result.duration}ms</p>
      </div>
      
      ${result.errors.length > 0 ? `
        <div class="result-errors">
          <h5>Errors:</h5>
          ${result.errors.map((err: any) => `
            <div class="error-item">
              <strong>${err.stepId}:</strong> ${err.message}
            </div>
          `).join('')}
        </div>
      ` : ''}
      
      <div class="result-outputs">
        <h5>Outputs:</h5>
        ${Object.entries(result.outputs).map(([key, value]) => `
          <div class="output-item">
            <strong>${key}:</strong> ${JSON.stringify(value, null, 2)}
          </div>
        `).join('')}
      </div>
      
      <div class="step-results">
        <h5>Step Results:</h5>
        ${result.steps.map((step: any) => `
          <div class="step-result ${step.success ? 'success' : 'error'}">
            <strong>${step.stepId}:</strong> 
            ${step.success ? '✓' : '✗'} 
            (${step.duration}ms)
            ${step.retries > 0 ? `- ${step.retries} retries` : ''}
          </div>
        `).join('')}
      </div>
    `;
  }

  /**
   * Load workflow templates
   */
  private loadTemplates(): void {
    const modal = document.getElementById('template-modal');
    const list = document.getElementById('template-list');
    
    if (!modal || !list) return;
    
    const templates = workflowEngine.getWorkflowTemplates();
    
    list.innerHTML = templates.map((template, index) => `
      <div class="template-item" data-index="${index}">
        <h4>${template.name}</h4>
        <p>${template.description}</p>
        <button class="ms-Button ms-Button--small use-template-btn" data-index="${index}">
          Use This Template
        </button>
      </div>
    `).join('');
    
    list.querySelectorAll('.use-template-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const index = parseInt((e.target as HTMLElement).dataset.index || '0');
        this.useTemplate(templates[index]);
      });
    });
    
    modal.style.display = 'flex';
  }

  /**
   * Use a workflow template
   */
  private useTemplate(template: Workflow): void {
    this.currentWorkflow = { ...template, id: `workflow_${Date.now()}` };
    this.selectedStep = null;
    this.renderWorkflow();
    this.renderStepProperties();
    
    const modal = document.getElementById('template-modal');
    if (modal) modal.style.display = 'none';
    
    this.showSuccess('Template loaded');
  }

  /**
   * Save workflow as template
   */
  private async saveWorkflow(): Promise<void> {
    try {
      await workflowEngine.saveWorkflowTemplate(this.currentWorkflow);
      this.showSuccess('Workflow saved as template');
    } catch (error) {
      this.showError('Failed to save workflow');
    }
  }

  /**
   * Drag and drop handlers
   */
  private handleDragStart(e: DragEvent): void {
    const target = e.target as HTMLElement;
    const type = target.dataset.type;
    
    if (type) {
      this.draggedStep = STEP_TEMPLATES.find(t => t.type === type) || null;
      target.classList.add('dragging');
      e.dataTransfer!.effectAllowed = 'copy';
    }
  }

  private handleDragEnd(): void {
    document.querySelectorAll('.dragging').forEach(el => {
      el.classList.remove('dragging');
    });
    document.querySelectorAll('.drag-over').forEach(el => {
      el.classList.remove('drag-over');
    });
    this.draggedStep = null;
  }

  private handleDragOver(e: DragEvent): void {
    e.preventDefault();
    e.dataTransfer!.dropEffect = 'copy';
    (e.target as HTMLElement).classList.add('drag-over');
  }

  private handleDragLeave(e: DragEvent): void {
    (e.target as HTMLElement).classList.remove('drag-over');
  }

  private handleDrop(e: DragEvent): void {
    e.preventDefault();
    const target = e.target as HTMLElement;
    target.classList.remove('drag-over');
    
    if (this.draggedStep) {
      const position = parseInt(target.dataset.position || '0');
      this.addStep(this.draggedStep.type, position);
    }
  }

  /**
   * Attach main event listeners
   */
  private attachEventListeners(): void {
    // Header buttons
    const newBtn = document.getElementById('new-workflow-btn');
    if (newBtn) {
      newBtn.onclick = () => {
        this.currentWorkflow = this.createEmptyWorkflow();
        this.selectedStep = null;
        this.renderWorkflow();
        this.renderStepProperties();
      };
    }

    const loadBtn = document.getElementById('load-template-btn');
    if (loadBtn) {
      loadBtn.onclick = () => this.loadTemplates();
    }

    const saveBtn = document.getElementById('save-workflow-btn');
    if (saveBtn) {
      saveBtn.onclick = () => this.saveWorkflow();
    }

    // Execute buttons
    const executeBtn = document.getElementById('execute-workflow-btn');
    if (executeBtn) {
      executeBtn.onclick = () => this.executeWorkflow();
    }

    const batchBtn = document.getElementById('execute-batch-btn');
    if (batchBtn) {
      batchBtn.onclick = () => this.showBatchModal();
    }

    // Batch modal
    const startBatchBtn = document.getElementById('start-batch-btn');
    if (startBatchBtn) {
      startBatchBtn.onclick = () => this.executeBatch();
    }

    // Close modals
    document.querySelectorAll('.close-modal').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const modal = (e.target as HTMLElement).closest('.modal');
        if (modal) modal.style.display = 'none';
      });
    });
  }

  /**
   * Show status messages
   */
  private showExecutionStatus(message: string): void {
    const resultsDiv = document.getElementById('execution-results');
    const content = document.getElementById('results-content');
    
    if (resultsDiv && content) {
      resultsDiv.style.display = 'block';
      content.innerHTML = `<div class="status-message">${message}</div>`;
    }
  }

  private showSuccess(message: string): void {
    const notification = document.createElement('div');
    notification.className = 'success-notification';
    notification.textContent = message;
    document.body.appendChild(notification);
    
    setTimeout(() => notification.remove(), 3000);
  }

  private showError(message: string): void {
    const notification = document.createElement('div');
    notification.className = 'error-notification';
    notification.textContent = message;
    document.body.appendChild(notification);
    
    setTimeout(() => notification.remove(), 5000);
  }

  /**
   * Add custom styles
   */
  private addStyles(): void {
    const style = document.createElement('style');
    style.textContent = `
      .workflow-designer {
        height: 100%;
        display: flex;
        flex-direction: column;
      }
      
      .designer-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 10px;
        border-bottom: 1px solid #e0e0e0;
      }
      
      .header-actions {
        display: flex;
        gap: 8px;
      }
      
      .designer-body {
        flex: 1;
        display: grid;
        grid-template-columns: 200px 1fr 250px;
        gap: 10px;
        padding: 10px;
        overflow: hidden;
      }
      
      .step-palette {
        background: #f5f5f5;
        padding: 10px;
        border-radius: 4px;
        overflow-y: auto;
      }
      
      .step-template {
        background: white;
        padding: 8px;
        margin-bottom: 8px;
        border-radius: 4px;
        cursor: move;
        display: flex;
        align-items: center;
        gap: 8px;
        transition: transform 0.2s;
      }
      
      .step-template:hover {
        transform: translateX(5px);
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      }
      
      .step-template.dragging {
        opacity: 0.5;
      }
      
      .step-icon {
        font-size: 20px;
      }
      
      .workflow-canvas {
        background: white;
        padding: 15px;
        border-radius: 4px;
        border: 1px solid #e0e0e0;
        overflow-y: auto;
      }
      
      .workflow-info {
        margin-bottom: 20px;
      }
      
      .workflow-info input,
      .workflow-info textarea {
        width: 100%;
        margin-bottom: 10px;
      }
      
      .workflow-steps {
        min-height: 200px;
        margin-bottom: 20px;
      }
      
      .drop-zone {
        border: 2px dashed #d0d0d0;
        padding: 20px;
        text-align: center;
        color: #999;
        margin: 10px 0;
        border-radius: 4px;
        transition: all 0.2s;
      }
      
      .drop-zone.drag-over {
        border-color: #0078d4;
        background: #f0f9ff;
        color: #0078d4;
      }
      
      .workflow-step {
        background: #f8f8f8;
        padding: 12px;
        margin: 10px 0;
        border-radius: 4px;
        border: 2px solid transparent;
        cursor: pointer;
        transition: all 0.2s;
      }
      
      .workflow-step:hover {
        border-color: #0078d4;
      }
      
      .workflow-step.selected {
        border-color: #0078d4;
        background: #e3f2fd;
      }
      
      .step-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 5px;
      }
      
      .step-header .step-name {
        flex: 1;
        font-weight: 600;
      }
      
      .remove-step-btn {
        background: none;
        border: none;
        color: #999;
        cursor: pointer;
        padding: 4px;
      }
      
      .remove-step-btn:hover {
        color: #d32f2f;
      }
      
      .step-details {
        display: flex;
        gap: 10px;
        font-size: 12px;
        color: #666;
      }
      
      .step-properties {
        background: #f5f5f5;
        padding: 15px;
        border-radius: 4px;
        overflow-y: auto;
      }
      
      .property-group {
        margin-bottom: 15px;
      }
      
      .property-group label {
        display: block;
        font-weight: 600;
        margin-bottom: 5px;
      }
      
      .property-group input,
      .property-group select {
        width: 100%;
      }
      
      .input-item,
      .output-item {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr auto;
        gap: 5px;
        margin-bottom: 5px;
      }
      
      .execution-results {
        background: #f5f5f5;
        padding: 15px;
        margin-top: 10px;
        border-radius: 4px;
        max-height: 300px;
        overflow-y: auto;
      }
      
      .result-summary {
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 10px;
      }
      
      .result-summary.success {
        background: #e8f5e9;
        color: #2e7d32;
      }
      
      .result-summary.error {
        background: #ffebee;
        color: #c62828;
      }
      
      .modal {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.5);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 1000;
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
      
      .template-item {
        background: #f5f5f5;
        padding: 15px;
        margin-bottom: 10px;
        border-radius: 4px;
      }
      
      .template-item h4 {
        margin: 0 0 5px 0;
      }
      
      .batch-config label {
        display: block;
        margin-top: 10px;
        font-weight: 600;
      }
      
      .batch-config textarea {
        width: 100%;
        margin-top: 5px;
      }
      
      .modal-actions {
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        margin-top: 20px;
      }
      
      .progress-bar {
        background: #e0e0e0;
        height: 20px;
        border-radius: 10px;
        overflow: hidden;
        margin: 20px 0;
      }
      
      .progress-fill {
        background: #4CAF50;
        height: 100%;
        transition: width 0.3s;
      }
      
      .progress-text {
        text-align: center;
        color: #666;
      }
      
      .no-selection {
        color: #999;
        text-align: center;
        padding: 40px;
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
    `;
    document.head.appendChild(style);
  }
}

// Export for use in main taskpane
export function initializeWorkflowDesigner(containerId: string): WorkflowDesigner {
  return new WorkflowDesigner(containerId);
}