/**
 * Workflow Automation & Chaining Engine
 * Enables creation and execution of multi-step AI workflows with batch processing
 */

import { aiService } from './ai-service';
import { smartDetection } from './smart-detection';
import { ExcelHelpers } from '../utils/excel-helpers';

/* global Excel, Office */

// Workflow type definitions
export interface Workflow {
  id: string;
  name: string;
  description: string;
  steps: WorkflowStep[];
  triggers?: WorkflowTrigger[];
  variables: Record<string, any>;
  errorHandling: ErrorStrategy;
  created: Date;
  lastModified: Date;
  author?: string;
  isTemplate?: boolean;
}

export interface WorkflowStep {
  id: string;
  name: string;
  type: StepType;
  operation: string;
  inputs: StepInput[];
  outputs: StepOutput[];
  conditions?: StepCondition[];
  retryPolicy?: RetryPolicy;
  parallel?: boolean;
  timeout?: number;
}

export type StepType = 
  | 'ai-analysis' 
  | 'ai-generation'
  | 'data-transform' 
  | 'excel-operation' 
  | 'conditional' 
  | 'loop'
  | 'validation'
  | 'formatting';

export interface StepInput {
  name: string;
  type: 'range' | 'value' | 'variable' | 'previous-output';
  source: string;
  required: boolean;
  defaultValue?: any;
}

export interface StepOutput {
  name: string;
  type: 'range' | 'value' | 'variable';
  target: string;
}

export interface StepCondition {
  type: 'if' | 'unless' | 'while';
  expression: string;
  thenStep?: string;
  elseStep?: string;
}

export interface WorkflowTrigger {
  type: 'manual' | 'schedule' | 'data-change' | 'event';
  config: any;
}

export interface RetryPolicy {
  maxAttempts: number;
  delayMs: number;
  backoffMultiplier: number;
}

export type ErrorStrategy = 'stop' | 'continue' | 'retry' | 'fallback';

export interface BatchOperation {
  ranges: string[];
  sheets: string[];
  operation: WorkflowStep | Workflow;
  parallelExecution: boolean;
  progressCallback?: (progress: BatchProgress) => void;
  errorCallback?: (error: BatchError) => void;
}

export interface BatchProgress {
  total: number;
  completed: number;
  failed: number;
  currentItem: string;
  percentage: number;
}

export interface BatchError {
  item: string;
  error: Error;
  canContinue: boolean;
}

export interface WorkflowResult {
  success: boolean;
  outputs: Record<string, any>;
  errors: WorkflowError[];
  duration: number;
  steps: StepResult[];
}

export interface StepResult {
  stepId: string;
  success: boolean;
  output?: any;
  error?: Error;
  duration: number;
  retries: number;
}

export interface WorkflowError {
  stepId: string;
  message: string;
  details?: any;
  timestamp: Date;
}

export interface ExecutionContext {
  workbook: Excel.Workbook;
  worksheet: Excel.Worksheet;
  variables: Map<string, any>;
  outputs: Map<string, any>;
  user?: string;
  startTime: number;
}

// Pre-built workflow templates
export const WORKFLOW_TEMPLATES: Partial<Workflow>[] = [
  {
    name: 'Data Analysis Pipeline',
    description: 'Analyze data, generate insights, and create visualizations',
    steps: [
      {
        id: 'detect',
        name: 'Detect Data Types',
        type: 'data-transform',
        operation: 'smart-detect',
        inputs: [{ name: 'range', type: 'range', source: 'selection', required: true }],
        outputs: [{ name: 'dataInfo', type: 'variable', target: 'detectedInfo' }]
      },
      {
        id: 'analyze',
        name: 'AI Analysis',
        type: 'ai-analysis',
        operation: 'analyze',
        inputs: [
          { name: 'data', type: 'range', source: 'selection', required: true },
          { name: 'context', type: 'variable', source: 'detectedInfo', required: false }
        ],
        outputs: [{ name: 'insights', type: 'variable', target: 'analysisResults' }]
      },
      {
        id: 'format',
        name: 'Apply Smart Formatting',
        type: 'formatting',
        operation: 'conditional-format',
        inputs: [
          { name: 'range', type: 'range', source: 'selection', required: true },
          { name: 'rules', type: 'variable', source: 'analysisResults.formatting', required: false }
        ],
        outputs: []
      }
    ],
    variables: {},
    errorHandling: 'continue'
  },
  {
    name: 'Monthly Report Generator',
    description: 'Generate comprehensive monthly reports from raw data',
    steps: [
      {
        id: 'clean',
        name: 'Clean Data',
        type: 'ai-generation',
        operation: 'clean',
        inputs: [{ name: 'data', type: 'range', source: 'selection', required: true }],
        outputs: [{ name: 'cleanedData', type: 'range', target: 'Sheet2!A1' }]
      },
      {
        id: 'summarize',
        name: 'Generate Summary',
        type: 'ai-analysis',
        operation: 'summarize',
        inputs: [{ name: 'data', type: 'previous-output', source: 'clean.cleanedData', required: true }],
        outputs: [{ name: 'summary', type: 'value', target: 'summary' }]
      },
      {
        id: 'chart',
        name: 'Create Charts',
        type: 'excel-operation',
        operation: 'create-chart',
        inputs: [{ name: 'data', type: 'previous-output', source: 'clean.cleanedData', required: true }],
        outputs: []
      }
    ],
    variables: {},
    errorHandling: 'stop'
  }
];

export class WorkflowEngine {
  private excelHelpers: ExcelHelpers;
  private runningWorkflows: Map<string, AbortController> = new Map();
  private workflowHistory: WorkflowResult[] = [];

  constructor() {
    this.excelHelpers = new ExcelHelpers();
  }

  /**
   * Execute a workflow
   */
  async executeWorkflow(workflow: Workflow, initialVariables?: Record<string, any>): Promise<WorkflowResult> {
    const workflowId = `${workflow.id}_${Date.now()}`;
    const abortController = new AbortController();
    this.runningWorkflows.set(workflowId, abortController);

    const result: WorkflowResult = {
      success: true,
      outputs: {},
      errors: [],
      duration: 0,
      steps: []
    };

    const startTime = Date.now();

    try {
      await Excel.run(async (excelContext) => {
        const context: ExecutionContext = {
          workbook: excelContext.workbook,
          worksheet: excelContext.workbook.worksheets.getActiveWorksheet(),
          variables: new Map(Object.entries({ ...workflow.variables, ...initialVariables })),
          outputs: new Map(),
          startTime
        };

        // Execute steps
        for (const step of workflow.steps) {
          if (abortController.signal.aborted) {
            throw new Error('Workflow aborted');
          }

          const stepResult = await this.executeStep(step, context, workflow.errorHandling);
          result.steps.push(stepResult);

          if (!stepResult.success && workflow.errorHandling === 'stop') {
            result.success = false;
            break;
          }
        }

        // Collect outputs
        context.outputs.forEach((value, key) => {
          result.outputs[key] = value;
        });
      });
    } catch (error: any) {
      result.success = false;
      result.errors.push({
        stepId: 'workflow',
        message: error.message,
        details: error,
        timestamp: new Date()
      });
    } finally {
      result.duration = Date.now() - startTime;
      this.runningWorkflows.delete(workflowId);
      this.workflowHistory.push(result);
    }

    return result;
  }

  /**
   * Execute a single workflow step
   */
  private async executeStep(
    step: WorkflowStep,
    context: ExecutionContext,
    errorStrategy: ErrorStrategy
  ): Promise<StepResult> {
    const stepStartTime = Date.now();
    let retries = 0;
    const maxRetries = step.retryPolicy?.maxAttempts || 1;

    while (retries < maxRetries) {
      try {
        // Check conditions
        if (step.conditions && !this.evaluateConditions(step.conditions, context)) {
          return {
            stepId: step.id,
            success: true,
            output: null,
            duration: Date.now() - stepStartTime,
            retries
          };
        }

        // Prepare inputs
        const inputs = await this.prepareInputs(step.inputs, context);

        // Execute operation
        const output = await this.executeOperation(step, inputs, context);

        // Store outputs
        await this.storeOutputs(step.outputs, output, context);

        return {
          stepId: step.id,
          success: true,
          output,
          duration: Date.now() - stepStartTime,
          retries
        };

      } catch (error: any) {
        retries++;
        
        if (retries >= maxRetries) {
          if (errorStrategy === 'continue') {
            return {
              stepId: step.id,
              success: false,
              error,
              duration: Date.now() - stepStartTime,
              retries
            };
          } else {
            throw error;
          }
        }

        // Wait before retry
        if (step.retryPolicy) {
          const delay = step.retryPolicy.delayMs * Math.pow(step.retryPolicy.backoffMultiplier || 1, retries - 1);
          await new Promise(resolve => setTimeout(resolve, delay));
        }
      }
    }

    throw new Error(`Step ${step.id} failed after ${retries} retries`);
  }

  /**
   * Execute the actual operation for a step
   */
  private async executeOperation(
    step: WorkflowStep,
    inputs: Record<string, any>,
    context: ExecutionContext
  ): Promise<any> {
    switch (step.type) {
      case 'ai-analysis':
        return await this.executeAIAnalysis(step.operation, inputs);
      
      case 'ai-generation':
        return await this.executeAIGeneration(step.operation, inputs);
      
      case 'data-transform':
        return await this.executeDataTransform(step.operation, inputs, context);
      
      case 'excel-operation':
        return await this.executeExcelOperation(step.operation, inputs, context);
      
      case 'conditional':
        return await this.executeConditional(step, inputs, context);
      
      case 'loop':
        return await this.executeLoop(step, inputs, context);
      
      case 'validation':
        return await this.executeValidation(step.operation, inputs);
      
      case 'formatting':
        return await this.executeFormatting(step.operation, inputs, context);
      
      default:
        throw new Error(`Unknown step type: ${step.type}`);
    }
  }

  /**
   * Execute AI analysis operations
   */
  private async executeAIAnalysis(operation: string, inputs: Record<string, any>): Promise<any> {
    switch (operation) {
      case 'analyze':
        return await aiService.analyzeData(inputs.data, inputs.analysisType || 'summary');
      
      case 'summarize':
        return await aiService.callAI(
          `Summarize this data: ${JSON.stringify(inputs.data)}`,
          { type: 'analyze', maxTokens: 500 }
        );
      
      case 'insights':
        return await aiService.callAI(
          `Extract key insights from: ${JSON.stringify(inputs.data)}`,
          { type: 'analyze', maxTokens: 600 }
        );
      
      default:
        throw new Error(`Unknown AI analysis operation: ${operation}`);
    }
  }

  /**
   * Execute AI generation operations
   */
  private async executeAIGeneration(operation: string, inputs: Record<string, any>): Promise<any> {
    switch (operation) {
      case 'generate':
        return await aiService.callAI(inputs.prompt, { type: 'generate' });
      
      case 'clean':
        return await aiService.callAI(
          `Clean and standardize this data: ${JSON.stringify(inputs.data)}`,
          { type: 'generate' }
        );
      
      case 'formula':
        return await aiService.generateFormula(inputs.description, inputs.range);
      
      default:
        throw new Error(`Unknown AI generation operation: ${operation}`);
    }
  }

  /**
   * Execute data transformation operations
   */
  private async executeDataTransform(
    operation: string,
    inputs: Record<string, any>,
    context: ExecutionContext
  ): Promise<any> {
    switch (operation) {
      case 'smart-detect':
        return await smartDetection.analyzeSelection();
      
      case 'pivot':
        // Create pivot table configuration
        return {
          sourceData: inputs.range,
          rows: inputs.rows || [],
          columns: inputs.columns || [],
          values: inputs.values || []
        };
      
      case 'filter':
        // Apply filters
        const range = context.worksheet.getRange(inputs.range);
        range.load('values');
        await context.workbook.sync();
        
        // Filter logic would go here
        return range.values;
      
      default:
        throw new Error(`Unknown data transform operation: ${operation}`);
    }
  }

  /**
   * Execute Excel operations
   */
  private async executeExcelOperation(
    operation: string,
    inputs: Record<string, any>,
    context: ExecutionContext
  ): Promise<any> {
    switch (operation) {
      case 'create-chart':
        const chartRange = context.worksheet.getRange(inputs.data);
        const chart = context.worksheet.charts.add(
          Excel.ChartType.columnClustered,
          chartRange,
          Excel.ChartSeriesBy.auto
        );
        chart.setPosition(inputs.position || { top: 100, left: 100 });
        await context.workbook.sync();
        return chart;
      
      case 'create-table':
        const tableRange = context.worksheet.getRange(inputs.range);
        const table = context.worksheet.tables.add(tableRange, inputs.hasHeaders !== false);
        table.name = inputs.name || `Table_${Date.now()}`;
        table.style = inputs.style || 'TableStyleMedium2';
        await context.workbook.sync();
        return table;
      
      case 'copy-range':
        const sourceRange = context.worksheet.getRange(inputs.source);
        const targetRange = context.worksheet.getRange(inputs.target);
        sourceRange.copyTo(targetRange);
        await context.workbook.sync();
        return true;
      
      default:
        throw new Error(`Unknown Excel operation: ${operation}`);
    }
  }

  /**
   * Execute conditional operations
   */
  private async executeConditional(
    step: WorkflowStep,
    inputs: Record<string, any>,
    context: ExecutionContext
  ): Promise<any> {
    // Evaluate condition
    const condition = this.evaluateExpression(step.operation, context);
    
    if (condition) {
      // Execute then branch
      if (step.outputs[0]?.target) {
        return step.outputs[0].target;
      }
    } else {
      // Execute else branch
      if (step.outputs[1]?.target) {
        return step.outputs[1].target;
      }
    }
    
    return condition;
  }

  /**
   * Execute loop operations
   */
  private async executeLoop(
    step: WorkflowStep,
    inputs: Record<string, any>,
    context: ExecutionContext
  ): Promise<any> {
    const results: any[] = [];
    const items = inputs.items || [];
    
    for (const item of items) {
      context.variables.set('loopItem', item);
      // Execute loop body (would need to be defined in step configuration)
      results.push(item);
    }
    
    return results;
  }

  /**
   * Execute validation operations
   */
  private async executeValidation(operation: string, inputs: Record<string, any>): Promise<any> {
    switch (operation) {
      case 'validate-emails':
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return inputs.data.map((email: string) => ({
          value: email,
          valid: emailRegex.test(email)
        }));
      
      case 'validate-numbers':
        return inputs.data.map((value: any) => ({
          value,
          valid: !isNaN(Number(value))
        }));
      
      case 'validate-required':
        return inputs.data.map((value: any) => ({
          value,
          valid: value != null && value !== ''
        }));
      
      default:
        throw new Error(`Unknown validation operation: ${operation}`);
    }
  }

  /**
   * Execute formatting operations
   */
  private async executeFormatting(
    operation: string,
    inputs: Record<string, any>,
    context: ExecutionContext
  ): Promise<any> {
    const range = context.worksheet.getRange(inputs.range);
    
    switch (operation) {
      case 'conditional-format':
        const format = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        format.colorScale.criteria = {
          minimum: { 
            color: '#F8696B', 
            type: Excel.ConditionalFormatColorCriterionType.lowestValue 
          },
          midpoint: { 
            color: '#FFEB84', 
            type: Excel.ConditionalFormatColorCriterionType.percentile, 
            value: 50 
          },
          maximum: { 
            color: '#63BE7B', 
            type: Excel.ConditionalFormatColorCriterionType.highestValue 
          }
        };
        await context.workbook.sync();
        return true;
      
      case 'number-format':
        range.numberFormat = [[inputs.format || '#,##0.00']];
        await context.workbook.sync();
        return true;
      
      case 'auto-fit':
        range.format.autofitColumns();
        range.format.autofitRows();
        await context.workbook.sync();
        return true;
      
      default:
        throw new Error(`Unknown formatting operation: ${operation}`);
    }
  }

  /**
   * Prepare inputs for a step
   */
  private async prepareInputs(
    inputDefs: StepInput[],
    context: ExecutionContext
  ): Promise<Record<string, any>> {
    const inputs: Record<string, any> = {};

    for (const inputDef of inputDefs) {
      let value: any;

      switch (inputDef.type) {
        case 'range':
          if (inputDef.source === 'selection') {
            const selection = context.workbook.getSelectedRange();
            selection.load('values');
            await context.workbook.sync();
            value = selection.values;
          } else {
            const range = context.worksheet.getRange(inputDef.source);
            range.load('values');
            await context.workbook.sync();
            value = range.values;
          }
          break;
        
        case 'value':
          value = inputDef.source;
          break;
        
        case 'variable':
          value = context.variables.get(inputDef.source);
          break;
        
        case 'previous-output':
          const [stepId, outputName] = inputDef.source.split('.');
          value = context.outputs.get(`${stepId}.${outputName}`);
          break;
      }

      if (value === undefined && inputDef.required && inputDef.defaultValue === undefined) {
        throw new Error(`Required input ${inputDef.name} is missing`);
      }

      inputs[inputDef.name] = value ?? inputDef.defaultValue;
    }

    return inputs;
  }

  /**
   * Store outputs from a step
   */
  private async storeOutputs(
    outputDefs: StepOutput[],
    output: any,
    context: ExecutionContext
  ): Promise<void> {
    for (const outputDef of outputDefs) {
      const value = outputDef.name ? output[outputDef.name] : output;

      switch (outputDef.type) {
        case 'range':
          const range = context.worksheet.getRange(outputDef.target);
          range.values = Array.isArray(value) ? value : [[value]];
          await context.workbook.sync();
          break;
        
        case 'value':
        case 'variable':
          context.variables.set(outputDef.target, value);
          context.outputs.set(outputDef.target, value);
          break;
      }
    }
  }

  /**
   * Evaluate conditions
   */
  private evaluateConditions(conditions: StepCondition[], context: ExecutionContext): boolean {
    for (const condition of conditions) {
      const result = this.evaluateExpression(condition.expression, context);
      
      switch (condition.type) {
        case 'if':
          if (!result) return false;
          break;
        case 'unless':
          if (result) return false;
          break;
        case 'while':
          if (!result) return false;
          break;
      }
    }
    
    return true;
  }

  /**
   * Evaluate expressions (simple implementation)
   */
  private evaluateExpression(expression: string, context: ExecutionContext): boolean {
    // This is a simplified expression evaluator
    // In production, you'd want a proper expression parser
    try {
      // Replace variable references
      let evaluableExpression = expression;
      context.variables.forEach((value, key) => {
        evaluableExpression = evaluableExpression.replace(
          new RegExp(`\\$${key}`, 'g'),
          JSON.stringify(value)
        );
      });

      // Safety check - only allow certain operations
      if (!/^[\w\s\$\.\[\]<>=!&|()'"0-9]+$/.test(evaluableExpression)) {
        throw new Error('Invalid expression');
      }

      // Evaluate
      return Function('"use strict"; return (' + evaluableExpression + ')')();
    } catch (error) {
      console.error('Expression evaluation error:', error);
      return false;
    }
  }

  /**
   * Execute batch operations
   */
  async executeBatch(batch: BatchOperation): Promise<WorkflowResult[]> {
    const results: WorkflowResult[] = [];
    const items: Array<{ type: 'range' | 'sheet'; value: string }> = [];

    // Collect all items to process
    batch.ranges.forEach(range => items.push({ type: 'range', value: range }));
    batch.sheets.forEach(sheet => items.push({ type: 'sheet', value: sheet }));

    const total = items.length;
    let completed = 0;
    let failed = 0;

    // Process items
    const processItem = async (item: { type: string; value: string }) => {
      try {
        // Update progress
        if (batch.progressCallback) {
          batch.progressCallback({
            total,
            completed,
            failed,
            currentItem: item.value,
            percentage: (completed / total) * 100
          });
        }

        // Execute operation
        let result: WorkflowResult;
        
        if ('steps' in batch.operation) {
          // It's a workflow
          result = await this.executeWorkflow(batch.operation as Workflow, {
            currentRange: item.value,
            currentSheet: item.type === 'sheet' ? item.value : undefined
          });
        } else {
          // It's a single step
          result = await this.executeWorkflow({
            id: `batch_${Date.now()}`,
            name: 'Batch Operation',
            description: 'Batch operation',
            steps: [batch.operation as WorkflowStep],
            variables: { currentItem: item.value },
            errorHandling: 'continue',
            created: new Date(),
            lastModified: new Date()
          });
        }

        results.push(result);
        completed++;
        
        if (!result.success) {
          failed++;
        }

      } catch (error: any) {
        failed++;
        
        if (batch.errorCallback) {
          const canContinue = batch.operation.errorHandling !== 'stop';
          batch.errorCallback({
            item: item.value,
            error,
            canContinue
          });
          
          if (!canContinue) {
            throw error;
          }
        }
      }
    };

    // Execute in parallel or sequence
    if (batch.parallelExecution) {
      // Limit concurrent executions to prevent overwhelming Excel
      const concurrencyLimit = 5;
      const chunks: Array<Array<{ type: string; value: string }>> = [];
      
      for (let i = 0; i < items.length; i += concurrencyLimit) {
        chunks.push(items.slice(i, i + concurrencyLimit));
      }

      for (const chunk of chunks) {
        await Promise.all(chunk.map(processItem));
      }
    } else {
      for (const item of items) {
        await processItem(item);
      }
    }

    // Final progress update
    if (batch.progressCallback) {
      batch.progressCallback({
        total,
        completed,
        failed,
        currentItem: 'Complete',
        percentage: 100
      });
    }

    return results;
  }

  /**
   * Abort a running workflow
   */
  abortWorkflow(workflowId: string): boolean {
    const controller = this.runningWorkflows.get(workflowId);
    if (controller) {
      controller.abort();
      this.runningWorkflows.delete(workflowId);
      return true;
    }
    return false;
  }

  /**
   * Get workflow history
   */
  getHistory(): WorkflowResult[] {
    return this.workflowHistory;
  }

  /**
   * Clear workflow history
   */
  clearHistory(): void {
    this.workflowHistory = [];
  }

  /**
   * Save workflow as template
   */
  async saveWorkflowTemplate(workflow: Workflow): Promise<void> {
    // In a real implementation, this would save to a database or file
    const templates = this.getStoredTemplates();
    templates.push(workflow);
    Office.context.document.settings.set('workflowTemplates', JSON.stringify(templates));
    await new Promise<void>((resolve) => {
      Office.context.document.settings.saveAsync(() => resolve());
    });
  }

  /**
   * Load workflow templates
   */
  getWorkflowTemplates(): Workflow[] {
    const stored = this.getStoredTemplates();
    return [
      ...WORKFLOW_TEMPLATES.map(t => this.createWorkflowFromTemplate(t)),
      ...stored
    ];
  }

  /**
   * Get stored templates from settings
   */
  private getStoredTemplates(): Workflow[] {
    try {
      const stored = Office.context.document.settings.get('workflowTemplates');
      return stored ? JSON.parse(stored) : [];
    } catch {
      return [];
    }
  }

  /**
   * Create a full workflow from a template
   */
  private createWorkflowFromTemplate(template: Partial<Workflow>): Workflow {
    return {
      id: `template_${Date.now()}`,
      name: template.name || 'Unnamed Workflow',
      description: template.description || '',
      steps: template.steps || [],
      triggers: template.triggers,
      variables: template.variables || {},
      errorHandling: template.errorHandling || 'stop',
      created: new Date(),
      lastModified: new Date(),
      isTemplate: true,
      ...template
    };
  }
}

// Export singleton instance
export const workflowEngine = new WorkflowEngine();