/**
 * Advanced Excel Integration Service
 * AI-powered PivotTables, formula dependencies, and intelligent formatting
 */

import { aiService } from './ai-service';
import { smartDetection, DetectedDataInfo } from './smart-detection';
import { ExcelHelpers } from '../utils/excel-helpers';

/* global Excel, Office */

// Type definitions
export interface PivotConfig {
  sourceData: string;
  rows: PivotField[];
  columns: PivotField[];
  values: PivotValue[];
  filters: PivotFilter[];
  recommendations?: PivotRecommendation[];
  style?: string;
  showTotals?: boolean;
}

export interface PivotField {
  name: string;
  displayName?: string;
  sortOrder?: 'ascending' | 'descending';
  showSubtotals?: boolean;
}

export interface PivotValue {
  field: string;
  operation: 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNumbers' | 'stdDev' | 'variance';
  displayName?: string;
  numberFormat?: string;
}

export interface PivotFilter {
  field: string;
  type: 'value' | 'label' | 'date';
  criteria: any;
}

export interface PivotRecommendation {
  confidence: number;
  reasoning: string;
  config: Partial<PivotConfig>;
  insights: string[];
  expectedValue: string;
}

export interface FormulaDependency {
  cell: string;
  formula: string;
  precedents: string[];
  dependents: string[];
  depth: number;
  hasCircularReference: boolean;
  errors?: string[];
}

export interface DependencyGraph {
  nodes: DependencyNode[];
  edges: DependencyEdge[];
  stats: DependencyStats;
  criticalPath: string[];
  circularReferences: string[][];
}

export interface DependencyNode {
  id: string;
  address: string;
  formula?: string;
  value?: any;
  type: 'input' | 'formula' | 'output';
  sheet: string;
  depth: number;
}

export interface DependencyEdge {
  source: string;
  target: string;
  type: 'direct' | 'indirect';
}

export interface DependencyStats {
  totalCells: number;
  formulaCells: number;
  maxDepth: number;
  averageDepth: number;
  circularCount: number;
  volatileFunctions: number;
}

export interface SmartFormattingRule {
  type: 'conditional' | 'dataBar' | 'colorScale' | 'iconSet' | 'custom';
  condition?: string;
  aiGenerated: boolean;
  reasoning: string;
  parameters: any;
  priority: number;
  appliesTo: string;
}

export interface TableEnhancement {
  calculatedColumns: CalculatedColumn[];
  suggestedTotals: TotalRow[];
  slicers: SlicerConfig[];
  formatting: SmartFormattingRule[];
}

export interface CalculatedColumn {
  name: string;
  formula: string;
  description: string;
  dataType: string;
}

export interface TotalRow {
  column: string;
  function: string;
  customFormula?: string;
}

export interface SlicerConfig {
  field: string;
  style: string;
  position: { top: number; left: number };
}

export interface VisualizationConfig {
  type: 'network' | 'tree' | 'sankey' | 'heatmap';
  data: any;
  options: Record<string, any>;
  container: string;
}

export class AdvancedExcelService {
  private excelHelpers: ExcelHelpers;
  private formulaCache: Map<string, FormulaDependency> = new Map();
  private pivotCache: Map<string, PivotConfig> = new Map();

  constructor() {
    this.excelHelpers = new ExcelHelpers();
  }

  /**
   * Analyze data and recommend optimal PivotTable configuration
   */
  async analyzeDateForPivot(rangeAddress: string): Promise<PivotRecommendation[]> {
    return await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(rangeAddress);
      range.load(['values', 'address']);
      await context.sync();

      // Get smart detection info
      const dataInfo = await smartDetection.detectDataInfo(range, context);
      
      // Prepare data for AI analysis
      const dataPreview = this.prepareDataForAI(range.values, dataInfo);
      
      // Get AI recommendations
      const prompt = `Analyze this Excel data and recommend the best PivotTable configuration:
      
      Data Info:
      - Range: ${rangeAddress}
      - Has Headers: ${dataInfo.hasHeaders}
      - Column Types: ${JSON.stringify(dataInfo.columnTypes)}
      - Headers: ${JSON.stringify(dataInfo.headers)}
      - Row Count: ${range.values.length}
      
      Data Preview:
      ${dataPreview}
      
      Please recommend:
      1. Which fields should be rows (dimensions)
      2. Which fields should be columns
      3. Which fields should be values (measures) and what operations
      4. Any useful filters
      5. Key insights this pivot would reveal
      
      Return as JSON with structure:
      {
        "recommendations": [{
          "confidence": 0-1,
          "reasoning": "explanation",
          "config": {
            "rows": [{"name": "field", "sortOrder": "ascending"}],
            "columns": [],
            "values": [{"field": "field", "operation": "sum"}],
            "filters": []
          },
          "insights": ["insight1", "insight2"],
          "expectedValue": "what user will learn"
        }]
      }`;

      const response = await aiService.callAI(prompt, {
        type: 'analyze',
        maxTokens: 800,
        systemPrompt: 'You are an expert in Excel PivotTables and data analysis. Provide practical, actionable recommendations.'
      });

      // Parse AI response
      try {
        const parsed = JSON.parse(response.content);
        return parsed.recommendations.map((rec: any) => ({
          confidence: rec.confidence,
          reasoning: rec.reasoning,
          config: {
            sourceData: rangeAddress,
            rows: rec.config.rows || [],
            columns: rec.config.columns || [],
            values: rec.config.values || [],
            filters: rec.config.filters || []
          },
          insights: rec.insights || [],
          expectedValue: rec.expectedValue
        }));
      } catch (error) {
        console.error('Failed to parse AI recommendations:', error);
        return this.generateFallbackRecommendations(dataInfo, rangeAddress);
      }
    });
  }

  /**
   * Generate PivotTable using Office Scripts or VBA
   */
  async generatePivotTable(config: PivotConfig, targetCell: string): Promise<string> {
    // Since Office.js doesn't directly support PivotTable creation,
    // we'll generate Office Script code that can be executed
    
    const script = `
function main(workbook: ExcelScript.Workbook) {
  // Get the source data
  const sourceSheet = workbook.getActiveWorksheet();
  const sourceRange = sourceSheet.getRange("${config.sourceData}");
  
  // Create or get target sheet
  let pivotSheet: ExcelScript.Worksheet;
  try {
    pivotSheet = workbook.getWorksheet("Pivot Analysis");
  } catch {
    pivotSheet = workbook.addWorksheet("Pivot Analysis");
  }
  
  // Create PivotTable
  const pivotTable = pivotSheet.addPivotTable(
    "PivotTable_" + Date.now(),
    sourceRange,
    pivotSheet.getRange("${targetCell}")
  );
  
  // Add row fields
  ${config.rows.map(row => `
  pivotTable.addRowHierarchy(
    pivotTable.getHierarchy("${row.name}")
  );`).join('')}
  
  // Add column fields
  ${config.columns.map(col => `
  pivotTable.addColumnHierarchy(
    pivotTable.getHierarchy("${col.name}")
  );`).join('')}
  
  // Add value fields
  ${config.values.map(val => `
  pivotTable.addDataHierarchy(
    pivotTable.getHierarchy("${val.field}"),
    ExcelScript.AggregationFunction.${this.mapAggregationFunction(val.operation)}
  );`).join('')}
  
  // Apply style
  pivotTable.setPivotTableStyle("${config.style || 'PivotStyleMedium2'}");
  
  // Show totals
  pivotTable.setShowGrandTotalForRows(${config.showTotals !== false});
  pivotTable.setShowGrandTotalForColumns(${config.showTotals !== false});
  
  return "PivotTable created successfully!";
}`;

    // Store the script for execution
    this.storePivotScript(script, config);
    
    // Return instructions for the user
    return `PivotTable configuration ready! 
    
To create the PivotTable:
1. Open Excel Online or Desktop
2. Go to Automate tab
3. Click "New Script"
4. Paste the generated script
5. Click "Run"

The PivotTable will analyze:
- Rows: ${config.rows.map(r => r.name).join(', ')}
- Columns: ${config.columns.map(c => c.name).join(', ')}
- Values: ${config.values.map(v => `${v.operation}(${v.field})`).join(', ')}`;
  }

  /**
   * Map formula dependencies in a worksheet
   */
  async mapFormulaDependencies(worksheet?: Excel.Worksheet): Promise<DependencyGraph> {
    return await Excel.run(async (context) => {
      const ws = worksheet || context.workbook.worksheets.getActiveWorksheet();
      ws.load('name');
      
      // Get all formulas in the worksheet
      const usedRange = ws.getUsedRange();
      usedRange.load(['formulas', 'formulasR1C1', 'address', 'values']);
      await context.sync();

      const nodes: DependencyNode[] = [];
      const edges: DependencyEdge[] = [];
      const formulaCells: FormulaDependency[] = [];
      
      // Process each cell
      const formulas = usedRange.formulas;
      const values = usedRange.values;
      const baseAddress = this.parseAddress(usedRange.address);
      
      for (let row = 0; row < formulas.length; row++) {
        for (let col = 0; col < formulas[row].length; col++) {
          const formula = formulas[row][col];
          const value = values[row][col];
          
          if (formula && typeof formula === 'string' && formula.startsWith('=')) {
            const cellAddress = this.getCellAddress(baseAddress, row, col);
            
            // Extract dependencies
            const dependency = await this.analyzeFormula(cellAddress, formula, ws, context);
            formulaCells.push(dependency);
            
            // Add to graph
            nodes.push({
              id: cellAddress,
              address: cellAddress,
              formula: formula,
              value: value,
              type: 'formula',
              sheet: ws.name,
              depth: 0 // Will be calculated later
            });
            
            // Add edges for precedents
            dependency.precedents.forEach(precedent => {
              edges.push({
                source: precedent,
                target: cellAddress,
                type: 'direct'
              });
            });
          } else if (value != null) {
            const cellAddress = this.getCellAddress(baseAddress, row, col);
            nodes.push({
              id: cellAddress,
              address: cellAddress,
              value: value,
              type: 'input',
              sheet: ws.name,
              depth: 0
            });
          }
        }
      }

      // Calculate depths and detect circular references
      const { depths, circularRefs } = this.calculateDepths(nodes, edges);
      nodes.forEach(node => {
        node.depth = depths.get(node.id) || 0;
      });

      // Calculate statistics
      const stats = this.calculateDependencyStats(nodes, formulaCells);
      
      // Find critical path
      const criticalPath = this.findCriticalPath(nodes, edges);

      return {
        nodes,
        edges,
        stats,
        criticalPath,
        circularReferences: circularRefs
      };
    });
  }

  /**
   * Visualize formula dependencies
   */
  async visualizeDependencies(graph: DependencyGraph, containerId: string): Promise<VisualizationConfig> {
    // Generate D3.js visualization configuration
    const visConfig: VisualizationConfig = {
      type: 'network',
      container: containerId,
      data: {
        nodes: graph.nodes.map(node => ({
          id: node.id,
          label: node.address,
          group: node.type,
          level: node.depth,
          value: node.value,
          formula: node.formula
        })),
        links: graph.edges.map(edge => ({
          source: edge.source,
          target: edge.target,
          value: 1
        }))
      },
      options: {
        width: 800,
        height: 600,
        nodeRadius: 15,
        simulation: {
          forceStrength: -300,
          linkDistance: 100
        },
        colors: {
          input: '#4CAF50',
          formula: '#2196F3',
          output: '#FF9800',
          error: '#F44336'
        },
        showLabels: true,
        interactive: true,
        zoom: true
      }
    };

    // Generate the visualization code
    const d3Code = this.generateD3VisualizationCode(visConfig);
    
    // Store for execution
    Office.context.document.settings.set('dependencyVisualization', d3Code);
    await new Promise<void>((resolve) => {
      Office.context.document.settings.saveAsync(() => resolve());
    });

    return visConfig;
  }

  /**
   * Suggest intelligent conditional formatting rules
   */
  async suggestSmartFormatting(rangeAddress: string): Promise<SmartFormattingRule[]> {
    return await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(rangeAddress);
      range.load(['values', 'numberFormat']);
      await context.sync();

      // Analyze data
      const dataInfo = await smartDetection.detectDataInfo(range, context);
      const stats = dataInfo.statistics;
      const rules: SmartFormattingRule[] = [];

      // Get AI suggestions
      const prompt = `Suggest conditional formatting rules for this Excel data:
      
      Data Type: ${dataInfo.dataType}
      Statistics: ${JSON.stringify(stats)}
      Column Types: ${JSON.stringify(dataInfo.columnTypes)}
      Null Count: ${dataInfo.nullCount}
      
      Suggest formatting rules that would:
      1. Highlight important patterns
      2. Identify outliers
      3. Show data quality issues
      4. Make the data more readable
      
      Return as JSON array of rules with reasoning.`;

      const response = await aiService.callAI(prompt, {
        type: 'analyze',
        maxTokens: 600
      });

      // Parse AI suggestions
      try {
        const suggestions = JSON.parse(response.content);
        rules.push(...suggestions.map((s: any, index: number) => ({
          type: s.type || 'conditional',
          condition: s.condition,
          aiGenerated: true,
          reasoning: s.reasoning,
          parameters: s.parameters,
          priority: index + 1,
          appliesTo: rangeAddress
        })));
      } catch (error) {
        console.error('Failed to parse AI formatting suggestions:', error);
      }

      // Add statistical rules
      if (stats) {
        // Outlier detection
        if (stats.average && stats.stdDev) {
          const upperBound = stats.average + (2 * stats.stdDev);
          const lowerBound = stats.average - (2 * stats.stdDev);
          
          rules.push({
            type: 'conditional',
            condition: `value > ${upperBound} OR value < ${lowerBound}`,
            aiGenerated: false,
            reasoning: 'Highlight statistical outliers (2 standard deviations from mean)',
            parameters: {
              format: { backgroundColor: '#FFEB3B', fontColor: '#000000' }
            },
            priority: rules.length + 1,
            appliesTo: rangeAddress
          });
        }

        // Data bars for numeric data
        if (dataInfo.dataType === 'numeric' || dataInfo.dataType === 'currency') {
          rules.push({
            type: 'dataBar',
            aiGenerated: false,
            reasoning: 'Visualize relative values with data bars',
            parameters: {
              showValue: true,
              gradient: true,
              color: '#2196F3'
            },
            priority: rules.length + 1,
            appliesTo: rangeAddress
          });
        }

        // Color scale for percentages
        if (dataInfo.dataType === 'percentage') {
          rules.push({
            type: 'colorScale',
            aiGenerated: false,
            reasoning: 'Use color scale to show percentage ranges',
            parameters: {
              minColor: '#F44336',
              midColor: '#FFEB3B',
              maxColor: '#4CAF50',
              midpoint: 50
            },
            priority: rules.length + 1,
            appliesTo: rangeAddress
          });
        }
      }

      // Null value highlighting
      if (dataInfo.nullCount > 0) {
        rules.push({
          type: 'conditional',
          condition: 'ISBLANK(value)',
          aiGenerated: false,
          reasoning: `Highlight ${dataInfo.nullCount} empty cells for data quality`,
          parameters: {
            format: { backgroundColor: '#FFCDD2', pattern: 'lightGrid' }
          },
          priority: rules.length + 1,
          appliesTo: rangeAddress
        });
      }

      return rules.sort((a, b) => a.priority - b.priority);
    });
  }

  /**
   * Apply smart formatting rules
   */
  async applySmartFormatting(rules: SmartFormattingRule[]): Promise<void> {
    await Excel.run(async (context) => {
      for (const rule of rules) {
        try {
          const range = context.workbook.worksheets.getActiveWorksheet().getRange(rule.appliesTo);
          
          switch (rule.type) {
            case 'conditional':
              await this.applyConditionalRule(range, rule, context);
              break;
            
            case 'dataBar':
              await this.applyDataBar(range, rule.parameters);
              break;
            
            case 'colorScale':
              await this.applyColorScale(range, rule.parameters);
              break;
            
            case 'iconSet':
              await this.applyIconSet(range, rule.parameters);
              break;
          }
          
          await context.sync();
        } catch (error) {
          console.error(`Failed to apply rule: ${rule.reasoning}`, error);
        }
      }
    });
  }

  /**
   * Enhance Excel table with AI-powered features
   */
  async enhanceTable(tableName: string): Promise<TableEnhancement> {
    return await Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(tableName);
      table.load(['id', 'name']);
      
      const dataRange = table.getDataBodyRange();
      dataRange.load('values');
      
      const headerRange = table.getHeaderRowRange();
      headerRange.load('values');
      
      await context.sync();

      const headers = headerRange.values[0].map(h => String(h));
      const enhancement: TableEnhancement = {
        calculatedColumns: [],
        suggestedTotals: [],
        slicers: [],
        formatting: []
      };

      // Get AI suggestions for calculated columns
      const prompt = `Suggest calculated columns for this Excel table:
      
      Headers: ${JSON.stringify(headers)}
      Sample Data: ${JSON.stringify(dataRange.values.slice(0, 5))}
      
      Suggest:
      1. Useful calculated columns with Excel formulas
      2. Total row calculations
      3. Which columns would make good slicers
      
      Return as JSON with Excel formulas.`;

      const response = await aiService.callAI(prompt, {
        type: 'analyze',
        maxTokens: 600
      });

      try {
        const suggestions = JSON.parse(response.content);
        
        // Add calculated columns
        if (suggestions.calculatedColumns) {
          enhancement.calculatedColumns = suggestions.calculatedColumns.map((col: any) => ({
            name: col.name,
            formula: col.formula,
            description: col.description || '',
            dataType: col.dataType || 'general'
          }));
        }

        // Add total row suggestions
        if (suggestions.totals) {
          enhancement.suggestedTotals = suggestions.totals.map((total: any) => ({
            column: total.column,
            function: total.function,
            customFormula: total.customFormula
          }));
        }

        // Add slicer suggestions
        if (suggestions.slicers) {
          enhancement.slicers = suggestions.slicers.map((slicer: any, index: number) => ({
            field: slicer.field,
            style: 'SlicerStyleLight2',
            position: { 
              top: 100 + (index * 200), 
              left: 850 
            }
          }));
        }
      } catch (error) {
        console.error('Failed to parse AI suggestions:', error);
      }

      // Add smart formatting
      const formatSuggestions = await this.suggestSmartFormatting(table.getRange().address);
      enhancement.formatting = formatSuggestions;

      return enhancement;
    });
  }

  /**
   * Apply table enhancements
   */
  async applyTableEnhancements(tableName: string, enhancements: TableEnhancement): Promise<void> {
    await Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(tableName);
      
      // Add calculated columns
      for (const calcCol of enhancements.calculatedColumns) {
        const column = table.columns.add(null, [
          [calcCol.name],
          ...Array(table.rows.count).fill([calcCol.formula])
        ]);
        column.name = calcCol.name;
        await context.sync();
      }

      // Set up total row
      if (enhancements.suggestedTotals.length > 0) {
        table.showTotals = true;
        await context.sync();
        
        for (const total of enhancements.suggestedTotals) {
          const column = table.columns.getItem(total.column);
          column.getTotalRowRange().formulas = [[
            total.customFormula || `=SUBTOTAL(${this.getFunctionCode(total.function)},${column.getDataBodyRange().address})`
          ]];
        }
      }

      // Note: Slicers would require Office Scripts or VBA
      // Store slicer configuration for later execution
      if (enhancements.slicers.length > 0) {
        this.storeSlicerConfig(tableName, enhancements.slicers);
      }

      // Apply formatting
      if (enhancements.formatting.length > 0) {
        await this.applySmartFormatting(enhancements.formatting);
      }

      await context.sync();
    });
  }

  // Helper methods

  /**
   * Prepare data preview for AI analysis
   */
  private prepareDataForAI(values: any[][], dataInfo: DetectedDataInfo): string {
    const maxRows = 10;
    const preview = values.slice(0, maxRows);
    
    let result = 'Data Preview:\n';
    preview.forEach((row, index) => {
      result += `Row ${index + 1}: ${row.join('\t')}\n`;
    });
    
    if (values.length > maxRows) {
      result += `... (${values.length - maxRows} more rows)\n`;
    }
    
    return result;
  }

  /**
   * Generate fallback PivotTable recommendations
   */
  private generateFallbackRecommendations(dataInfo: DetectedDataInfo, sourceData: string): PivotRecommendation[] {
    const recommendations: PivotRecommendation[] = [];
    
    // Basic recommendation based on data types
    if (dataInfo.columnTypes) {
      const textColumns = dataInfo.headers.filter((_, i) => 
        dataInfo.columnTypes![i] === 'text' || dataInfo.columnTypes![i] === 'date'
      );
      const numericColumns = dataInfo.headers.filter((_, i) => 
        ['numeric', 'currency', 'percentage'].includes(dataInfo.columnTypes![i])
      );

      if (textColumns.length > 0 && numericColumns.length > 0) {
        recommendations.push({
          confidence: 0.7,
          reasoning: 'Basic pivot structure based on data types',
          config: {
            sourceData,
            rows: textColumns.slice(0, 2).map(name => ({ name })),
            columns: [],
            values: numericColumns.slice(0, 3).map(field => ({
              field,
              operation: 'sum'
            })),
            filters: []
          },
          insights: [
            'Group data by categorical fields',
            'Summarize numeric values',
            'Identify patterns and trends'
          ],
          expectedValue: 'Summary view of numeric data by categories'
        });
      }
    }
    
    return recommendations;
  }

  /**
   * Map aggregation function names
   */
  private mapAggregationFunction(operation: string): string {
    const mapping: Record<string, string> = {
      'sum': 'Sum',
      'count': 'Count',
      'average': 'Average',
      'max': 'Max',
      'min': 'Min',
      'product': 'Product',
      'countNumbers': 'CountNumbers',
      'stdDev': 'StandardDeviation',
      'variance': 'Variance'
    };
    return mapping[operation] || 'Sum';
  }

  /**
   * Store PivotTable script for execution
   */
  private storePivotScript(script: string, config: PivotConfig): void {
    const scripts = this.getStoredScripts();
    scripts.push({
      id: Date.now().toString(),
      type: 'pivot',
      script,
      config,
      created: new Date()
    });
    
    Office.context.document.settings.set('pivotScripts', JSON.stringify(scripts));
    Office.context.document.settings.saveAsync(() => {});
  }

  /**
   * Get stored scripts
   */
  private getStoredScripts(): any[] {
    try {
      const stored = Office.context.document.settings.get('pivotScripts');
      return stored ? JSON.parse(stored) : [];
    } catch {
      return [];
    }
  }

  /**
   * Analyze a formula and extract dependencies
   */
  private async analyzeFormula(
    cellAddress: string,
    formula: string,
    worksheet: Excel.Worksheet,
    context: Excel.RequestContext
  ): Promise<FormulaDependency> {
    // Check cache
    if (this.formulaCache.has(cellAddress)) {
      return this.formulaCache.get(cellAddress)!;
    }

    const dependency: FormulaDependency = {
      cell: cellAddress,
      formula,
      precedents: [],
      dependents: [],
      depth: 0,
      hasCircularReference: false,
      errors: []
    };

    // Extract cell references from formula
    const cellRefs = this.extractCellReferences(formula);
    dependency.precedents = cellRefs;

    // Check for errors
    if (formula.includes('#REF!')) {
      dependency.errors!.push('Reference error');
    }
    if (formula.includes('#VALUE!')) {
      dependency.errors!.push('Value error');
    }
    if (formula.includes('#DIV/0!')) {
      dependency.errors!.push('Division by zero');
    }

    this.formulaCache.set(cellAddress, dependency);
    return dependency;
  }

  /**
   * Extract cell references from a formula
   */
  private extractCellReferences(formula: string): string[] {
    const refs: string[] = [];
    
    // Basic regex for cell references (A1 style)
    const cellRegex = /\$?[A-Z]+\$?\d+/gi;
    const matches = formula.match(cellRegex);
    
    if (matches) {
      refs.push(...matches.map(ref => ref.replace(/\$/g, '')));
    }
    
    // Range references (A1:B10 style)
    const rangeRegex = /\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+/gi;
    const rangeMatches = formula.match(rangeRegex);
    
    if (rangeMatches) {
      rangeMatches.forEach(range => {
        const [start, end] = range.split(':').map(r => r.replace(/\$/g, ''));
        // For simplicity, just include start and end cells
        refs.push(start, end);
      });
    }
    
    return [...new Set(refs)];
  }

  /**
   * Parse Excel address to get sheet and range
   */
  private parseAddress(address: string): { sheet: string; range: string; startRow: number; startCol: number } {
    const parts = address.split('!');
    const sheet = parts.length > 1 ? parts[0] : '';
    const range = parts.length > 1 ? parts[1] : parts[0];
    
    // Extract start position
    const match = range.match(/([A-Z]+)(\d+)/);
    const startCol = match ? this.columnToNumber(match[1]) : 0;
    const startRow = match ? parseInt(match[2]) - 1 : 0;
    
    return { sheet, range, startRow, startCol };
  }

  /**
   * Convert column letter to number (A=0, B=1, etc.)
   */
  private columnToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 65);
    }
    return result;
  }

  /**
   * Convert number to column letter
   */
  private numberToColumn(num: number): string {
    let result = '';
    while (num >= 0) {
      result = String.fromCharCode(65 + (num % 26)) + result;
      num = Math.floor(num / 26) - 1;
    }
    return result;
  }

  /**
   * Get cell address from base address and offsets
   */
  private getCellAddress(base: any, row: number, col: number): string {
    const column = this.numberToColumn(base.startCol + col);
    const rowNum = base.startRow + row + 1;
    return `${column}${rowNum}`;
  }

  /**
   * Calculate node depths and detect circular references
   */
  private calculateDepths(
    nodes: DependencyNode[],
    edges: DependencyEdge[]
  ): { depths: Map<string, number>; circularRefs: string[][] } {
    const depths = new Map<string, number>();
    const visited = new Set<string>();
    const recursionStack = new Set<string>();
    const circularRefs: string[][] = [];
    
    // Build adjacency list
    const adjacencyList = new Map<string, string[]>();
    edges.forEach(edge => {
      if (!adjacencyList.has(edge.source)) {
        adjacencyList.set(edge.source, []);
      }
      adjacencyList.get(edge.source)!.push(edge.target);
    });
    
    // DFS to calculate depths and detect cycles
    const dfs = (nodeId: string, currentDepth: number, path: string[]): number => {
      if (recursionStack.has(nodeId)) {
        // Circular reference detected
        const cycleStart = path.indexOf(nodeId);
        circularRefs.push(path.slice(cycleStart));
        return currentDepth;
      }
      
      if (visited.has(nodeId)) {
        return depths.get(nodeId) || 0;
      }
      
      visited.add(nodeId);
      recursionStack.add(nodeId);
      path.push(nodeId);
      
      let maxChildDepth = currentDepth;
      const children = adjacencyList.get(nodeId) || [];
      
      for (const child of children) {
        const childDepth = dfs(child, currentDepth + 1, [...path]);
        maxChildDepth = Math.max(maxChildDepth, childDepth);
      }
      
      depths.set(nodeId, maxChildDepth);
      recursionStack.delete(nodeId);
      
      return maxChildDepth;
    };
    
    // Process all nodes
    nodes.forEach(node => {
      if (!visited.has(node.id)) {
        dfs(node.id, 0, []);
      }
    });
    
    return { depths, circularRefs };
  }

  /**
   * Calculate dependency statistics
   */
  private calculateDependencyStats(
    nodes: DependencyNode[],
    formulaCells: FormulaDependency[]
  ): DependencyStats {
    const depths = nodes.map(n => n.depth);
    const volatileFunctions = ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT'];
    
    let volatileCount = 0;
    formulaCells.forEach(cell => {
      if (volatileFunctions.some(fn => cell.formula.includes(fn))) {
        volatileCount++;
      }
    });
    
    return {
      totalCells: nodes.length,
      formulaCells: formulaCells.length,
      maxDepth: Math.max(...depths, 0),
      averageDepth: depths.reduce((a, b) => a + b, 0) / depths.length || 0,
      circularCount: formulaCells.filter(f => f.hasCircularReference).length,
      volatileFunctions: volatileCount
    };
  }

  /**
   * Find critical path in dependency graph
   */
  private findCriticalPath(nodes: DependencyNode[], edges: DependencyEdge[]): string[] {
    // Find nodes with maximum depth
    const maxDepth = Math.max(...nodes.map(n => n.depth), 0);
    const criticalNodes = nodes.filter(n => n.depth === maxDepth);
    
    if (criticalNodes.length === 0) {
      return [];
    }
    
    // Trace back to find a critical path
    const path: string[] = [];
    let currentNode = criticalNodes[0];
    path.push(currentNode.id);
    
    // Build reverse adjacency list
    const reverseAdjacency = new Map<string, string[]>();
    edges.forEach(edge => {
      if (!reverseAdjacency.has(edge.target)) {
        reverseAdjacency.set(edge.target, []);
      }
      reverseAdjacency.get(edge.target)!.push(edge.source);
    });
    
    // Trace back to root
    while (currentNode.depth > 0) {
      const parents = reverseAdjacency.get(currentNode.id) || [];
      if (parents.length === 0) break;
      
      // Choose parent with highest depth
      const parentNodes = parents.map(p => nodes.find(n => n.id === p)).filter(n => n);
      parentNodes.sort((a, b) => (b?.depth || 0) - (a?.depth || 0));
      
      if (parentNodes[0]) {
        currentNode = parentNodes[0];
        path.unshift(currentNode.id);
      } else {
        break;
      }
    }
    
    return path;
  }

  /**
   * Generate D3.js visualization code
   */
  private generateD3VisualizationCode(config: VisualizationConfig): string {
    return `
// D3.js Formula Dependency Visualization
const width = ${config.options.width};
const height = ${config.options.height};

const svg = d3.select("#${config.container}")
  .append("svg")
  .attr("width", width)
  .attr("height", height);

const simulation = d3.forceSimulation(${JSON.stringify(config.data.nodes)})
  .force("link", d3.forceLink(${JSON.stringify(config.data.links)}).id(d => d.id))
  .force("charge", d3.forceManyBody().strength(${config.options.simulation.forceStrength}))
  .force("center", d3.forceCenter(width / 2, height / 2));

// Add zoom behavior
const zoom = d3.zoom()
  .scaleExtent([0.1, 10])
  .on("zoom", (event) => {
    g.attr("transform", event.transform);
  });

svg.call(zoom);

const g = svg.append("g");

// Draw links
const link = g.append("g")
  .selectAll("line")
  .data(${JSON.stringify(config.data.links)})
  .enter().append("line")
  .attr("class", "link")
  .style("stroke", "#999")
  .style("stroke-opacity", 0.6)
  .style("stroke-width", d => Math.sqrt(d.value));

// Draw nodes
const node = g.append("g")
  .selectAll("circle")
  .data(${JSON.stringify(config.data.nodes)})
  .enter().append("circle")
  .attr("r", ${config.options.nodeRadius})
  .attr("fill", d => ${JSON.stringify(config.options.colors)}[d.group])
  .call(d3.drag()
    .on("start", dragstarted)
    .on("drag", dragged)
    .on("end", dragended));

// Add labels
const label = g.append("g")
  .selectAll("text")
  .data(${JSON.stringify(config.data.nodes)})
  .enter().append("text")
  .text(d => d.label)
  .style("font-size", "12px")
  .style("pointer-events", "none");

// Add tooltips
node.append("title")
  .text(d => d.formula || d.value);

// Simulation tick
simulation.on("tick", () => {
  link
    .attr("x1", d => d.source.x)
    .attr("y1", d => d.source.y)
    .attr("x2", d => d.target.x)
    .attr("y2", d => d.target.y);

  node
    .attr("cx", d => d.x)
    .attr("cy", d => d.y);

  label
    .attr("x", d => d.x + 15)
    .attr("y", d => d.y + 5);
});

// Drag functions
function dragstarted(event, d) {
  if (!event.active) simulation.alphaTarget(0.3).restart();
  d.fx = d.x;
  d.fy = d.y;
}

function dragged(event, d) {
  d.fx = event.x;
  d.fy = event.y;
}

function dragended(event, d) {
  if (!event.active) simulation.alphaTarget(0);
  d.fx = null;
  d.fy = null;
}
`;
  }

  /**
   * Apply conditional formatting rule
   */
  private async applyConditionalRule(
    range: Excel.Range,
    rule: SmartFormattingRule,
    context: Excel.RequestContext
  ): Promise<void> {
    // This is a simplified implementation
    // Full implementation would parse the condition and create appropriate Excel conditional format
    const format = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    format.custom.rule.formula = this.translateConditionToFormula(rule.condition || '');
    format.custom.format.fill.color = rule.parameters.format?.backgroundColor || '#FFEB3B';
    format.custom.format.font.color = rule.parameters.format?.fontColor || '#000000';
  }

  /**
   * Translate condition to Excel formula
   */
  private translateConditionToFormula(condition: string): string {
    // Simple translation - in practice, this would be more sophisticated
    return condition
      .replace(/value/g, 'A1')  // Replace with actual cell reference
      .replace(/OR/g, '+')
      .replace(/AND/g, '*')
      .replace(/=/g, '=');
  }

  /**
   * Apply data bar formatting
   */
  private async applyDataBar(range: Excel.Range, parameters: any): Promise<void> {
    const format = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    format.dataBar.barDirection = Excel.ConditionalDataBarDirection.context;
    format.dataBar.showDataBarOnly = !parameters.showValue;
    
    if (parameters.color) {
      format.dataBar.positiveFormat.fillColor = parameters.color;
    }
  }

  /**
   * Apply color scale formatting
   */
  private async applyColorScale(range: Excel.Range, parameters: any): Promise<void> {
    const format = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    format.colorScale.criteria = {
      minimum: { 
        color: parameters.minColor || '#F44336', 
        type: Excel.ConditionalFormatColorCriterionType.lowestValue 
      },
      midpoint: { 
        color: parameters.midColor || '#FFEB3B', 
        type: Excel.ConditionalFormatColorCriterionType.percentile, 
        value: parameters.midpoint || 50 
      },
      maximum: { 
        color: parameters.maxColor || '#4CAF50', 
        type: Excel.ConditionalFormatColorCriterionType.highestValue 
      }
    };
  }

  /**
   * Apply icon set formatting
   */
  private async applyIconSet(range: Excel.Range, parameters: any): Promise<void> {
    const format = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    format.iconSet.style = parameters.style || Excel.IconSet.threeTrafficLights1;
  }

  /**
   * Get SUBTOTAL function code
   */
  private getFunctionCode(functionName: string): number {
    const codes: Record<string, number> = {
      'average': 101,
      'count': 103,
      'counta': 103,
      'max': 104,
      'min': 105,
      'product': 106,
      'stdev': 107,
      'sum': 109,
      'var': 110
    };
    return codes[functionName.toLowerCase()] || 109;
  }

  /**
   * Store slicer configuration
   */
  private storeSlicerConfig(tableName: string, slicers: SlicerConfig[]): void {
    const configs = this.getStoredSlicerConfigs();
    configs[tableName] = slicers;
    
    Office.context.document.settings.set('slicerConfigs', JSON.stringify(configs));
    Office.context.document.settings.saveAsync(() => {});
  }

  /**
   * Get stored slicer configurations
   */
  private getStoredSlicerConfigs(): Record<string, SlicerConfig[]> {
    try {
      const stored = Office.context.document.settings.get('slicerConfigs');
      return stored ? JSON.parse(stored) : {};
    } catch {
      return {};
    }
  }
}

// Export singleton instance
export const advancedExcel = new AdvancedExcelService();