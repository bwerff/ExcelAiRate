/**
 * Dashboard Builder
 * Creates comprehensive dashboards in Excel
 */

import { ExcelHelpers } from './excel-helpers';

/* global Excel */

export class DashboardBuilder {
  private excelHelpers: ExcelHelpers;
  
  constructor() {
    this.excelHelpers = new ExcelHelpers();
  }

  /**
   * Create a comprehensive dashboard
   */
  async createDashboard(
    dataRange: string,
    dashboardType: string,
    designSpecs: string
  ): Promise<void> {
    try {
      await Excel.run(async (context) => {
        // Create dashboard worksheet
        const dashboardSheet = await this.excelHelpers.createWorksheet('AI Dashboard');
        
        // Get source data
        const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
        const sourceRange = sourceSheet.getRange(dataRange);
        sourceRange.load('values, address');
        
        await context.sync();
        
        const data = sourceRange.values;
        const headers = data[0];
        const dataRows = data.slice(1);
        
        // Add dashboard title
        await this.addDashboardTitle(dashboardSheet, dashboardType);
        
        // Create KPI cards
        await this.createKPICards(dashboardSheet, dataRows, headers);
        
        // Create charts based on data analysis
        await this.createDashboardCharts(context, dashboardSheet, data, dashboardType);
        
        // Apply dashboard theme
        await this.applyDashboardTheme(dashboardSheet, dashboardType);
        
        // Auto-fit columns
        await this.excelHelpers.autoFitColumns(dashboardSheet);
        
        await context.sync();
      });
    } catch (error) {
      console.error('Dashboard creation error:', error);
      throw error;
    }
  }

  /**
   * Add dashboard title
   */
  private async addDashboardTitle(
    worksheet: Excel.Worksheet,
    dashboardType: string
  ): Promise<void> {
    await Excel.run(async (context) => {
      const titleCell = worksheet.getCell(0, 0);
      titleCell.values = [[`${dashboardType.toUpperCase()} DASHBOARD`]];
      titleCell.format.font.size = 24;
      titleCell.format.font.bold = true;
      titleCell.format.font.color = '#1f4e79';
      titleCell.format.horizontalAlignment = 'Center';
      
      // Merge cells for title
      const titleRange = worksheet.getRange('A1:H1');
      titleRange.merge();
      
      // Add timestamp
      const timestampCell = worksheet.getCell(1, 0);
      timestampCell.formulas = [['=NOW()']];
      timestampCell.numberFormat = [['mm/dd/yyyy hh:mm AM/PM']];
      timestampCell.format.font.italic = true;
      timestampCell.format.font.color = '#666666';
      
      await context.sync();
    });
  }

  /**
   * Create KPI cards
   */
  private async createKPICards(
    worksheet: Excel.Worksheet,
    dataRows: any[][],
    headers: string[]
  ): Promise<void> {
    const kpis = this.calculateKPIs(dataRows, headers);
    let row = 3;
    let col = 0;
    
    await Excel.run(async (context) => {
      for (const kpi of kpis.slice(0, 4)) { // Limit to 4 KPIs
        // Create KPI card
        const cardRange = worksheet.getRangeByIndexes(row, col, 3, 3);
        
        // Apply card styling
        cardRange.format.fill.color = kpi.color || '#0078d4';
        cardRange.format.borders.getItem('EdgeBottom').style = 'Thick';
        cardRange.format.borders.getItem('EdgeBottom').color = '#ffffff';
        
        // KPI Title
        const titleCell = worksheet.getCell(row, col);
        titleCell.values = [[kpi.title]];
        titleCell.format.font.color = 'white';
        titleCell.format.font.bold = true;
        titleCell.format.font.size = 12;
        
        // KPI Value
        const valueCell = worksheet.getCell(row + 1, col);
        valueCell.values = [[kpi.value]];
        valueCell.format.font.color = 'white';
        valueCell.format.font.bold = true;
        valueCell.format.font.size = 20;
        
        // KPI Change/Trend
        if (kpi.change) {
          const changeCell = worksheet.getCell(row + 2, col);
          changeCell.values = [[kpi.change]];
          changeCell.format.font.color = kpi.change.startsWith('+') ? '#90EE90' : '#FFB6C1';
          changeCell.format.font.size = 10;
        }
        
        col += 4;
      }
      
      await context.sync();
    });
  }

  /**
   * Create dashboard charts
   */
  private async createDashboardCharts(
    context: Excel.RequestContext,
    worksheet: Excel.Worksheet,
    data: any[][],
    dashboardType: string
  ): Promise<void> {
    const chartConfigs = this.determineOptimalCharts(data, dashboardType);
    let chartRow = 8;
    let chartCol = 0;
    
    for (const config of chartConfigs.slice(0, 4)) { // Limit to 4 charts
      // Prepare data for chart
      const chartData = this.prepareChartData(worksheet, data, config);
      
      // Create chart
      const chart = worksheet.charts.add(
        config.type,
        chartData,
        Excel.ChartSeriesBy.columns
      );
      
      // Position and size chart
      chart.top = chartRow * 20;
      chart.left = chartCol * 100;
      chart.height = 250;
      chart.width = 400;
      
      // Style chart
      chart.title.text = config.title;
      chart.legend.visible = config.showLegend;
      
      // Apply chart style
      if (config.style) {
        chart.style = config.style;
      }
      
      // Move to next position
      chartCol += 5;
      if (chartCol > 10) {
        chartCol = 0;
        chartRow += 15;
      }
    }
    
    await context.sync();
  }

  /**
   * Calculate KPIs from data
   */
  private calculateKPIs(dataRows: any[][], headers: string[]): Array<{
    title: string;
    value: string;
    color?: string;
    change?: string;
  }> {
    const kpis: Array<any> = [];
    
    // Find numerical columns
    const numericalCols = headers
      .map((header, index) => ({ header, index }))
      .filter(({ index }) => {
        const sampleData = dataRows.slice(0, 5).map(row => row[index]);
        return sampleData.every(val => !isNaN(parseFloat(val)));
      });
    
    numericalCols.forEach(({ header, index }) => {
      const values = dataRows.map(row => parseFloat(row[index])).filter(v => !isNaN(v));
      
      if (values.length > 0) {
        const stats = this.excelHelpers.calculateStatistics(values);
        
        kpis.push({
          title: `Total ${header}`,
          value: this.formatKPIValue(stats.sum),
          color: '#0078d4',
          change: values.length > 1 ? this.calculateTrend(values) : null
        });
        
        if (values.length > 1) {
          kpis.push({
            title: `Avg ${header}`,
            value: this.formatKPIValue(stats.average),
            color: '#107c10'
          });
        }
      }
    });
    
    // Add record count KPI
    kpis.unshift({
      title: 'Total Records',
      value: dataRows.length.toString(),
      color: '#5c2d91'
    });
    
    return kpis;
  }

  /**
   * Determine optimal chart types
   */
  private determineOptimalCharts(data: any[][], dashboardType: string): Array<{
    type: Excel.ChartType;
    title: string;
    columns: number[];
    showLegend: boolean;
    style?: number;
  }> {
    const headers = data[0];
    const columnTypes = this.excelHelpers.detectColumnTypes(data);
    const charts: Array<any> = [];
    
    // Find date columns for time series
    const dateColumns = columnTypes
      .map((type, index) => ({ type, index }))
      .filter(({ type }) => type === 'date')
      .map(({ index }) => index);
    
    // Find numerical columns
    const numericalColumns = columnTypes
      .map((type, index) => ({ type, index }))
      .filter(({ type }) => type === 'number')
      .map(({ index }) => index);
    
    // Find categorical columns
    const categoricalColumns = columnTypes
      .map((type, index) => ({ type, index }))
      .filter(({ type }) => type === 'text')
      .map(({ index }) => index);
    
    // Time series chart if date column exists
    if (dateColumns.length > 0 && numericalColumns.length > 0) {
      charts.push({
        type: Excel.ChartType.line,
        title: 'Trend Analysis',
        columns: [dateColumns[0], ...numericalColumns.slice(0, 3)],
        showLegend: true,
        style: 4
      });
    }
    
    // Category comparison chart
    if (categoricalColumns.length > 0 && numericalColumns.length > 0) {
      charts.push({
        type: Excel.ChartType.columnClustered,
        title: 'Category Comparison',
        columns: [categoricalColumns[0], numericalColumns[0]],
        showLegend: false,
        style: 3
      });
    }
    
    // Distribution chart
    if (numericalColumns.length >= 2) {
      charts.push({
        type: Excel.ChartType.pie,
        title: 'Distribution Analysis',
        columns: [categoricalColumns[0] || 0, numericalColumns[0]],
        showLegend: true,
        style: 5
      });
    }
    
    // Correlation chart
    if (numericalColumns.length >= 2) {
      charts.push({
        type: Excel.ChartType.xyScatter,
        title: 'Correlation Analysis',
        columns: numericalColumns.slice(0, 2),
        showLegend: false,
        style: 6
      });
    }
    
    return charts;
  }

  /**
   * Prepare chart data
   */
  private prepareChartData(
    worksheet: Excel.Worksheet,
    data: any[][],
    config: any
  ): Excel.Range {
    // For simplicity, return a range with sample data
    // In production, this would prepare the actual data based on config
    const startRow = 20;
    const dataForChart = data.map(row => 
      config.columns.map((col: number) => row[col])
    );
    
    worksheet.getRangeByIndexes(startRow, 0, dataForChart.length, config.columns.length)
      .values = dataForChart;
    
    return worksheet.getRangeByIndexes(
      startRow, 
      0, 
      dataForChart.length, 
      config.columns.length
    );
  }

  /**
   * Apply dashboard theme
   */
  private async applyDashboardTheme(
    worksheet: Excel.Worksheet,
    dashboardType: string
  ): Promise<void> {
    const themes: Record<string, any> = {
      executive: { primary: '#1f4e79', secondary: '#70ad47', accent: '#ffc000' },
      sales: { primary: '#c55a11', secondary: '#70ad47', accent: '#264478' },
      financial: { primary: '#375623', secondary: '#843c0c', accent: '#3f3f3f' },
      operational: { primary: '#7030a0', secondary: '#0070c0', accent: '#00b050' }
    };
    
    const theme = themes[dashboardType] || themes.executive;
    
    await Excel.run(async (context) => {
      // Set worksheet tab color
      worksheet.tabColor = theme.primary;
      
      // Apply theme to headers
      const headerRange = worksheet.getRange('A3:H3');
      headerRange.format.fill.color = theme.primary;
      headerRange.format.font.color = 'white';
      headerRange.format.font.bold = true;
      
      await context.sync();
    });
  }

  /**
   * Format KPI value
   */
  private formatKPIValue(value: number): string {
    if (value >= 1000000) {
      return `${(value / 1000000).toFixed(1)}M`;
    } else if (value >= 1000) {
      return `${(value / 1000).toFixed(1)}K`;
    } else {
      return value.toFixed(2);
    }
  }

  /**
   * Calculate trend
   */
  private calculateTrend(values: number[]): string | null {
    if (values.length < 2) return null;
    
    const recent = values.slice(-Math.min(5, values.length));
    const earlier = values.slice(-Math.min(10, values.length), -5);
    
    if (earlier.length === 0) return null;
    
    const recentAvg = recent.reduce((a, b) => a + b, 0) / recent.length;
    const earlierAvg = earlier.reduce((a, b) => a + b, 0) / earlier.length;
    
    const change = ((recentAvg - earlierAvg) / earlierAvg * 100);
    return `${change >= 0 ? '+' : ''}${change.toFixed(1)}%`;
  }

  /**
   * Create a heat map
   */
  async createHeatMap(worksheet: Excel.Worksheet, data: any[][]): Promise<void> {
    await Excel.run(async (context) => {
      const range = worksheet.getRangeByIndexes(0, 0, data.length, data[0].length);
      range.values = data;
      
      await this.excelHelpers.applyConditionalFormatting(range, 'colorScale', {
        minColor: '#F8696B',
        midColor: '#FFEB84',
        maxColor: '#63BE7B'
      });
      
      await context.sync();
    });
  }

  /**
   * Create sparklines
   */
  async createSparklines(
    worksheet: Excel.Worksheet,
    dataRange: string,
    sparklineRange: string
  ): Promise<void> {
    await Excel.run(async (context) => {
      // Note: Office.js doesn't directly support sparklines
      // This would need to be implemented using Office Scripts or VBA
      const message = `Sparklines would be created here:
      Data: ${dataRange}
      Location: ${sparklineRange}`;
      
      worksheet.getRange(sparklineRange).values = [[message]];
      await context.sync();
    });
  }
}