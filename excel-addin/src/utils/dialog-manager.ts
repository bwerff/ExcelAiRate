/**
 * Dialog Manager
 * Handles all dialog interactions for the Excel AI Assistant
 */

export interface DialogResult<T> {
  confirmed: boolean;
  data?: T;
}

export class DialogManager {
  /**
   * Show dashboard configuration dialog
   */
  async showDashboardDialog(): Promise<DialogResult<{ type: string; chartPrefs?: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>Create Dashboard</h3>
        <div class="form-group">
          <label>Dashboard Type:</label>
          <select id="dashboard-type" class="form-control">
            <option value="executive">Executive Dashboard</option>
            <option value="sales">Sales Dashboard</option>
            <option value="financial">Financial Dashboard</option>
            <option value="operational">Operational Dashboard</option>
          </select>
        </div>
        <div class="form-group">
          <label>Chart Preferences (optional):</label>
          <input type="text" id="chart-prefs" class="form-control" 
                 placeholder="e.g., modern, colorful, minimal">
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const type = (dialog.querySelector('#dashboard-type') as HTMLSelectElement).value;
        const chartPrefs = (dialog.querySelector('#chart-prefs') as HTMLInputElement).value;
        this.closeDialog(dialog);
        resolve({ confirmed: true, data: { type, chartPrefs } });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show PowerQuery dialog
   */
  async showPowerQueryDialog(): Promise<DialogResult<{ transformation: string; source?: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>PowerQuery Transformation</h3>
        <div class="form-group">
          <label>What transformation do you need?</label>
          <textarea id="pq-transformation" class="form-control" rows="3" 
                    placeholder="e.g., Clean and pivot sales data, merge tables, remove duplicates..."></textarea>
        </div>
        <div class="form-group">
          <label>Data source (optional):</label>
          <input type="text" id="pq-source" class="form-control" 
                 placeholder="e.g., Excel table, CSV file, database...">
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const transformation = (dialog.querySelector('#pq-transformation') as HTMLTextAreaElement).value.trim();
        const source = (dialog.querySelector('#pq-source') as HTMLInputElement).value.trim();
        this.closeDialog(dialog);
        resolve({ 
          confirmed: !!transformation, 
          data: transformation ? { transformation, source } : undefined 
        });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show DAX dialog
   */
  async showDAXDialog(): Promise<DialogResult<{ measure: string; context?: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>DAX Formula Helper</h3>
        <div class="form-group">
          <label>What measure do you need?</label>
          <textarea id="dax-measure" class="form-control" rows="3" 
                    placeholder="e.g., Total Sales, YTD Revenue, Customer Count, Average Order Value..."></textarea>
        </div>
        <div class="form-group">
          <label>Table/Column context (optional):</label>
          <input type="text" id="dax-context" class="form-control" 
                 placeholder="e.g., Sales[Amount], Customer[ID]...">
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const measure = (dialog.querySelector('#dax-measure') as HTMLTextAreaElement).value.trim();
        const context = (dialog.querySelector('#dax-context') as HTMLInputElement).value.trim();
        this.closeDialog(dialog);
        resolve({ 
          confirmed: !!measure, 
          data: measure ? { measure, context } : undefined 
        });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show forecast dialog
   */
  async showForecastDialog(): Promise<DialogResult<{ type: string; periods: number }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>Forecasting Configuration</h3>
        <div class="form-group">
          <label>Forecast Type:</label>
          <select id="forecast-type" class="form-control">
            <option value="linear">Linear Trend</option>
            <option value="exponential">Exponential Growth</option>
            <option value="seasonal">Seasonal Pattern</option>
            <option value="arima">ARIMA Model</option>
          </select>
        </div>
        <div class="form-group">
          <label>Number of Periods:</label>
          <input type="number" id="forecast-periods" class="form-control" 
                 value="12" min="1" max="100">
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const type = (dialog.querySelector('#forecast-type') as HTMLSelectElement).value;
        const periods = parseInt((dialog.querySelector('#forecast-periods') as HTMLInputElement).value);
        this.closeDialog(dialog);
        resolve({ confirmed: true, data: { type, periods } });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show ML analysis dialog
   */
  async showMLDialog(): Promise<DialogResult<{ analysisType: string; targetVariable?: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>Machine Learning Analysis</h3>
        <div class="form-group">
          <label>Analysis Type:</label>
          <select id="ml-type" class="form-control">
            <option value="regression">Regression Analysis</option>
            <option value="correlation">Correlation Analysis</option>
            <option value="clustering">Clustering/Segmentation</option>
            <option value="anomaly">Anomaly Detection</option>
          </select>
        </div>
        <div class="form-group">
          <label>Target Variable (optional):</label>
          <input type="text" id="ml-target" class="form-control" 
                 placeholder="e.g., Sales, Revenue, Customer_Count">
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const analysisType = (dialog.querySelector('#ml-type') as HTMLSelectElement).value;
        const targetVariable = (dialog.querySelector('#ml-target') as HTMLInputElement).value.trim();
        this.closeDialog(dialog);
        resolve({ confirmed: true, data: { analysisType, targetVariable } });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show visualization dialog
   */
  async showVisualizationDialog(): Promise<DialogResult<{ vizType: string; styleOptions?: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>Advanced Visualization</h3>
        <div class="form-group">
          <label>Visualization Type:</label>
          <select id="viz-type" class="form-control">
            <option value="heatmap">Heat Map</option>
            <option value="treemap">Tree Map</option>
            <option value="waterfall">Waterfall Chart</option>
            <option value="funnel">Funnel Chart</option>
            <option value="network">Network Diagram</option>
            <option value="custom">Custom Visualization</option>
          </select>
        </div>
        <div class="form-group">
          <label>Style Options:</label>
          <input type="text" id="viz-style" class="form-control" 
                 placeholder="e.g., dark theme, interactive, colorful">
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const vizType = (dialog.querySelector('#viz-type') as HTMLSelectElement).value;
        const styleOptions = (dialog.querySelector('#viz-style') as HTMLInputElement).value.trim();
        this.closeDialog(dialog);
        resolve({ confirmed: true, data: { vizType, styleOptions } });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show VBA dialog
   */
  async showVBADialog(): Promise<DialogResult<{ task: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>VBA Code Generator</h3>
        <div class="form-group">
          <label>What task should the VBA code perform?</label>
          <textarea id="vba-task" class="form-control" rows="4" 
                    placeholder="e.g., Automate report generation, format data, create charts, send emails..."></textarea>
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const task = (dialog.querySelector('#vba-task') as HTMLTextAreaElement).value.trim();
        this.closeDialog(dialog);
        resolve({ 
          confirmed: !!task, 
          data: task ? { task } : undefined 
        });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show SQL dialog
   */
  async showSQLDialog(): Promise<DialogResult<{ query: string; dialect: string; schema?: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>SQL Query Generator</h3>
        <div class="form-group">
          <label>Query Description:</label>
          <textarea id="sql-query" class="form-control" rows="3" 
                    placeholder="e.g., Get top 10 customers by revenue, join sales and product tables..."></textarea>
        </div>
        <div class="form-group">
          <label>SQL Dialect:</label>
          <select id="sql-dialect" class="form-control">
            <option value="tsql">T-SQL (SQL Server)</option>
            <option value="mysql">MySQL</option>
            <option value="postgres">PostgreSQL</option>
            <option value="oracle">Oracle</option>
          </select>
        </div>
        <div class="form-group">
          <label>Schema Info (optional):</label>
          <input type="text" id="sql-schema" class="form-control" 
                 placeholder="e.g., table names, column names...">
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const query = (dialog.querySelector('#sql-query') as HTMLTextAreaElement).value.trim();
        const dialect = (dialog.querySelector('#sql-dialect') as HTMLSelectElement).value;
        const schema = (dialog.querySelector('#sql-schema') as HTMLInputElement).value.trim();
        this.closeDialog(dialog);
        resolve({ 
          confirmed: !!query, 
          data: query ? { query, dialect, schema } : undefined 
        });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Show API connection dialog
   */
  async showAPIDialog(): Promise<DialogResult<{ source: string; method: string }>> {
    return new Promise((resolve) => {
      const dialog = this.createDialog(`
        <h3>API Integration</h3>
        <div class="form-group">
          <label>Data Source:</label>
          <input type="text" id="api-source" class="form-control" 
                 placeholder="e.g., Salesforce, Google Analytics, REST API...">
        </div>
        <div class="form-group">
          <label>Connection Method:</label>
          <select id="api-method" class="form-control">
            <option value="api">REST API</option>
            <option value="web-scrape">Web Scraping</option>
            <option value="database">Database Connection</option>
            <option value="file">File Import</option>
          </select>
        </div>
      `);

      dialog.querySelector('.btn-primary')!.addEventListener('click', () => {
        const source = (dialog.querySelector('#api-source') as HTMLInputElement).value.trim();
        const method = (dialog.querySelector('#api-method') as HTMLSelectElement).value;
        this.closeDialog(dialog);
        resolve({ 
          confirmed: !!source, 
          data: source ? { source, method } : undefined 
        });
      });

      dialog.querySelector('.btn-secondary')!.addEventListener('click', () => {
        this.closeDialog(dialog);
        resolve({ confirmed: false });
      });
    });
  }

  /**
   * Create base dialog structure
   */
  private createDialog(content: string): HTMLElement {
    const dialog = document.createElement('div');
    dialog.className = 'dialog-overlay';
    dialog.innerHTML = `
      <div class="dialog-content">
        ${content}
        <div class="dialog-buttons">
          <button class="btn btn-secondary">Cancel</button>
          <button class="btn btn-primary">Confirm</button>
        </div>
      </div>
    `;
    
    document.body.appendChild(dialog);
    
    // Add styles if not already present
    if (!document.getElementById('dialog-styles')) {
      const style = document.createElement('style');
      style.id = 'dialog-styles';
      style.textContent = `
        .dialog-overlay {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(0, 0, 0, 0.5);
          z-index: 10000;
          display: flex;
          align-items: center;
          justify-content: center;
        }
        .dialog-content {
          background: white;
          padding: 20px;
          border-radius: 8px;
          width: 400px;
          max-width: 90vw;
          max-height: 80vh;
          overflow-y: auto;
        }
        .dialog-content h3 {
          margin-top: 0;
          color: #333;
        }
        .form-group {
          margin-bottom: 15px;
        }
        .form-group label {
          display: block;
          margin-bottom: 5px;
          font-weight: bold;
          color: #555;
        }
        .form-control {
          width: 100%;
          padding: 8px;
          border: 1px solid #ddd;
          border-radius: 4px;
          font-size: 14px;
          font-family: inherit;
        }
        .form-control:focus {
          outline: none;
          border-color: #0078d4;
        }
        .dialog-buttons {
          display: flex;
          gap: 10px;
          justify-content: flex-end;
          margin-top: 20px;
        }
        .btn {
          padding: 8px 16px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
          font-family: inherit;
        }
        .btn-primary {
          background: #0078d4;
          color: white;
        }
        .btn-primary:hover {
          background: #106ebe;
        }
        .btn-secondary {
          background: #666;
          color: white;
        }
        .btn-secondary:hover {
          background: #555;
        }
      `;
      document.head.appendChild(style);
    }
    
    return dialog;
  }

  /**
   * Close and remove dialog
   */
  private closeDialog(dialog: HTMLElement): void {
    document.body.removeChild(dialog);
  }
}