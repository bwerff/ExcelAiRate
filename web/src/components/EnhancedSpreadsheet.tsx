'use client'

import React, { useState, useCallback, useRef, useEffect } from 'react'
import { Download, Upload, Plus, Sparkles, Bold, Italic, Underline } from 'lucide-react'

const ROWS = 50
const COLS = 78 // Extends to column BZ (A=1, B=2, ..., Z=26, AA=27, ..., BZ=78)

interface EnhancedSpreadsheetProps {
  className?: string
  isDarkMode?: boolean
}

interface CellData {
  value: string
  formula?: string
  style?: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    backgroundColor?: string
    textColor?: string
    fontSize?: number
    textAlign?: 'left' | 'center' | 'right'
  }
}

export function EnhancedSpreadsheet({ className, isDarkMode = true }: EnhancedSpreadsheetProps) {
  // State for grid data and selection
  const [gridData, setGridData] = useState<CellData[][]>(() => 
    Array(ROWS).fill(null).map(() => 
      Array(COLS).fill(null).map(() => ({ value: '', style: {} }))
    )
  )
  const [selectedCell, setSelectedCell] = useState({ row: 0, col: 0 })
  const [selectedRange, setSelectedRange] = useState({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 }
  })
  const [isEditing, setIsEditing] = useState<{ row: number; col: number } | null>(null)
  const [columnWidths, setColumnWidths] = useState<number[]>(Array(COLS).fill(100))
  const [rowHeights, setRowHeights] = useState<number[]>(Array(ROWS).fill(32))
  const [formulaBarValue, setFormulaBarValue] = useState('')
  const [isEditingFormula, setIsEditingFormula] = useState(false)
  const [referencedCells, setReferencedCells] = useState<Array<{row: number, col: number}>>([])
  
  // UI state
  const [activeWorksheet, setActiveWorksheet] = useState('Sheet1')
  const [worksheets, setWorksheets] = useState(['Sheet1'])
  const [showAIPanel, setShowAIPanel] = useState(false)
  const [aiPrompt, setAiPrompt] = useState('')
  const [aiLoading, setAiLoading] = useState(false)

  const fileInputRef = useRef<HTMLInputElement>(null)
  const gridRef = useRef<HTMLDivElement>(null)
  const editInputRef = useRef<HTMLInputElement>(null)
  const formulaBarRef = useRef<HTMLInputElement>(null)

  // Helper functions
  const getCellRef = (row: number, col: number) => {
    let columnName = ''
    let tempCol = col
    
    if (tempCol < 26) {
      // Single letter columns (A-Z)
      columnName = String.fromCharCode(65 + tempCol)
    } else {
      // Multi-letter columns (AA-BZ)
      const firstLetter = Math.floor(tempCol / 26) - 1
      const secondLetter = tempCol % 26
      columnName = String.fromCharCode(65 + firstLetter) + String.fromCharCode(65 + secondLetter)
    }
    
    return `${columnName}${row + 1}`
  }

  const parseCellRef = (cellRef: string) => {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/)
    if (!match) return { row: 0, col: 0 }
    
    const columnName = match[1]
    const row = parseInt(match[2]) - 1
    
    let col = 0
    if (columnName.length === 1) {
      // Single letter (A-Z)
      col = columnName.charCodeAt(0) - 65
    } else if (columnName.length === 2) {
      // Two letters (AA-BZ)
      const firstLetter = columnName.charCodeAt(0) - 65
      const secondLetter = columnName.charCodeAt(1) - 65
      col = (firstLetter + 1) * 26 + secondLetter
    }
    
    return { row, col }
  }

  const evaluateFormula = (formula: string, row: number, col: number): string => {
    if (!formula.startsWith('=')) return formula
    
    try {
      // Simple formula evaluation - in a real app, you'd use a proper formula parser
      const expression = formula.slice(1)
      
      // Handle SUM function
      if (expression.startsWith('SUM(')) {
        const range = expression.match(/SUM\(([A-Z]+\d+):([A-Z]+\d+)\)/)?.[1]
        if (range) {
          const [start, end] = expression.match(/SUM\(([A-Z]+\d+):([A-Z]+\d+)\)/)?.slice(1, 3) || []
          if (start && end) {
            const startPos = parseCellRef(start)
            const endPos = parseCellRef(end)
            let sum = 0
            for (let r = startPos.row; r <= endPos.row; r++) {
              for (let c = startPos.col; c <= endPos.col; c++) {
                const cellValue = gridData[r]?.[c]?.value || '0'
                const num = parseFloat(cellValue)
                if (!isNaN(num)) sum += num
              }
            }
            return sum.toString()
          }
        }
      }
      
      // Handle AVERAGE function
      if (expression.startsWith('AVERAGE(')) {
        const [start, end] = expression.match(/AVERAGE\(([A-Z]+\d+):([A-Z]+\d+)\)/)?.slice(1, 3) || []
        if (start && end) {
          const startPos = parseCellRef(start)
          const endPos = parseCellRef(end)
          let sum = 0
          let count = 0
          for (let r = startPos.row; r <= endPos.row; r++) {
            for (let c = startPos.col; c <= endPos.col; c++) {
              const cellValue = gridData[r]?.[c]?.value || '0'
              const num = parseFloat(cellValue)
              if (!isNaN(num)) {
                sum += num
                count++
              }
            }
          }
          return count > 0 ? (sum / count).toString() : '0'
        }
      }
      
      // Handle simple arithmetic
      const result = eval(expression.replace(/[A-Z]+\d+/g, (match) => {
        const pos = parseCellRef(match)
        const cellValue = gridData[pos.row]?.[pos.col]?.value || '0'
        return parseFloat(cellValue) || 0
      }))
      
      return result.toString()
    } catch (error) {
      return '#ERROR!'
    }
  }

  const handleCellClick = useCallback((row: number, col: number, event?: React.MouseEvent) => {
    // If we're editing a formula in the formula bar, insert cell reference
    if (isEditingFormula && formulaBarValue.startsWith('=')) {
      event?.preventDefault()
      event?.stopPropagation()
      
      const cellRef = getCellRef(row, col)
      setFormulaBarValue(prev => prev + cellRef)
      
      // Track referenced cells for highlighting
      setReferencedCells(prev => [...prev, { row, col }])
      
      // Keep focus on formula bar
      setTimeout(() => formulaBarRef.current?.focus(), 10)
      
      return
    }
    
    if (event?.shiftKey && selectedCell) {
      // Range selection
      setSelectedRange({
        start: selectedCell,
        end: { row, col }
      })
    } else {
      setSelectedCell({ row, col })
      setSelectedRange({
        start: { row, col },
        end: { row, col }
      })
      setIsEditing(null)
      
      // Update formula bar with selected cell's content
      const cellData = gridData[row][col]
      setFormulaBarValue(cellData.formula || cellData.value)
      setIsEditingFormula(false)
      setReferencedCells([])
    }
  }, [selectedCell, isEditingFormula, formulaBarValue, gridData])

  const handleCellDoubleClick = useCallback((row: number, col: number) => {
    setSelectedCell({ row, col })
    setIsEditing({ row, col })
    setTimeout(() => editInputRef.current?.focus(), 0)
  }, [])

  const handleCellChange = useCallback((row: number, col: number, value: string) => {
    const newGridData = [...gridData]
    
    if (value.startsWith('=')) {
      // It's a formula
      newGridData[row][col] = {
        ...newGridData[row][col],
        formula: value,
        value: evaluateFormula(value, row, col)
      }
    } else {
      // Regular value
      newGridData[row][col] = {
        ...newGridData[row][col],
        value: value,
        formula: undefined
      }
    }
    
    setGridData(newGridData)
  }, [gridData])

  const handleEditComplete = useCallback(() => {
    setIsEditing(null)
  }, [])

  const handleKeyDown = useCallback((event: KeyboardEvent) => {
    if (isEditing) return

    const { row, col } = selectedCell
    let newRow = row
    let newCol = col

    switch (event.key) {
      case 'ArrowUp':
        newRow = Math.max(0, row - 1)
        break
      case 'ArrowDown':
        newRow = Math.min(ROWS - 1, row + 1)
        break
      case 'ArrowLeft':
        newCol = Math.max(0, col - 1)
        break
      case 'ArrowRight':
        newCol = Math.min(COLS - 1, col + 1)
        break
      case 'Enter':
        if (event.shiftKey) {
          newRow = Math.max(0, row - 1)
        } else {
          newRow = Math.min(ROWS - 1, row + 1)
        }
        break
      case 'Tab':
        event.preventDefault()
        if (event.shiftKey) {
          newCol = Math.max(0, col - 1)
        } else {
          newCol = Math.min(COLS - 1, col + 1)
        }
        break
      case 'F2':
        event.preventDefault()
        setIsEditing({ row, col })
        setTimeout(() => editInputRef.current?.focus(), 0)
        return
      case 'Delete':
        handleCellChange(row, col, '')
        return
      default:
        if (event.key.length === 1 && !event.ctrlKey && !event.altKey) {
          setIsEditing({ row, col })
          handleCellChange(row, col, event.key)
          setTimeout(() => editInputRef.current?.focus(), 0)
          return
        }
    }

    if (newRow !== row || newCol !== col) {
      setSelectedCell({ row: newRow, col: newCol })
      setSelectedRange({
        start: { row: newRow, col: newCol },
        end: { row: newRow, col: newCol }
      })
    }
  }, [selectedCell, isEditing, handleCellChange])

  // Add keyboard event listener
  useEffect(() => {
    const handleGlobalKeyDown = (event: KeyboardEvent) => {
      handleKeyDown(event)
    }

    document.addEventListener('keydown', handleGlobalKeyDown)
    return () => document.removeEventListener('keydown', handleGlobalKeyDown)
  }, [handleKeyDown])

  // Sync formula bar with selected cell (only when not editing)
  useEffect(() => {
    if (!isEditingFormula) {
      const cellData = gridData[selectedCell.row][selectedCell.col]
      setFormulaBarValue(cellData.formula || cellData.value)
    }
  }, [selectedCell, gridData, isEditingFormula])

  const handleImport = useCallback(() => {
    fileInputRef.current?.click()
  }, [])

  const handleFileUpload = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (file) {
      console.log('Importing file:', file.name)
      // Simulate file import
      const sampleData = [
        ['Product', 'Q1 Sales', 'Q2 Sales', 'Q3 Sales', 'Q4 Sales', 'Total'],
        ['Widget A', '15000', '18000', '22000', '25000', '=SUM(B2:E2)'],
        ['Widget B', '8000', '12000', '15000', '18000', '=SUM(B3:E3)'],
        ['Widget C', '5000', '7000', '9000', '11000', '=SUM(B4:E4)'],
        ['Total', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)', '=SUM(E2:E4)', '=SUM(F2:F4)'],
      ]
      
      const newGridData = [...gridData]
      sampleData.forEach((row, rowIndex) => {
        if (rowIndex < newGridData.length) {
          row.forEach((cellValue, colIndex) => {
            if (colIndex < newGridData[rowIndex].length) {
              if (cellValue.startsWith('=')) {
                newGridData[rowIndex][colIndex] = {
                  ...newGridData[rowIndex][colIndex],
                  formula: cellValue,
                  value: evaluateFormula(cellValue, rowIndex, colIndex)
                }
              } else {
                newGridData[rowIndex][colIndex] = {
                  ...newGridData[rowIndex][colIndex],
                  value: cellValue,
                  formula: undefined
                }
              }
            }
          })
        }
      })
      setGridData(newGridData)
    }
  }, [gridData])

  const handleExport = useCallback((format: string) => {
    console.log('Exporting as:', format)
    // Export the actual values, not formulas
    const csvContent = gridData.map(row => 
      row.map(cell => cell.value).join(',')
    ).join('\n')
    const blob = new Blob([csvContent], { type: 'text/csv' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'spreadsheet.csv'
    a.click()
    URL.revokeObjectURL(url)
  }, [gridData])

  const formatCellStyle = (style: any) => ({
    fontWeight: style.bold ? 'bold' : 'normal',
    fontStyle: style.italic ? 'italic' : 'normal',
    textDecoration: style.underline ? 'underline' : 'none',
    backgroundColor: style.backgroundColor || 'transparent',
    color: style.textColor || (style.color || 'inherit'),
    textAlign: style.textAlign || 'left',
    fontSize: style.fontSize || '14px',
  })

  const renderCell = (cell: CellData, rowIndex: number, colIndex: number) => {
    const cellRef = getCellRef(rowIndex, colIndex)
    const isSelected = selectedRange.start.row <= rowIndex && 
                      rowIndex <= selectedRange.end.row &&
                      selectedRange.start.col <= colIndex && 
                      colIndex <= selectedRange.end.col
    
    const isReferenced = referencedCells.some(ref => ref.row === rowIndex && ref.col === colIndex)
    
    return (
      <div
        key={`${rowIndex}-${colIndex}`}
        className={`
          border-r border-b flex items-center px-1 text-xs cursor-cell relative transition-colors
          ${isDarkMode ? 'border-gray-700' : 'border-gray-200'}
          ${isReferenced
            ? isDarkMode
              ? 'bg-purple-900/30 ring-2 ring-purple-400 ring-inset z-10'
              : 'bg-purple-50 ring-2 ring-purple-400 ring-inset z-10'
            : ''}
          ${isSelected && !isReferenced
            ? isDarkMode 
              ? 'bg-blue-900/30 ring-2 ring-blue-400 ring-inset z-10' 
              : 'bg-blue-50 ring-2 ring-blue-400 ring-inset z-10'
            : ''}
          ${isEditing && isEditing.row === rowIndex && isEditing.col === colIndex 
            ? isDarkMode ? 'bg-gray-700' : 'bg-white' 
            : isDarkMode ? 'hover:bg-gray-800' : 'hover:bg-gray-25'}
        `}
        style={{
          width: `${columnWidths[colIndex] * 0.7}px`,
          height: `${rowHeights[rowIndex] * 0.7}px`,
          fontSize: '11px',
          lineHeight: '1.2',
          ...formatCellStyle(cell.style)
        }}
        onClick={(e) => {
          if (!isEditingFormula || !formulaBarValue.startsWith('=')) {
            handleCellClick(rowIndex, colIndex, e)
          } else {
            e.preventDefault()
            handleCellClick(rowIndex, colIndex, e)
          }
        }}
        onDoubleClick={() => handleCellDoubleClick(rowIndex, colIndex)}
        title={cell.formula ? `Formula: ${cell.formula}` : cell.value}
      >
        {isEditing && isEditing.row === rowIndex && isEditing.col === colIndex ? (
          <input
            ref={editInputRef}
            type="text"
            value={cell.formula || cell.value}
            onChange={(e) => handleCellChange(rowIndex, colIndex, e.target.value)}
            onBlur={() => handleEditComplete()}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === 'Tab') {
                e.preventDefault()
                handleEditComplete()
              } else if (e.key === 'Escape') {
                setIsEditing(null)
              }
            }}
            className={`w-full h-full border-none outline-none text-xs px-1 ${
              isDarkMode ? 'bg-gray-700 text-gray-200' : 'bg-white text-gray-800'
            }`}
            style={{ fontSize: '11px' }}
            autoFocus
          />
        ) : (
          <span className={`truncate w-full ${
            isDarkMode ? 'text-gray-200' : 'text-gray-800'
          }`}>
            {cell.value}
          </span>
        )}
      </div>
    )
  }

  const renderColumnHeader = (colIndex: number) => (
    <div
      key={`col-${colIndex}`}
      className={`border-r border-b flex items-center justify-center text-xs font-medium transition-colors ${
        isDarkMode 
          ? 'bg-gray-800 border-gray-700 text-gray-300 hover:bg-gray-700' 
          : 'bg-gray-100 border-gray-300 text-gray-700 hover:bg-gray-150'
      }`}
      style={{ 
        width: `${columnWidths[colIndex] * 0.7}px`,
        height: '18px',
        fontSize: '10px'
      }}
    >
      {getCellRef(0, colIndex).replace(/\d+$/, '')}
    </div>
  )

  const renderRowHeader = (rowIndex: number) => (
    <div
      key={`row-${rowIndex}`}
      className={`border-r border-b flex items-center justify-center text-xs font-medium transition-colors ${
        isDarkMode 
          ? 'bg-gray-800 border-gray-700 text-gray-300 hover:bg-gray-700' 
          : 'bg-gray-100 border-gray-300 text-gray-700 hover:bg-gray-150'
      }`}
      style={{ 
        width: '32px',
        height: `${rowHeights[rowIndex] * 0.7}px`,
        fontSize: '10px'
      }}
    >
      {rowIndex + 1}
    </div>
  )

  const handleAIPrompt = useCallback(async () => {
    if (!aiPrompt.trim()) return
    
    setAiLoading(true)
    
    // Simulate AI processing
    setTimeout(() => {
      const aiResponses = [
        'SUM(B2:E2)',
        'AVERAGE(B2:E2)',
        'MAX(B2:E2)',
        'MIN(B2:E2)',
      ]
      
      const response = aiResponses[Math.floor(Math.random() * aiResponses.length)]
      
      // Add AI response to next available cell
      const newGridData = [...gridData]
      let emptyCellFound = false
      
      for (let row = 0; row < newGridData.length; row++) {
        for (let col = 0; col < newGridData[row].length; col++) {
          if (!newGridData[row][col].value) {
            newGridData[row][col] = { value: response, style: {} }
            emptyCellFound = true
            break
          }
        }
        if (emptyCellFound) break
      }
      
      setGridData(newGridData)
      setAiLoading(false)
      setAiPrompt('')
    }, 1000)
  }, [aiPrompt, gridData])

  const handleAddSheet = useCallback(() => {
    const newSheet = `Sheet${worksheets.length + 1}`
    setWorksheets([...worksheets, newSheet])
    setActiveWorksheet(newSheet)
  }, [worksheets])

  return (
    <>
      <style jsx>{`
        .bg-blue-25 {
          background-color: #f0f8ff;
        }
        .bg-gray-25 {
          background-color: #fafafa;
        }
        .hover\\:bg-gray-150:hover {
          background-color: #f5f5f5;
        }
        .excel-grid {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        /* Dark mode cell hover */
        .dark-hover\\:bg-gray-800:hover {
          background-color: #1f2937;
        }
        /* Ensure proper contrast in cells */
        .cell-dark {
          color: #e5e7eb;
        }
        .cell-light {
          color: #1f2937;
        }
      `}</style>
      <div className={`flex flex-col h-full border rounded-lg shadow-sm excel-grid transition-colors duration-300 ${
        isDarkMode 
          ? 'bg-gray-800 border-gray-700' 
          : 'bg-white border-gray-200'
      } ${className || ''}`}>
      {/* Toolbar */}
      <div className={`flex items-center gap-1 px-2 py-1 border-b transition-colors duration-300 ${
        isDarkMode 
          ? 'bg-gray-900 border-gray-700' 
          : 'bg-gray-50 border-gray-200'
      }`}>
        <button
          onClick={handleImport}
          className={`flex items-center gap-1 px-2 py-1 text-xs border rounded transition-colors ${
            isDarkMode
              ? 'border-gray-600 hover:bg-gray-700 bg-gray-800 text-gray-200'
              : 'border-gray-300 hover:bg-gray-100 bg-white text-gray-700'
          }`}
        >
          <Upload className="w-3 h-3" />
          Import
        </button>
        
        <button
          onClick={() => handleExport('csv')}
          className={`flex items-center gap-1 px-2 py-1 text-xs border rounded transition-colors ${
            isDarkMode
              ? 'border-gray-600 hover:bg-gray-700 bg-gray-800 text-gray-200'
              : 'border-gray-300 hover:bg-gray-100 bg-white text-gray-700'
          }`}
        >
          <Download className="w-3 h-3" />
          Export
        </button>

        <div className={`h-4 w-px mx-1 transition-colors ${
          isDarkMode ? 'bg-gray-600' : 'bg-gray-300'
        }`} />

        <button
          onClick={() => setShowAIPanel(!showAIPanel)}
          className={`flex items-center gap-1 px-2 py-1 text-xs border rounded transition-colors ${
            isDarkMode
              ? 'border-purple-500 hover:bg-purple-900/30 bg-purple-900/20 text-purple-300'
              : 'border-purple-300 hover:bg-purple-100 bg-purple-50 text-purple-700'
          }`}
        >
          <Sparkles className="w-3 h-3" />
          AI
        </button>

        <div className="flex-1" />

        <button
          onClick={handleAddSheet}
          className={`flex items-center gap-1 px-2 py-1 text-xs border rounded transition-colors ${
            isDarkMode
              ? 'border-gray-600 hover:bg-gray-700 bg-gray-800 text-gray-200'
              : 'border-gray-300 hover:bg-gray-100 bg-white text-gray-700'
          }`}
        >
          <Plus className="w-3 h-3" />
          New Sheet
        </button>
      </div>

      {/* Sheet Tabs */}
      <div className={`flex items-center gap-1 px-2 py-1 border-b transition-colors duration-300 ${
        isDarkMode 
          ? 'bg-gray-900 border-gray-700' 
          : 'bg-gray-100 border-gray-200'
      }`}>
        {worksheets.map(sheet => (
          <button
            key={sheet}
            onClick={() => setActiveWorksheet(sheet)}
            className={`px-2 py-1 text-xs border rounded transition-colors ${
              activeWorksheet === sheet 
                ? 'bg-blue-500 text-white border-blue-500' 
                : isDarkMode
                  ? 'bg-gray-800 hover:bg-gray-700 border-gray-600 text-gray-200'
                  : 'bg-white hover:bg-gray-200 border-gray-300 text-gray-700'
            }`}
          >
            {sheet}
          </button>
        ))}
      </div>

      {/* Formula Bar */}
      <div className={`flex items-center px-2 py-1 border-b transition-colors duration-300 ${
        isDarkMode 
          ? 'border-gray-700 bg-gray-800' 
          : 'border-gray-200 bg-white'
      }`}>
        <div className={`px-2 py-1 border rounded text-xs font-mono mr-2 min-w-[50px] text-center transition-colors ${
          isDarkMode 
            ? 'bg-gray-700 border-gray-600 text-gray-200' 
            : 'bg-gray-100 border-gray-300 text-gray-700'
        }`}>
          {getCellRef(selectedCell.row, selectedCell.col)}
        </div>
        <input
          ref={formulaBarRef}
          value={formulaBarValue}
          onChange={(e) => {
            setFormulaBarValue(e.target.value)
            setIsEditingFormula(true)
            
            // Clear referenced cells when starting fresh
            if (e.target.value === '=') {
              setReferencedCells([])
            }
          }}
          onFocus={() => {
            setIsEditingFormula(true)
            if (formulaBarValue.startsWith('=')) {
              // Parse existing formula for cell references
              const refs = formulaBarValue.match(/[A-Z]+\d+/g) || []
              const cells = refs.map(ref => parseCellRef(ref))
              setReferencedCells(cells)
            }
          }}
          onBlur={(e) => {
            // Don't blur if clicking within the spreadsheet
            const relatedTarget = e.relatedTarget as HTMLElement
            const isClickingCell = relatedTarget?.closest('.excel-grid')
            
            if (!isClickingCell) {
              setIsEditingFormula(false)
              setReferencedCells([])
            }
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              handleCellChange(selectedCell.row, selectedCell.col, formulaBarValue)
              setIsEditingFormula(false)
              setReferencedCells([])
            } else if (e.key === 'Escape') {
              e.preventDefault()
              const cellData = gridData[selectedCell.row][selectedCell.col]
              setFormulaBarValue(cellData.formula || cellData.value)
              setIsEditingFormula(false)
              setReferencedCells([])
            }
          }}
          className={`flex-1 px-2 py-1 border rounded font-mono text-xs focus:outline-none focus:ring-1 focus:ring-blue-400 transition-colors ${
            isDarkMode 
              ? 'border-gray-600 bg-gray-700 text-gray-200' 
              : 'border-gray-300 bg-white text-gray-700'
          }`}
          placeholder="Enter value or formula..."
        />
      </div>

      <div className="flex flex-1 overflow-hidden">
        {/* Spreadsheet Grid */}
        <div className={`flex-1 overflow-auto transition-colors ${
          isDarkMode ? 'bg-gray-900' : 'bg-white'
        }`}>
          <div className="inline-block min-w-full">
            {/* Column Headers */}
             <div className={`flex sticky top-0 z-20 transition-colors ${
               isDarkMode ? 'bg-gray-900' : 'bg-white'
             }`}>
               <div className={`w-8 h-[18px] border-r border-b flex items-center justify-center transition-colors ${
                 isDarkMode 
                   ? 'bg-gray-800 border-gray-700' 
                   : 'bg-gray-200 border-gray-300'
               }`}>
                 <div className={`w-2 h-2 rounded-sm transition-colors ${
                   isDarkMode ? 'bg-gray-600' : 'bg-gray-400'
                 }`}></div>
               </div>
               {Array.from({ length: COLS }, (_, colIndex) => renderColumnHeader(colIndex))}
             </div>
            
            {/* Grid Rows */}
            {gridData.map((row, rowIndex) => (
              <div key={rowIndex} className="flex">
                {renderRowHeader(rowIndex)}
                {row.map((cell, colIndex) => renderCell(cell, rowIndex, colIndex))}
              </div>
            ))}
          </div>
        </div>

        {/* AI Panel */}
        {showAIPanel && (
          <div className={`w-80 p-4 ml-2 border rounded shadow transition-colors ${
            isDarkMode 
              ? 'bg-gray-800 border-gray-700' 
              : 'bg-white border-gray-200'
          }`}>
            <div className="flex items-center justify-between mb-4">
              <h3 className="font-semibold">AI Assistant</h3>
              <button 
                onClick={() => setShowAIPanel(false)}
                className={`px-2 py-1 text-sm border rounded transition-colors ${
                  isDarkMode 
                    ? 'border-gray-600 hover:bg-gray-700' 
                    : 'border-gray-300 hover:bg-gray-100'
                }`}
              >
                ×
              </button>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className={`block text-sm font-medium mb-2 ${
                  isDarkMode ? 'text-gray-200' : 'text-gray-700'
                }`}>
                  Ask AI to enhance your spreadsheet:
                </label>
                <textarea
                  value={aiPrompt}
                  onChange={(e) => setAiPrompt(e.target.value)}
                  placeholder="e.g., 'Calculate total sales for Q1-Q4'"
                  className={`w-full p-2 border rounded text-sm transition-colors ${
                    isDarkMode 
                      ? 'border-gray-600 bg-gray-700 text-gray-200 placeholder-gray-400' 
                      : 'border-gray-300 bg-white text-gray-700 placeholder-gray-500'
                  }`}
                  rows={3}
                />
              </div>
              
              <button 
                onClick={handleAIPrompt}
                disabled={aiLoading}
                className={`w-full px-3 py-2 text-sm border rounded text-white disabled:opacity-50 transition-colors ${
                  isDarkMode 
                    ? 'bg-blue-600 hover:bg-blue-700 border-blue-500' 
                    : 'bg-blue-500 hover:bg-blue-600 border-blue-400'
                }`}
              >
                {aiLoading ? 'Processing...' : 'Apply AI Enhancement'}
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Hidden File Input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,.xlsx,.json"
        onChange={handleFileUpload}
        className="hidden"
      />
    </div>
    </>
  )
}