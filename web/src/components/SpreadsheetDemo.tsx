'use client'

import React, { useState, useCallback } from 'react'
import { Button } from '@/components'

interface Cell {
  value: string
  formula?: string
  type?: 'text' | 'number' | 'formula'
}

interface ChatMessage {
  id: string
  type: 'user' | 'assistant'
  content: string
  timestamp: Date
}

const SpreadsheetDemo: React.FC = () => {
  const [cells, setCells] = useState<{ [key: string]: Cell }>({})
  const [selectedCell, setSelectedCell] = useState<string>('A1')
  const [isProcessing, setIsProcessing] = useState(false)
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>([
    {
      id: '1',
      type: 'assistant',
      content: 'Hi! I\'m your AI Excel agent. I can help you build financial models, analyze data, create projections, and much more. What would you like to work on?',
      timestamp: new Date()
    }
  ])
  const [chatInput, setChatInput] = useState('')
  const [isChatOpen, setIsChatOpen] = useState(true)

  // Generate column letters (A, B, C, ..., Z, AA, AB, etc.)
  const getColumnLetter = (index: number): string => {
    let result = ''
    while (index >= 0) {
      result = String.fromCharCode(65 + (index % 26)) + result
      index = Math.floor(index / 26) - 1
    }
    return result
  }

  // Generate cell reference (A1, B2, etc.)
  const getCellRef = (row: number, col: number): string => {
    return `${getColumnLetter(col)}${row + 1}`
  }

  // Handle cell value change
  const handleCellChange = useCallback((cellRef: string, value: string) => {
    setCells(prev => ({
      ...prev,
      [cellRef]: {
        value,
        type: value.startsWith('=') ? 'formula' : isNaN(Number(value)) ? 'text' : 'number'
      }
    }))
  }, [])

  // Simulate AI processing for chat
  const handleChatSubmit = async () => {
    if (!chatInput.trim()) return

    const userMessage: ChatMessage = {
      id: Date.now().toString(),
      type: 'user',
      content: chatInput,
      timestamp: new Date()
    }

    setChatMessages(prev => [...prev, userMessage])
    setChatInput('')
    setIsProcessing(true)

    // Simulate AI processing delay
    await new Promise(resolve => setTimeout(resolve, 1500))

    // Generate AI response based on input
    let aiResponse = ''
    const input = chatInput.toLowerCase()

    if (input.includes('financial model') || input.includes('dcf') || input.includes('cash flow')) {
      aiResponse = 'I\'ll create a DCF financial model for you. Let me populate the spreadsheet with revenue projections, operating expenses, and cash flow calculations.'
      
      // Populate spreadsheet with financial model
      const financialData = {
        'A1': { value: 'DCF Financial Model', type: 'text' as const },
        'A3': { value: 'Year', type: 'text' as const },
        'B3': { value: '2024', type: 'text' as const },
        'C3': { value: '2025', type: 'text' as const },
        'D3': { value: '2026', type: 'text' as const },
        'E3': { value: '2027', type: 'text' as const },
        'A4': { value: 'Revenue', type: 'text' as const },
        'B4': { value: '1000000', type: 'number' as const },
        'C4': { value: '1200000', type: 'number' as const },
        'D4': { value: '1440000', type: 'number' as const },
        'E4': { value: '1728000', type: 'number' as const },
        'A5': { value: 'Operating Expenses', type: 'text' as const },
        'B5': { value: '600000', type: 'number' as const },
        'C5': { value: '720000', type: 'number' as const },
        'D5': { value: '864000', type: 'number' as const },
        'E5': { value: '1036800', type: 'number' as const },
        'A6': { value: 'EBITDA', type: 'text' as const },
        'B6': { value: '=B4-B5', type: 'formula' as const },
        'C6': { value: '=C4-C5', type: 'formula' as const },
        'D6': { value: '=D4-D5', type: 'formula' as const },
        'E6': { value: '=E4-E5', type: 'formula' as const },
      }
      setCells(prev => ({ ...prev, ...financialData }))
    } else if (input.includes('budget') || input.includes('expense')) {
      aiResponse = 'I\'ll create a budget template with expense categories and monthly tracking.'
      
      const budgetData = {
        'A1': { value: 'Monthly Budget Template', type: 'text' as const },
        'A3': { value: 'Category', type: 'text' as const },
        'B3': { value: 'Budgeted', type: 'text' as const },
        'C3': { value: 'Actual', type: 'text' as const },
        'D3': { value: 'Variance', type: 'text' as const },
        'A4': { value: 'Housing', type: 'text' as const },
        'B4': { value: '2000', type: 'number' as const },
        'C4': { value: '1950', type: 'number' as const },
        'D4': { value: '=C4-B4', type: 'formula' as const },
        'A5': { value: 'Transportation', type: 'text' as const },
        'B5': { value: '500', type: 'number' as const },
        'C5': { value: '520', type: 'number' as const },
        'D5': { value: '=C5-B5', type: 'formula' as const },
        'A6': { value: 'Food', type: 'text' as const },
        'B6': { value: '600', type: 'number' as const },
        'C6': { value: '580', type: 'number' as const },
        'D6': { value: '=C6-B6', type: 'formula' as const },
      }
      setCells(prev => ({ ...prev, ...budgetData }))
    } else if (input.includes('sales') || input.includes('revenue')) {
      aiResponse = 'I\'ll generate a sales forecast with growth projections and seasonal adjustments.'
      
      const salesData = {
        'A1': { value: 'Sales Forecast', type: 'text' as const },
        'A3': { value: 'Month', type: 'text' as const },
        'B3': { value: 'Base Sales', type: 'text' as const },
        'C3': { value: 'Growth %', type: 'text' as const },
        'D3': { value: 'Projected Sales', type: 'text' as const },
        'A4': { value: 'Jan', type: 'text' as const },
        'B4': { value: '50000', type: 'number' as const },
        'C4': { value: '0.05', type: 'number' as const },
        'D4': { value: '=B4*(1+C4)', type: 'formula' as const },
        'A5': { value: 'Feb', type: 'text' as const },
        'B5': { value: '52000', type: 'number' as const },
        'C5': { value: '0.07', type: 'number' as const },
        'D5': { value: '=B5*(1+C5)', type: 'formula' as const },
      }
      setCells(prev => ({ ...prev, ...salesData }))
    } else {
      aiResponse = 'I can help you with financial modeling, data analysis, budget creation, sales forecasting, and much more. Try asking me to "create a DCF model" or "build a budget template".'
    }

    const assistantMessage: ChatMessage = {
      id: (Date.now() + 1).toString(),
      type: 'assistant',
      content: aiResponse,
      timestamp: new Date()
    }

    setChatMessages(prev => [...prev, assistantMessage])
    setIsProcessing(false)
  }

  // Render spreadsheet grid
  const renderGrid = () => {
    const rows = 20
    const cols = 10

    return (
      <div className="overflow-auto border border-gray-300">
        {/* Header row with column letters */}
        <div className="flex bg-gray-100 sticky top-0 z-10">
          <div className="w-12 h-8 border-r border-gray-300 bg-gray-200"></div>
          {Array.from({ length: cols }, (_, colIndex) => (
            <div
              key={colIndex}
              className="w-24 h-8 border-r border-gray-300 flex items-center justify-center text-sm font-medium bg-gray-100"
            >
              {getColumnLetter(colIndex)}
            </div>
          ))}
        </div>

        {/* Data rows */}
        {Array.from({ length: rows }, (_, rowIndex) => (
          <div key={rowIndex} className="flex">
            {/* Row number */}
            <div className="w-12 h-8 border-r border-b border-gray-300 flex items-center justify-center text-sm font-medium bg-gray-100">
              {rowIndex + 1}
            </div>
            
            {/* Data cells */}
            {Array.from({ length: cols }, (_, colIndex) => {
              const cellRef = getCellRef(rowIndex, colIndex)
              const cell = cells[cellRef]
              const isSelected = selectedCell === cellRef

              return (
                <div
                  key={cellRef}
                  className={`w-24 h-8 border-r border-b border-gray-300 ${
                    isSelected ? 'bg-blue-100 border-blue-500' : 'bg-white hover:bg-gray-50'
                  }`}
                  onClick={() => setSelectedCell(cellRef)}
                >
                  <input
                    type="text"
                    value={cell?.value || ''}
                    onChange={(e) => handleCellChange(cellRef, e.target.value)}
                    className="w-full h-full px-1 text-sm border-none outline-none bg-transparent"
                    placeholder=""
                  />
                </div>
              )
            })}
          </div>
        ))}
      </div>
    )
  }

  return (
    <div className="flex h-screen bg-gray-50">
      {/* Main spreadsheet area */}
      <div className={`flex-1 flex flex-col ${isChatOpen ? 'mr-80' : ''} transition-all duration-300`}>
        {/* Toolbar */}
        <div className="bg-white border-b border-gray-300 p-2 flex items-center gap-4">
          <div className="text-lg font-semibold">Excel AI Agent Demo</div>
          <div className="flex items-center gap-2">
            <span className="text-sm text-gray-600">Selected:</span>
            <span className="font-mono text-sm bg-gray-100 px-2 py-1 rounded">{selectedCell}</span>
          </div>
          <Button
            onClick={() => setIsChatOpen(!isChatOpen)}
            variant="outline"
            size="sm"
          >
            {isChatOpen ? 'Hide AI Chat' : 'Show AI Chat'}
          </Button>
        </div>

        {/* Formula bar */}
        <div className="bg-white border-b border-gray-300 p-2">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium w-12">{selectedCell}</span>
            <input
              type="text"
              value={cells[selectedCell]?.value || ''}
              onChange={(e) => handleCellChange(selectedCell, e.target.value)}
              className="flex-1 px-2 py-1 border border-gray-300 rounded text-sm"
              placeholder="Enter value or formula..."
            />
          </div>
        </div>

        {/* Spreadsheet grid */}
        <div className="flex-1 p-4">
          {renderGrid()}
        </div>
      </div>

      {/* AI Chat Sidebar */}
      {isChatOpen && (
        <div className="w-80 bg-white border-l border-gray-300 flex flex-col fixed right-0 top-0 h-full">
          {/* Chat header */}
          <div className="p-4 border-b border-gray-300">
            <h3 className="font-semibold text-lg">AI Excel Agent</h3>
            <p className="text-sm text-gray-600">Ask me to build models, analyze data, or create templates</p>
          </div>

          {/* Chat messages */}
          <div className="flex-1 overflow-y-auto p-4 space-y-4">
            {chatMessages.map((message) => (
              <div
                key={message.id}
                className={`flex ${message.type === 'user' ? 'justify-end' : 'justify-start'}`}
              >
                <div
                  className={`max-w-[80%] p-3 rounded-lg ${
                    message.type === 'user'
                      ? 'bg-blue-500 text-white'
                      : 'bg-gray-100 text-gray-900'
                  }`}
                >
                  <p className="text-sm">{message.content}</p>
                </div>
              </div>
            ))}
            {isProcessing && (
              <div className="flex justify-start">
                <div className="bg-gray-100 text-gray-900 p-3 rounded-lg">
                  <div className="flex items-center gap-2">
                    <div className="animate-spin w-4 h-4 border-2 border-blue-500 border-t-transparent rounded-full"></div>
                    <span className="text-sm">AI is working...</span>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Chat input */}
          <div className="p-4 border-t border-gray-300">
            <div className="flex gap-2">
              <input
                type="text"
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                onKeyPress={(e) => e.key === 'Enter' && handleChatSubmit()}
                placeholder="Ask me to create a financial model..."
                className="flex-1 px-3 py-2 border border-gray-300 rounded-lg text-sm"
                disabled={isProcessing}
              />
              <Button
                onClick={handleChatSubmit}
                disabled={isProcessing || !chatInput.trim()}
                size="sm"
              >
                Send
              </Button>
            </div>
            <div className="mt-2 text-xs text-gray-500">
              Try: "Create a DCF model", "Build a budget template", "Generate sales forecast"
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default SpreadsheetDemo