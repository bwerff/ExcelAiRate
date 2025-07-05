import React, { useState, useCallback } from 'react'
import { Button, Card, CardContent, TextArea, Alert, Input } from './components'
import { aiService } from './services/ai'
import { authService } from './services/auth'
import type { AIResponse } from '../../shared/types'
import './App.css'

type OperationType = 'analyze' | 'generate' | 'explain'

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false)
  const [email, setEmail] = useState('')
  const [operation, setOperation] = useState<OperationType>('analyze')
  const [prompt, setPrompt] = useState('')
  const [selectedData, setSelectedData] = useState('')
  const [result, setResult] = useState<AIResponse | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [authLoading, setAuthLoading] = useState(false)
  const [authMessage, setAuthMessage] = useState('')

  const handleAuth = async () => {
    if (!email) {
      setError('Please enter your email')
      return
    }

    setAuthLoading(true)
    setError('')
    setAuthMessage('')

    try {
      await authService.signIn(email)
      setAuthMessage('Check your email for the magic link to sign in!')
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to send magic link')
    } finally {
      setAuthLoading(false)
    }
  }

  const getSelectedData = useCallback(async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange()
        range.load(['values', 'address'])
        await context.sync()
        
        const values = range.values
        const data = values.map(row => row.join('\t')).join('\n')
        setSelectedData(data)
      })
    } catch (err) {
      setError('Please select data in Excel first')
    }
  }, [])

  const handleOperation = async () => {
    if (!prompt) {
      setError('Please enter a prompt')
      return
    }

    if (operation !== 'generate' && !selectedData) {
      await getSelectedData()
      if (!selectedData) return
    }

    setLoading(true)
    setError('')
    setResult(null)

    try {
      const response = await aiService[operation](prompt, selectedData)
      setResult(response)

      if (operation === 'generate' && response && 'content' in response) {
        await insertGeneratedData(response.content)
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : `Failed to ${operation} data`)
    } finally {
      setLoading(false)
    }
  }

  const insertGeneratedData = async (data: string) => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        const range = context.workbook.getSelectedRange()
        range.load('address')
        await context.sync()

        const rows = data.split('\n').map(row => row.split('\t'))
        const targetRange = sheet.getRangeByIndexes(
          range.rowIndex,
          range.columnIndex,
          rows.length,
          rows[0].length
        )
        
        targetRange.values = rows
        await context.sync()
      })
    } catch (err) {
      setError('Failed to insert generated data')
    }
  }

  React.useEffect(() => {
    Office.onReady(() => {
      const session = authService.getSession()
      setIsAuthenticated(!!session)
    })
  }, [])

  if (!isAuthenticated) {
    return (
      <div className="app-container">
        <Card>
          <CardContent>
            <h1 className="text-xl font-bold mb-4">Excel AI Assistant</h1>
            <p className="text-gray-600 mb-6">
              Sign in with your email to start using AI features
            </p>
            
            <Input
              type="email"
              label="Email"
              placeholder="your@email.com"
              value={email}
              onChange={(e: React.ChangeEvent<HTMLInputElement>) => setEmail(e.target.value)}
              onKeyDown={(e: React.KeyboardEvent) => e.key === 'Enter' && handleAuth()}
            />
            
            <Button
              onClick={handleAuth}
              loading={authLoading}
              className="w-full mt-4"
            >
              Send Magic Link
            </Button>

            {authMessage && (
              <Alert variant="success" className="mt-4">
                {authMessage}
              </Alert>
            )}

            {error && (
              <Alert variant="error" className="mt-4">
                {error}
              </Alert>
            )}
          </CardContent>
        </Card>
      </div>
    )
  }

  return (
    <div className="app-container">
      <Card>
        <CardContent>
          <h1 className="text-xl font-bold mb-4">Excel AI Assistant</h1>
          
          <div className="mb-4">
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Operation
            </label>
            <div className="grid grid-cols-3 gap-2">
              {(['analyze', 'generate', 'explain'] as OperationType[]).map(op => (
                <Button
                  key={op}
                  variant={operation === op ? 'primary' : 'secondary'}
                  size="sm"
                  onClick={() => setOperation(op)}
                >
                  {op.charAt(0).toUpperCase() + op.slice(1)}
                </Button>
              ))}
            </div>
          </div>

          {operation !== 'generate' && (
            <div className="mb-4">
              <Button
                variant="ghost"
                size="sm"
                onClick={getSelectedData}
                className="w-full"
              >
                Get Selected Data
              </Button>
              {selectedData && (
                <div className="mt-2 p-2 bg-gray-50 rounded text-xs text-gray-600 max-h-20 overflow-auto">
                  {selectedData}
                </div>
              )}
            </div>
          )}

          <TextArea
            label="Prompt"
            placeholder={
              operation === 'analyze' 
                ? "What patterns do you see in this data?"
                : operation === 'generate'
                ? "Generate a monthly sales forecast"
                : "Explain what this data represents"
            }
            value={prompt}
            onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setPrompt(e.target.value)}
            rows={3}
          />

          <Button
            onClick={handleOperation}
            loading={loading}
            className="w-full mt-4"
          >
            {operation.charAt(0).toUpperCase() + operation.slice(1)} Data
          </Button>

          {error && (
            <Alert variant="error" className="mt-4">
              {error}
            </Alert>
          )}

          {result && (
            <div className="mt-4">
              <h3 className="font-medium text-gray-900 mb-2">Result:</h3>
              <div className="p-3 bg-gray-50 rounded text-sm">
                {operation === 'analyze' && 'summary' in result ? (
                  <div>
                    <p className="font-medium mb-2">{result.summary}</p>
                    {result.insights && (
                      <>
                        <h4 className="font-medium mt-3 mb-1">Insights:</h4>
                        <ul className="list-disc list-inside space-y-1">
                          {result.insights.map((insight: string, i: number) => (
                            <li key={i}>{insight}</li>
                          ))}
                        </ul>
                      </>
                    )}
                    {result.recommendations && (
                      <>
                        <h4 className="font-medium mt-3 mb-1">Recommendations:</h4>
                        <ul className="list-disc list-inside space-y-1">
                          {result.recommendations.map((rec: string, i: number) => (
                            <li key={i}>{rec}</li>
                          ))}
                        </ul>
                      </>
                    )}
                  </div>
                ) : operation === 'generate' && 'content' in result ? (
                  <pre className="whitespace-pre-wrap">{result.content}</pre>
                ) : operation === 'explain' && 'explanation' in result ? (
                  <div>
                    <p className="mb-2">{result.explanation}</p>
                    {result.examples && (
                      <>
                        <h4 className="font-medium mt-3 mb-1">Examples:</h4>
                        <ul className="list-disc list-inside space-y-1">
                          {result.examples.map((example: string, i: number) => (
                            <li key={i}>{example}</li>
                          ))}
                        </ul>
                      </>
                    )}
                  </div>
                ) : (
                  <pre className="whitespace-pre-wrap">{JSON.stringify(result, null, 2)}</pre>
                )}
              </div>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  )
}