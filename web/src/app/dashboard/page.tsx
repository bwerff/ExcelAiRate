'use client'

import { useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'
import { Button, Card, CardContent, CardHeader, CardTitle, Alert } from '@/components'
import { useAuthStore } from '../../store/auth'
import { supabaseClient } from '@/lib/supabase'

interface UsageLog {
  id: string
  action_type: string
  prompt: string
  response: Record<string, unknown>
  created_at: string
  from_cache: boolean
}

export default function Dashboard() {
  const router = useRouter()
  const { user, checkSession } = useAuthStore()
  const [usageLogs, setUsageLogs] = useState<UsageLog[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    const loadData = async () => {
      await checkSession()
      
      if (!user) {
        router.push('/auth/signin')
        return
      }

      try {
        const { data: logs } = await supabaseClient
          .from('usage_logs')
          .select('*')
          .eq('user_id', user.id)
          .order('created_at', { ascending: false })
          .limit(20)

        if (logs) {
          setUsageLogs(logs)
        }
      } catch (error) {
        console.error('Error loading usage logs:', error)
      } finally {
        setLoading(false)
      }
    }

    loadData()
  }, [checkSession, router, user])

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="animate-spin h-8 w-8 border-4 border-blue-600 border-t-transparent rounded-full"></div>
      </div>
    )
  }

  if (!user) {
    return null
  }

  const usagePercentage = (user.queries_used / user.queries_limit) * 100

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <h1 className="text-2xl font-bold text-gray-900">Dashboard</h1>
            <nav className="flex items-center space-x-4">
              <Button
                variant="ghost"
                onClick={() => router.push('/account')}
              >
                Account
              </Button>
              <Button
                variant="ghost"
                onClick={async () => {
                  await supabaseClient.auth.signOut()
                  router.push('/')
                }}
              >
                Sign Out
              </Button>
            </nav>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className="grid md:grid-cols-3 gap-6 mb-8">
          <Card>
            <CardHeader>
              <CardTitle>Current Plan</CardTitle>
            </CardHeader>
            <CardContent>
              <p className="text-2xl font-bold capitalize">{user.plan}</p>
              <p className="text-gray-600 mt-1">
                {user.plan === 'free' ? (
                  <Button
                    size="sm"
                    variant="secondary"
                    className="mt-2"
                    onClick={() => router.push('/account/upgrade')}
                  >
                    Upgrade Plan
                  </Button>
                ) : (
                  'Active subscription'
                )}
              </p>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Usage This Month</CardTitle>
            </CardHeader>
            <CardContent>
              <p className="text-2xl font-bold">
                {user.queries_used} / {user.queries_limit}
              </p>
              <div className="mt-2 w-full bg-gray-200 rounded-full h-2">
                <div
                  className={`h-2 rounded-full transition-all ${
                    usagePercentage > 80 ? 'bg-red-500' : 'bg-blue-500'
                  }`}
                  style={{ width: `${Math.min(usagePercentage, 100)}%` }}
                />
              </div>
              {usagePercentage > 80 && (
                <p className="text-sm text-red-600 mt-1">
                  Running low on queries
                </p>
              )}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Quick Actions</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-2">
                <Button
                  className="w-full"
                  onClick={() => window.open('https://aka.ms/sideload-addins')}
                >
                  Install Excel Add-in
                </Button>
                <Button
                  variant="secondary"
                  className="w-full"
                  onClick={() => router.push('/docs')}
                >
                  View Documentation
                </Button>
              </div>
            </CardContent>
          </Card>
        </div>

        <Card>
          <CardHeader>
            <CardTitle>Recent Activity</CardTitle>
          </CardHeader>
          <CardContent>
            {usageLogs.length === 0 ? (
              <Alert variant="info">
                No activity yet. Install the Excel add-in to start using AI features.
              </Alert>
            ) : (
              <div className="space-y-4">
                {usageLogs.map((log) => (
                  <div key={log.id} className="border-b pb-4 last:border-0">
                    <div className="flex justify-between items-start">
                      <div className="flex-1">
                        <p className="font-medium capitalize">
                          {log.action_type} Request
                          {log.from_cache && (
                            <span className="ml-2 text-xs bg-green-100 text-green-800 px-2 py-1 rounded">
                              Cached
                            </span>
                          )}
                        </p>
                        <p className="text-sm text-gray-600 mt-1">{log.prompt}</p>
                      </div>
                      <p className="text-sm text-gray-500">
                        {new Date(log.created_at).toLocaleDateString()}
                      </p>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  )
}