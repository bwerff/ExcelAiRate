'use client'

import { useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'
import { Button, Card, CardContent, CardHeader, CardTitle, Alert } from '@/components'
import { useAuthStore } from '../../store/auth'
import { stripe } from '@/lib/stripe'
import { PRICING_PLANS } from '@/lib/pricing'
import { supabaseClient } from '@/lib/supabase'

export default function Account() {
  const router = useRouter()
  const { user, checkSession } = useAuthStore()
  const [loading, setLoading] = useState(false)
  const [portalLoading, setPortalLoading] = useState(false)

  useEffect(() => {
    const loadData = async () => {
      await checkSession()
      
      if (!user) {
        router.push('/auth/signin')
      }
    }

    loadData()
  }, [checkSession, router, user])

  if (!user) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="animate-spin h-8 w-8 border-4 border-blue-600 border-t-transparent rounded-full"></div>
      </div>
    )
  }

  const handleUpgrade = async (plan: 'pro' | 'team') => {
    setLoading(true)
    try {
      const { url } = await stripe.createCheckout(plan)
      
      if (url) {
        window.location.href = url
      }
    } catch (error) {
      console.error('Upgrade error:', error)
      alert('Failed to start upgrade process. Please try again.')
    } finally {
      setLoading(false)
    }
  }

  const handleManageSubscription = async () => {
    setPortalLoading(true)
    try {
      const response = await fetch('/api/stripe/portal', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${await (await supabaseClient.auth.getSession()).data.session?.access_token}`
        }
      })
      
      const { url } = await response.json()
      if (url) {
        window.location.href = url
      }
    } catch (error) {
      console.error('Portal error:', error)
      alert('Failed to open billing portal. Please try again.')
    } finally {
      setPortalLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <h1 className="text-2xl font-bold text-gray-900">Account Settings</h1>
            <nav className="flex items-center space-x-4">
              <Button
                variant="ghost"
                onClick={() => router.push('/dashboard')}
              >
                Back to Dashboard
              </Button>
            </nav>
          </div>
        </div>
      </header>

      <div className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <Card className="mb-6">
          <CardHeader>
            <CardTitle>Account Information</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700">Email</label>
                <p className="mt-1 text-gray-900">{user.email}</p>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">User ID</label>
                <p className="mt-1 text-gray-500 text-sm font-mono">{user.id}</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card className="mb-6">
          <CardHeader>
            <CardTitle>Subscription</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700">Current Plan</label>
                <p className="mt-1 text-2xl font-bold capitalize">{user.plan}</p>
              </div>
              
              <div>
                <label className="block text-sm font-medium text-gray-700">Monthly Usage</label>
                <p className="mt-1">
                  {user.queries_used} / {user.queries_limit} queries used
                </p>
                <div className="mt-2 w-full bg-gray-200 rounded-full h-2">
                  <div
                    className="h-2 bg-blue-500 rounded-full transition-all"
                    style={{ width: `${Math.min((user.queries_used / user.queries_limit) * 100, 100)}%` }}
                  />
                </div>
              </div>

              {user.plan !== 'free' ? (
                <Button
                  onClick={handleManageSubscription}
                  loading={portalLoading}
                  variant="secondary"
                >
                  Manage Subscription
                </Button>
              ) : (
                <div className="pt-4">
                  <h4 className="font-medium mb-4">Upgrade Your Plan</h4>
                  <div className="grid sm:grid-cols-2 gap-4">
                    <Card>
                      <CardHeader>
                        <CardTitle>Pro</CardTitle>
                        <p className="text-2xl font-bold">${PRICING_PLANS.pro.price}/mo</p>
                      </CardHeader>
                      <CardContent>
                        <ul className="space-y-2 text-sm mb-4">
                          <li>✓ {PRICING_PLANS.pro.limits.queries} queries/month</li>
                          <li>✓ Priority support</li>
                          <li>✓ All AI features</li>
                        </ul>
                        <Button
                          onClick={() => handleUpgrade('pro')}
                          loading={loading}
                          className="w-full"
                        >
                          Upgrade to Pro
                        </Button>
                      </CardContent>
                    </Card>

                    <Card>
                      <CardHeader>
                        <CardTitle>Team</CardTitle>
                        <p className="text-2xl font-bold">${PRICING_PLANS.team.price}/mo</p>
                      </CardHeader>
                      <CardContent>
                        <ul className="space-y-2 text-sm mb-4">
                          <li>✓ {PRICING_PLANS.team.limits.queries} queries/month</li>
                          <li>✓ Priority support</li>
                          <li>✓ Team collaboration</li>
                        </ul>
                        <Button
                          onClick={() => handleUpgrade('team')}
                          loading={loading}
                          className="w-full"
                        >
                          Upgrade to Team
                        </Button>
                      </CardContent>
                    </Card>
                  </div>
                </div>
              )}
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Danger Zone</CardTitle>
          </CardHeader>
          <CardContent>
            <Alert variant="warning">
              Deleting your account will permanently remove all your data and cannot be undone.
            </Alert>
            <Button
              variant="secondary"
              className="mt-4"
              onClick={() => {
                if (confirm('Are you sure you want to delete your account? This cannot be undone.')) {
                  // Implement account deletion
                  alert('Please contact support to delete your account.')
                }
              }}
            >
              Delete Account
            </Button>
          </CardContent>
        </Card>
      </div>
    </div>
  )
}