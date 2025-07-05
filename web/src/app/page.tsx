'use client'

import { useState } from 'react'
import Link from 'next/link'
import { Button, Card, CardContent, CardHeader, CardTitle } from '../../shared/components'
import { useAuthStore } from '../store/auth'
import { PRICING_PLANS } from '../../lib/pricing'

export default function Home() {
  const { user } = useAuthStore()
  const [loading, setLoading] = useState(false)

  const handleGetStarted = async () => {
    if (user) {
      window.location.href = '/dashboard'
    } else {
      window.location.href = '/auth/signin'
    }
  }

  const features = [
    {
      title: 'Smart Analysis',
      description: 'Get instant insights from your Excel data with AI-powered analysis',
      icon: '📊'
    },
    {
      title: 'Data Generation',
      description: 'Generate realistic test data or projections based on your requirements',
      icon: '🔮'
    },
    {
      title: 'Plain English Explanations',
      description: 'Understand complex data patterns with simple, clear explanations',
      icon: '💬'
    }
  ]

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <div className="flex items-center">
              <h1 className="text-2xl font-bold text-gray-900">Excel AI Assistant</h1>
            </div>
            <nav className="flex items-center space-x-4">
              {user ? (
                <>
                  <Link href="/dashboard">
                    <Button variant="ghost">Dashboard</Button>
                  </Link>
                  <Link href="/account">
                    <Button variant="ghost">Account</Button>
                  </Link>
                </>
              ) : (
                <>
                  <Link href="/auth/signin">
                    <Button variant="ghost">Sign In</Button>
                  </Link>
                  <Link href="/auth/signup">
                    <Button>Get Started</Button>
                  </Link>
                </>
              )}
            </nav>
          </div>
        </div>
      </header>

      {/* Hero Section */}
      <section className="py-20 px-4">
        <div className="max-w-4xl mx-auto text-center">
          <h2 className="text-5xl font-bold text-gray-900 mb-6">
            Transform Your Excel Experience with AI
          </h2>
          <p className="text-xl text-gray-600 mb-8">
            Analyze data, generate insights, and get explanations in plain English.
            All within your familiar Excel environment.
          </p>
          <Button size="lg" onClick={handleGetStarted} loading={loading}>
            Get Started for Free
          </Button>
          <p className="mt-4 text-sm text-gray-500">
            No credit card required • 10 free queries
          </p>
        </div>
      </section>

      {/* Features Section */}
      <section className="py-20 px-4 bg-white">
        <div className="max-w-6xl mx-auto">
          <h3 className="text-3xl font-bold text-center text-gray-900 mb-12">
            Powerful AI Features for Excel
          </h3>
          <div className="grid md:grid-cols-3 gap-8">
            {features.map((feature, i) => (
              <Card key={i}>
                <CardHeader>
                  <div className="text-4xl mb-4">{feature.icon}</div>
                  <CardTitle>{feature.title}</CardTitle>
                </CardHeader>
                <CardContent>
                  <p className="text-gray-600">{feature.description}</p>
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      </section>

      {/* Pricing Section */}
      <section className="py-20 px-4">
        <div className="max-w-6xl mx-auto">
          <h3 className="text-3xl font-bold text-center text-gray-900 mb-12">
            Simple, Transparent Pricing
          </h3>
          <div className="grid md:grid-cols-3 gap-8">
            {Object.entries(PRICING_PLANS).map(([key, plan]) => (
              <Card key={key} className={key === 'pro' ? 'border-blue-500 border-2' : ''}>
                <CardHeader>
                  <CardTitle className="text-2xl">
                    {key.charAt(0).toUpperCase() + key.slice(1)}
                  </CardTitle>
                  <p className="text-3xl font-bold mt-2">
                    ${plan.price}
                    {plan.price > 0 && <span className="text-lg font-normal">/month</span>}
                  </p>
                </CardHeader>
                <CardContent>
                  <ul className="space-y-3">
                    <li className="flex items-start">
                      <span className="text-green-500 mr-2">✓</span>
                      <span>{plan.limits.queries} AI queries per month</span>
                    </li>
                    <li className="flex items-start">
                      <span className="text-green-500 mr-2">✓</span>
                      <span>All AI features included</span>
                    </li>
                    {key !== 'free' && (
                      <li className="flex items-start">
                        <span className="text-green-500 mr-2">✓</span>
                        <span>Priority support</span>
                      </li>
                    )}
                    {key === 'team' && (
                      <>
                        <li className="flex items-start">
                          <span className="text-green-500 mr-2">✓</span>
                          <span>Team collaboration</span>
                        </li>
                        <li className="flex items-start">
                          <span className="text-green-500 mr-2">✓</span>
                          <span>Advanced analytics</span>
                        </li>
                      </>
                    )}
                  </ul>
                  <Button 
                    className="w-full mt-6" 
                    variant={key === 'pro' ? 'primary' : 'secondary'}
                    onClick={() => window.location.href = user ? `/account/upgrade?plan=${key}` : '/auth/signup'}
                  >
                    {key === 'free' ? 'Get Started' : 'Upgrade'}
                  </Button>
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer className="bg-gray-900 text-white py-12 px-4">
        <div className="max-w-6xl mx-auto">
          <div className="grid md:grid-cols-4 gap-8">
            <div>
              <h4 className="font-semibold mb-4">Excel AI Assistant</h4>
              <p className="text-gray-400 text-sm">
                AI-powered tools to supercharge your Excel productivity
              </p>
            </div>
            <div>
              <h5 className="font-semibold mb-4">Product</h5>
              <ul className="space-y-2 text-sm text-gray-400">
                <li><Link href="/features">Features</Link></li>
                <li><Link href="/pricing">Pricing</Link></li>
                <li><Link href="/docs">Documentation</Link></li>
              </ul>
            </div>
            <div>
              <h5 className="font-semibold mb-4">Support</h5>
              <ul className="space-y-2 text-sm text-gray-400">
                <li><Link href="/help">Help Center</Link></li>
                <li><Link href="/contact">Contact Us</Link></li>
                <li><Link href="/status">Status</Link></li>
              </ul>
            </div>
            <div>
              <h5 className="font-semibold mb-4">Legal</h5>
              <ul className="space-y-2 text-sm text-gray-400">
                <li><Link href="/privacy">Privacy Policy</Link></li>
                <li><Link href="/terms">Terms of Service</Link></li>
              </ul>
            </div>
          </div>
          <div className="mt-8 pt-8 border-t border-gray-800 text-center text-sm text-gray-400">
            © 2025 Excel AI Assistant. All rights reserved.
          </div>
        </div>
      </footer>
    </div>
  )
}