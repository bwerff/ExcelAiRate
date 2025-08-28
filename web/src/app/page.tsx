'use client'

import { useState } from 'react'
import Link from 'next/link'
import { Button, Card, CardContent, CardHeader, CardTitle } from '@/components'
import { useAuthStore } from '@/store/auth'
import { PRICING_PLANS } from '@/lib/pricing'

export default function Home() {
  const { user } = useAuthStore()
  const [loading] = useState(false)

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
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-slate-100">
      {/* Header */}
      <header className="bg-white/80 backdrop-blur-md border-b border-slate-200/50 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <div className="flex items-center">
              <h1 className="text-2xl font-bold bg-gradient-to-r from-slate-900 to-slate-700 bg-clip-text text-transparent">
                ExcelAI Rate
              </h1>
            </div>
            <nav className="flex items-center space-x-2">
              <Link href="/demo">
                <Button variant="ghost" className="text-slate-700 hover:text-slate-900 hover:bg-slate-100/50">
                  Try Demo
                </Button>
              </Link>
              {user ? (
                <>
                  <Link href="/dashboard">
                    <Button variant="ghost" className="text-slate-700 hover:text-slate-900 hover:bg-slate-100/50">
                      Dashboard
                    </Button>
                  </Link>
                  <Link href="/account">
                    <Button variant="ghost" className="text-slate-700 hover:text-slate-900 hover:bg-slate-100/50">
                      Account
                    </Button>
                  </Link>
                </>
              ) : (
                <>
                  <Link href="/auth/signin">
                    <Button variant="ghost" className="text-slate-700 hover:text-slate-900 hover:bg-slate-100/50">
                      Sign In
                    </Button>
                  </Link>
                  <Link href="/auth/signup">
                    <Button className="bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 text-white shadow-lg">
                      Get Started
                    </Button>
                  </Link>
                </>
              )}
            </nav>
          </div>
        </div>
      </header>

      {/* Hero Section */}
      <section className="py-24 px-6 relative overflow-hidden">
        <div className="absolute inset-0 bg-gradient-to-r from-blue-50/50 to-indigo-50/50 -skew-y-1 transform origin-top-left"></div>
        <div className="max-w-5xl mx-auto text-center relative z-10">
          <h2 className="text-6xl font-bold bg-gradient-to-r from-slate-900 via-slate-800 to-slate-700 bg-clip-text text-transparent mb-8 leading-tight">
            Transform Your Excel Experience with AI
          </h2>
          <p className="text-xl text-slate-600 mb-10 max-w-3xl mx-auto leading-relaxed">
            Analyze data, generate insights, and get explanations in plain English.
            All within your familiar Excel environment.
          </p>
          <div className="flex flex-col sm:flex-row gap-4 justify-center">
            <Button 
              size="lg" 
              onClick={handleGetStarted} 
              loading={loading}
              className="bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 text-white shadow-xl px-8 py-4 text-lg"
            >
              Get Started for Free
            </Button>
            <Link href="/demo">
              <Button 
                size="lg" 
                variant="outline"
                className="border-slate-300 text-slate-700 hover:bg-slate-50 hover:border-slate-400 px-8 py-4 text-lg"
              >
                Try Interactive Demo
              </Button>
            </Link>
          </div>
          <p className="mt-6 text-sm text-slate-500">
            No credit card required • 10 free queries
          </p>
        </div>
      </section>

      {/* Features Section */}
      <section className="py-24 px-6 bg-white/50">
        <div className="max-w-6xl mx-auto">
          <h3 className="text-4xl font-bold text-center bg-gradient-to-r from-slate-900 to-slate-700 bg-clip-text text-transparent mb-16">
            Powerful AI Features for Excel
          </h3>
          <div className="grid md:grid-cols-3 gap-8">
            {features.map((feature, i) => (
              <Card key={i} className="border-slate-200/50 bg-white/80 backdrop-blur-sm hover:shadow-xl transition-all duration-300 hover:-translate-y-1">
                <CardHeader className="pb-4">
                  <div className="text-5xl mb-6">{feature.icon}</div>
                  <CardTitle className="text-xl text-slate-900">{feature.title}</CardTitle>
                </CardHeader>
                <CardContent>
                  <p className="text-slate-600 leading-relaxed">{feature.description}</p>
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      </section>

      {/* Pricing Section */}
      <section className="py-24 px-6 bg-gradient-to-br from-slate-50 to-white">
        <div className="max-w-6xl mx-auto">
          <h3 className="text-4xl font-bold text-center bg-gradient-to-r from-slate-900 to-slate-700 bg-clip-text text-transparent mb-16">
            Simple, Transparent Pricing
          </h3>
          <div className="grid md:grid-cols-3 gap-8">
            {Object.entries(PRICING_PLANS).map(([key, plan]) => (
              <Card 
                key={key} 
                className={`border-slate-200/50 bg-white/80 backdrop-blur-sm hover:shadow-xl transition-all duration-300 hover:-translate-y-1 ${
                  key === 'pro' ? 'ring-2 ring-blue-500/50 border-blue-200' : ''
                }`}
              >
                <CardHeader className="pb-4">
                  <CardTitle className="text-2xl text-slate-900">
                    {key.charAt(0).toUpperCase() + key.slice(1)}
                  </CardTitle>
                  <p className="text-4xl font-bold mt-4 text-slate-900">
                    ${(plan as any).price}
                    {(plan as any).price > 0 && <span className="text-lg font-normal text-slate-600">/month</span>}
                  </p>
                </CardHeader>
                <CardContent>
                  <ul className="space-y-4">
                    <li className="flex items-start">
                      <span className="text-green-500 mr-3 text-lg">✓</span>
                      <span className="text-slate-700">{(plan as any).limits.queries} AI queries per month</span>
                    </li>
                    <li className="flex items-start">
                      <span className="text-green-500 mr-3 text-lg">✓</span>
                      <span className="text-slate-700">All AI features included</span>
                    </li>
                    {key !== 'free' && (
                      <li className="flex items-start">
                        <span className="text-green-500 mr-3 text-lg">✓</span>
                        <span className="text-slate-700">Priority support</span>
                      </li>
                    )}
                    {key === 'team' && (
                      <>
                        <li className="flex items-start">
                          <span className="text-green-500 mr-3 text-lg">✓</span>
                          <span className="text-slate-700">Team collaboration</span>
                        </li>
                        <li className="flex items-start">
                          <span className="text-green-500 mr-3 text-lg">✓</span>
                          <span className="text-slate-700">Advanced analytics</span>
                        </li>
                      </>
                    )}
                  </ul>
                  <Button 
                    className={`w-full mt-8 ${
                      key === 'pro' 
                        ? 'bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 text-white shadow-lg' 
                        : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300'
                    }`}
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
      <footer className="bg-slate-900 text-white py-16 px-6">
        <div className="max-w-6xl mx-auto">
          <div className="grid md:grid-cols-4 gap-8">
            <div>
              <h4 className="font-semibold mb-6 text-lg">ExcelAI Rate</h4>
              <p className="text-slate-400 text-sm leading-relaxed">
                AI-powered tools to supercharge your Excel productivity
              </p>
            </div>
            <div>
              <h5 className="font-semibold mb-6 text-slate-300">Product</h5>
              <ul className="space-y-3 text-sm text-slate-400">
                <li><Link href="/features" className="hover:text-white transition-colors">Features</Link></li>
                <li><Link href="/pricing" className="hover:text-white transition-colors">Pricing</Link></li>
                <li><Link href="/docs" className="hover:text-white transition-colors">Documentation</Link></li>
              </ul>
            </div>
            <div>
              <h5 className="font-semibold mb-6 text-slate-300">Support</h5>
              <ul className="space-y-3 text-sm text-slate-400">
                <li><Link href="/help" className="hover:text-white transition-colors">Help Center</Link></li>
                <li><Link href="/contact" className="hover:text-white transition-colors">Contact Us</Link></li>
                <li><Link href="/status" className="hover:text-white transition-colors">Status</Link></li>
              </ul>
            </div>
            <div>
              <h5 className="font-semibold mb-6 text-slate-300">Legal</h5>
              <ul className="space-y-3 text-sm text-slate-400">
                <li><Link href="/privacy" className="hover:text-white transition-colors">Privacy Policy</Link></li>
                <li><Link href="/terms" className="hover:text-white transition-colors">Terms of Service</Link></li>
              </ul>
            </div>
          </div>
          <div className="mt-12 pt-8 border-t border-slate-800 text-center text-sm text-slate-400">
            © 2025 ExcelAI Rate. All rights reserved.
          </div>
        </div>
      </footer>
    </div>
  )
}