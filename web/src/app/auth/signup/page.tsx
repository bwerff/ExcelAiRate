'use client'

import { useState } from 'react'
import Link from 'next/link'
import { Button, Card, CardContent, CardHeader, CardTitle, Input, Alert } from '../../../../shared/components'
import { useAuthStore } from '../../../store/auth'

export default function SignUp() {
  const [email, setEmail] = useState('')
  const [submitted, setSubmitted] = useState(false)
  const { signIn, loading, error } = useAuthStore()

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    
    try {
      await signIn(email)
      setSubmitted(true)
    } catch (err) {
      // Error is handled in the store
    }
  }

  if (submitted) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
        <Card className="w-full max-w-md">
          <CardHeader>
            <CardTitle>Check Your Email</CardTitle>
          </CardHeader>
          <CardContent>
            <Alert variant="success">
              We've sent a magic link to <strong>{email}</strong>. 
              Click the link in the email to complete your signup.
            </Alert>
            <p className="text-sm text-gray-600 mt-4">
              Didn't receive the email? Check your spam folder or{' '}
              <button
                onClick={() => {
                  setSubmitted(false)
                  setEmail('')
                }}
                className="text-blue-600 hover:underline"
              >
                try again
              </button>
            </p>
          </CardContent>
        </Card>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
      <Card className="w-full max-w-md">
        <CardHeader>
          <CardTitle>Create Your Account</CardTitle>
        </CardHeader>
        <CardContent>
          <form onSubmit={handleSubmit} className="space-y-4">
            <Input
              type="email"
              label="Email"
              placeholder="you@example.com"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              required
            />
            
            {error && (
              <Alert variant="error">
                {error}
              </Alert>
            )}
            
            <Button
              type="submit"
              loading={loading}
              disabled={!email}
              className="w-full"
            >
              Create Account
            </Button>
          </form>
          
          <div className="mt-6">
            <p className="text-xs text-gray-600 text-center mb-4">
              By signing up, you agree to our{' '}
              <Link href="/terms" className="text-blue-600 hover:underline">
                Terms of Service
              </Link>{' '}
              and{' '}
              <Link href="/privacy" className="text-blue-600 hover:underline">
                Privacy Policy
              </Link>
            </p>
            
            <div className="text-center text-sm text-gray-600">
              Already have an account?{' '}
              <Link href="/auth/signin" className="text-blue-600 hover:underline">
                Sign in
              </Link>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  )
}