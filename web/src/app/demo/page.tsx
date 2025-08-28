'use client';

import { EnhancedSpreadsheet } from '@/components/EnhancedSpreadsheet';
import { useState } from 'react';
import { Moon, Sun } from 'lucide-react';

export default function DemoPage() {
  const [isDarkMode, setIsDarkMode] = useState(true);
  return (
    <div className={`h-screen flex flex-col overflow-hidden transition-colors duration-300 ${
      isDarkMode ? 'bg-gray-900 text-white' : 'bg-gray-50 text-gray-900'
    }`}>
      {/* Minimal Header */}
      <header className={`flex-shrink-0 border-b transition-colors duration-300 ${
        isDarkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'
      }`}>
        <div className="flex justify-between items-center h-12 px-4">
          <div className="flex items-center space-x-3">
            <div className="w-6 h-6 bg-gradient-to-br from-blue-400 to-blue-600 rounded"></div>
            <h1 className="text-lg font-semibold">Excel AI Assistant</h1>
          </div>
          <nav className="flex items-center space-x-4">
            <button
              onClick={() => setIsDarkMode(!isDarkMode)}
              className={`p-2 rounded-lg transition-colors ${
                isDarkMode
                  ? 'hover:bg-gray-700 text-gray-300 hover:text-white'
                  : 'hover:bg-gray-100 text-gray-600 hover:text-gray-900'
              }`}
              aria-label="Toggle theme"
            >
              {isDarkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
            </button>
            <a 
              href="/" 
              className={`transition-colors text-sm ${
                isDarkMode ? 'text-gray-300 hover:text-white' : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              Home
            </a>
            <a 
              href="/auth/signin" 
              className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1 rounded text-sm transition-colors"
            >
              Sign In
            </a>
          </nav>
        </div>
      </header>

      {/* Compact Hero Section */}
      <section className={`flex-shrink-0 border-b transition-colors duration-300 ${
        isDarkMode 
          ? 'bg-gradient-to-r from-gray-800 to-gray-900 border-gray-700' 
          : 'bg-gradient-to-r from-gray-100 to-gray-50 border-gray-200'
      }`}>
        <div className="px-4 py-3">
          <div className="text-center">
            <h2 className="text-xl font-bold mb-1">
              Enhanced Excel Experience
            </h2>
            <p className={`text-sm ${
              isDarkMode ? 'text-gray-300' : 'text-gray-600'
            }`}>
              Build spreadsheets with 100% Excel compatibility and AI-powered enhancements
            </p>
          </div>
        </div>
      </section>

      {/* Excel Grid - 95% of remaining space */}
      <main className="flex-1 min-h-0 p-0 m-0">
        <div className="h-full w-full">
          <EnhancedSpreadsheet className="h-full border-0 rounded-none" isDarkMode={isDarkMode} />
        </div>
      </main>
    </div>
  );
}