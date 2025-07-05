import ReactDOM from 'react-dom/client'
import App from './App'
import './style.css'

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('Excel Add-in is ready!')
    
    // Initialize React app
    const root = ReactDOM.createRoot(document.getElementById('app')!)
    root.render(<App />)
  }
})