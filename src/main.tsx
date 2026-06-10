import React from 'react'
import ReactDOM from 'react-dom/client'
import axios from 'axios'
import App from './App'
import './index.css'

// In production the frontend is on GitHub Pages and the API is on Render — different origins.
// Setting baseURL here so all axios calls prepend the Render URL in production.
// In dev the value is empty and the Vite proxy handles /api/* → localhost:3001.
axios.defaults.baseURL = import.meta.env.VITE_API_URL || ''
// Required for cross-origin cookie (giw_token) to be sent in production.
axios.defaults.withCredentials = true

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
