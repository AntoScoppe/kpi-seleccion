import { useState, useEffect } from 'react'
import { readSheet } from './api'

export default function App() {
  const [estado, setEstado] = useState('Conectando con Google Sheets...')

  useEffect(() => {
    readSheet('Vacantes')
      .then(rows => setEstado(`✅ Conexión OK — ${rows.length} vacantes encontradas`))
      .catch(() => setEstado('❌ Error de conexión'))
  }, [])

  return (
    <div style={{ fontFamily: 'sans-serif', padding: 40 }}>
      <h1>KPI Selección</h1>
      <p>{estado}</p>
    </div>
  )
}
