const API_URL = 'https://script.google.com/macros/s/AKfycbwSzif5rbsKcuGAG0bXSpsHzP1nXdev95Kxs4y1L17ssRkW6TlqnX7nInSOuf0hMCDemQ/exec'

const FECHA_KEYS = ['fechaSolicitud','fechaAprobacion','fechaInicio','fechaLimite',
  'fechaSeleccion','fechaIncorpPrev','fechaIncorpReal','fechaFactura']

function fixFecha(v) {
  if (!v || v === '') return ''
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v
  const d = new Date(v)
  if (!isNaN(d.getTime())) {
    return d.getFullYear() + '-' +
      String(d.getMonth()+1).padStart(2,'0') + '-' +
      String(d.getDate()).padStart(2,'0')
  }
  return v
}

function fixRow(row) {
  const r = {...row}
  FECHA_KEYS.forEach(k => { if (r[k]) r[k] = fixFecha(r[k]) })
  return r
}

async function call(params) {
  const url = new URL(API_URL)
  Object.entries(params).forEach(([k,v]) => url.searchParams.append(k,v))
  const res = await fetch(url.toString())
  return res.json()
}

export async function readSheet(sheet) {
  const data = await call({ action:'read', sheet })
  return (data.rows || []).map(fixRow)
}

export async function writeRow(sheet, row) {
  return call({ action:'write', sheet, data:JSON.stringify(row) })
}

export async function deleteRow(sheet, id) {
  return call({ action:'delete', sheet, id })
}
