const API_URL = 'https://script.google.com/macros/s/AKfycbxs8O0mPHdYJrzQP1p2hic0J43_4BgcFvXBSPpG11Gx5xzJ8CjJ4z_rv-rLujBTDWR5AA/exec'

async function call(params) {
  const url = new URL(API_URL)
  Object.entries(params).forEach(([k, v]) => url.searchParams.append(k, v))
  const res = await fetch(url.toString())
  return res.json()
}

export async function readSheet(sheet) {
  const data = await call({ action: 'read', sheet })
  return data.rows || []
}

export async function writeRow(sheet, row) {
  return call({ action: 'write', sheet, data: JSON.stringify(row) })
}

export async function deleteRow(sheet, id) {
  return call({ action: 'delete', sheet, id })
}
