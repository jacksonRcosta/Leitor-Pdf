import { runConverter } from './_logic.js'

// 10 MB PDF → ~13,4 MB base64 — aumenta o limite do body parser do Vercel
export const config = {
  api: {
    bodyParser: {
      sizeLimit: '15mb',
    },
  },
}

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
}

export default async function handler(req, res) {
  Object.entries(CORS).forEach(([k, v]) => res.setHeader(k, v))

  if (req.method === 'OPTIONS') return res.status(200).end()
  if (req.method !== 'POST')
    return res.status(405).json({ erro: 'Método não permitido.' })

  try {
    const result = await runConverter(req.body)
    return res.status(200).json(result)
  } catch (err) {
    return res.status(err.status ?? 500).json({ erro: err.message })
  }
}
