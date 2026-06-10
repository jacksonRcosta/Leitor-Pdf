import { runConverter } from '../../api/_logic.js'

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
}

export const handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: CORS, body: '' }
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify({ erro: 'Método não permitido.' }),
    }
  }

  try {
    const body   = JSON.parse(event.body ?? '{}')
    const result = await runConverter(body)
    return {
      statusCode: 200,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify(result),
    }
  } catch (err) {
    return {
      statusCode: err.status ?? 500,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify({ erro: err.message }),
    }
  }
}
