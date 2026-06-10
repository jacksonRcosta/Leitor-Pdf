import { useState } from 'react'
import { FileSearch, Loader2, AlertCircle } from 'lucide-react'
import AppShell from '@/components/AppShell'
import UploadZone from '@/components/UploadZone'
import ResultTable from '@/components/ResultTable'
import { Button } from '@/components/ui/button'
import { Card, CardContent } from '@/components/ui/card'

const FUNCTION_URL = '/.netlify/functions/converter'

const NAV = [
  { label: 'Conversor', href: '#', icon: FileSearch, active: true },
]

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload  = () => resolve(reader.result.split(',')[1])
    reader.onerror = () => reject(new Error('Falha ao ler o arquivo.'))
    reader.readAsDataURL(file)
  })
}

export default function App() {
  const [state,      setState]      = useState('upload')  // upload | loading | result | error
  const [loadingMsg, setLoadingMsg] = useState('')
  const [resultData, setResultData] = useState(null)
  const [errorMsg,   setErrorMsg]   = useState('')
  const [fileName,   setFileName]   = useState('')

  async function handleConvert(file) {
    setFileName(file.name)
    setState('loading')
    setLoadingMsg('Lendo PDF…')

    const controller = new AbortController()
    const timeout    = setTimeout(() => controller.abort(), 55000)

    try {
      const b64 = await fileToBase64(file)
      setLoadingMsg('Processando pedidos…')

      let res
      try {
        res = await fetch(FUNCTION_URL, {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify({ pdf: b64 }),
          signal:  controller.signal,
        })
      } finally {
        clearTimeout(timeout)
      }

      const text = await res.text()

      if (!text?.trim()) {
        throw new Error(
          res.status === 413 ? 'Arquivo muito grande. O limite é 10 MB.' :
          res.status === 504 || res.status === 524 ? 'Tempo de processamento excedido.' :
          `Sem resposta do servidor (HTTP ${res.status}).`
        )
      }

      let data
      try { data = JSON.parse(text) }
      catch { throw new Error('Resposta inválida do servidor.') }

      if (!res.ok) throw new Error(data.erro || `Erro ${res.status}`)

      setResultData(data)
      setState('result')

    } catch (err) {
      setErrorMsg(
        err.name === 'AbortError'
          ? 'Tempo limite atingido (55s). Tente com um PDF menor.'
          : err.message || 'Erro inesperado. Tente novamente.'
      )
      setState('error')
    }
  }

  function reset() {
    setState('upload')
    setResultData(null)
    setErrorMsg('')
  }

  return (
    <AppShell appName="Conversor de Pedidos" nav={NAV}>
      {state === 'upload' && (
        <div className="max-w-lg mx-auto space-y-6">
          <div>
            <h1 className="text-2xl font-semibold text-text">Converter PDF</h1>
            <p className="text-sm text-text-muted mt-0.5">
              Selecione um pedido de compra em PDF para gerar XLS, XML e CSV.
            </p>
          </div>
          <UploadZone onConvert={handleConvert} loading={false} />
        </div>
      )}

      {state === 'loading' && (
        <div className="flex flex-col items-center justify-center py-24 gap-4">
          <Loader2 className="h-10 w-10 animate-spin text-primary" />
          <p className="text-sm text-text-muted">{loadingMsg}</p>
        </div>
      )}

      {state === 'result' && resultData && (
        <ResultTable data={resultData} fileName={fileName} onReset={reset} />
      )}

      {state === 'error' && (
        <div className="flex flex-col items-center justify-center py-24 gap-4 text-center">
          <AlertCircle className="h-10 w-10 text-danger" />
          <div>
            <p className="font-medium text-text">Não foi possível processar o PDF</p>
            <p className="text-sm text-text-muted mt-1">{errorMsg}</p>
          </div>
          <Button variant="outline" onClick={reset}>Tentar novamente</Button>
        </div>
      )}
    </AppShell>
  )
}
