import { useRef, useState } from 'react'
import { UploadCloud, FileText, X } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, CardContent } from '@/components/ui/card'
import { cn } from '@/lib/utils'

const MAX_MB = 10

export default function UploadZone({ onConvert, loading }) {
  const [file, setFile]       = useState(null)
  const [dragOver, setDragOver] = useState(false)
  const inputRef              = useRef()

  function handleFile(f) {
    if (!f?.name.toLowerCase().endsWith('.pdf')) return alert('Selecione um arquivo .pdf')
    if (f.size > MAX_MB * 1024 * 1024) return alert(`O arquivo excede ${MAX_MB} MB.`)
    setFile(f)
  }

  function formatSize(b) {
    return b < 1024 * 1024 ? `${(b / 1024).toFixed(1)} KB` : `${(b / (1024 * 1024)).toFixed(2)} MB`
  }

  return (
    <div className="space-y-4">
      {!file ? (
        <Card
          className={cn(
            'cursor-pointer border-2 border-dashed transition-colors duration-150',
            dragOver ? 'border-primary bg-primary/5' : 'border-border hover:border-primary/50'
          )}
          onClick={() => inputRef.current?.click()}
          onDragOver={e => { e.preventDefault(); setDragOver(true) }}
          onDragLeave={() => setDragOver(false)}
          onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]) }}
          role="button"
          tabIndex={0}
          onKeyDown={e => e.key === 'Enter' && inputRef.current?.click()}
          aria-label="Área de upload de PDF"
        >
          <CardContent className="flex flex-col items-center gap-3 py-12">
            <UploadCloud className="h-12 w-12 text-text-muted" />
            <p className="text-sm text-text-muted">
              Arraste o PDF aqui ou{' '}
              <span className="text-primary underline underline-offset-2">clique para selecionar</span>
            </p>
            <p className="text-xs text-text-muted">Apenas .pdf · Máx. {MAX_MB} MB</p>
          </CardContent>
        </Card>
      ) : (
        <Card className="border-success/30 bg-success/5">
          <CardContent className="flex items-center gap-3 py-4">
            <FileText className="h-6 w-6 shrink-0 text-success" />
            <div className="flex-1 min-w-0">
              <p className="text-sm font-medium truncate">{file.name}</p>
              <p className="text-xs text-text-muted">{formatSize(file.size)}</p>
            </div>
            <button
              onClick={() => setFile(null)}
              className="rounded p-1 text-text-muted hover:text-danger transition-colors"
              aria-label="Remover arquivo"
            >
              <X className="h-4 w-4" />
            </button>
          </CardContent>
        </Card>
      )}

      <input ref={inputRef} type="file" accept=".pdf" className="hidden" onChange={e => handleFile(e.target.files[0])} />

      <Button
        className="w-full"
        disabled={!file || loading}
        onClick={() => file && onConvert(file)}
      >
        {loading ? 'Processando…' : 'Converter PDF'}
      </Button>
    </div>
  )
}
