import { Download } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Table, TableHeader, TableBody, TableRow, TableHead, TableCell } from '@/components/ui/table'

function downloadBase64(b64, mime, filename) {
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0))
  const url   = URL.createObjectURL(new Blob([bytes], { type: mime }))
  Object.assign(document.createElement('a'), { href: url, download: filename }).click()
  URL.revokeObjectURL(url)
}

export default function ResultTable({ data, fileName, onReset }) {
  const { pedidos, zip, csv, csv_nome } = data
  const zipName = fileName?.replace(/\.pdf$/i, '') + '_convertido.zip' || 'pedidos_convertidos.zip'

  return (
    <div className="space-y-6">
      {/* Cabeçalho de página */}
      <div className="flex items-start justify-between gap-4 flex-wrap">
        <div>
          <h1 className="text-2xl font-semibold text-text">Conversão concluída</h1>
          <p className="text-sm text-text-muted mt-0.5">
            {pedidos.length} pedido{pedidos.length !== 1 ? 's' : ''} extraído{pedidos.length !== 1 ? 's' : ''}
          </p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <Button variant="outline" size="sm" onClick={() => downloadBase64(csv, 'text/csv', csv_nome)}>
            <Download className="h-4 w-4" /> CSV
          </Button>
          <Button size="sm" onClick={() => downloadBase64(zip, 'application/zip', zipName)}>
            <Download className="h-4 w-4" /> ZIP (todos)
          </Button>
        </div>
      </div>

      {/* Tabela */}
      <div className="rounded-lg border border-border bg-surface shadow-sm">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Loja</TableHead>
              <TableHead>Razão Social</TableHead>
              <TableHead className="font-mono">Nº Pedido</TableHead>
              <TableHead>Data Compra</TableHead>
              <TableHead>Data Entrega</TableHead>
              <TableHead className="text-right">Itens</TableHead>
              <TableHead>Downloads</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {pedidos.map(p => (
              <TableRow key={p.loja}>
                <TableCell>
                  <Badge>{p.loja}</Badge>
                </TableCell>
                <TableCell className="max-w-[200px] truncate text-sm">{p.razao_social}</TableCell>
                <TableCell className="font-mono text-xs">{p.num_pedido || '—'}</TableCell>
                <TableCell className="text-sm">{p.data_compra || '—'}</TableCell>
                <TableCell className="text-sm">{p.data_entrega || '—'}</TableCell>
                <TableCell className="text-right font-mono text-sm">{p.total_itens}</TableCell>
                <TableCell>
                  <div className="flex items-center gap-1.5">
                    <Button
                      variant="outline" size="sm"
                      className="h-7 px-2 text-xs border-success/40 text-success hover:bg-success/5"
                      onClick={() => downloadBase64(p.xls, 'application/vnd.ms-excel', p.xls_nome)}
                      title="XLS padrão (campos selecionados vazios)"
                    >
                      XLS
                    </Button>
                    <Button
                      variant="outline" size="sm"
                      className="h-7 px-2 text-xs border-primary/40 text-primary hover:bg-primary/5"
                      onClick={() => downloadBase64(p.xls_completo, 'application/vnd.ms-excel', p.xls_completo_nome)}
                      title="XLS com todos os campos preenchidos"
                    >
                      XLS+
                    </Button>
                    <Button
                      variant="outline" size="sm"
                      className="h-7 px-2 text-xs border-warning/40 text-warning hover:bg-warning/5"
                      onClick={() => downloadBase64(p.xml, 'application/xml', p.xml_nome)}
                      title="Baixar XML"
                    >
                      XML
                    </Button>
                  </div>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      <Button variant="outline" onClick={onReset}>Nova conversão</Button>
    </div>
  )
}
