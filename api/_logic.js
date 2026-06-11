import pdfParse from 'pdf-parse'
import XLSX from 'xlsx'
import JSZip from 'jszip'

// ── Layout 5 — ArquivoImportado ───────────────────────────────────────────────
const XLS_HEADERS = [
  'codproduto', 'codembalagem', 'quantidade', 'descricao',
  'emba', 'qtUnit', 'precoVenda', 'preço emba', 'preço emba st',
  'preço unit', 'preço tot', 'preco tot ion', 'preco tot ion st',
]
const XLS_KEYS = [
  'codproduto', 'codembalagem', 'quantidade', 'descricao',
  'emba', 'qtUnit', 'precoVenda', 'preco_emba', 'preco_emba_st',
  'preco_unit', 'preco_tot', 'preco_tot_ion', 'preco_tot_ion_st',
]

// Colunas deixadas em branco na versão XLS padrão (ambos os formatos usam o mesmo layout)
const COLS_VAZIAS_PEDIDO = new Set(['codproduto', 'descricao', 'emba', 'preco_unit', 'preco_tot'])

// ── Utilitários ───────────────────────────────────────────────────────────────
const esc     = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')
const cnpjNum = s => s.replace(/\D/g, '')
const fmtCNPJ = c => /^\d{14}$/.test(c)
  ? c.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5')
  : c

// Remove "R$", converte separador de milhar (ponto) e decimal (vírgula → ponto)
// Exemplos: "R$ 2,00" → "2.00" | "1.234,56" → "1234.56" | "2,00" → "2.00"
const normalizarPreco = v => {
  const s = String(v || '').replace(/R\$\s*/gi, '').trim()
  if (!s) return ''
  return s.replace(/\.(?=\d{3}[,])/g, '').replace(',', '.')
}

// ── Extração de CNPJ do texto ─────────────────────────────────────────────────
function extrairPrimeiroCnpj(texto) {
  const m = texto.match(/\b(\d{2}[\.\s]?\d{3}[\.\s]?\d{3}[\/\s]?\d{4}[-\s]?\d{2})\b/)
  if (!m) return null
  const n = m[1].replace(/\D/g, '')
  return n.length === 14 ? { formatado: fmtCNPJ(n), numerico: n } : null
}

// ── Sanitiza string para tag XML válida ───────────────────────────────────────
function sanitizarTag(s) {
  const t = String(s || 'Campo').replace(/[^a-zA-Z0-9_]/g, '_').replace(/^(\d)/, 'C_$1')
  return t || 'Campo'
}

// ── Remove linhas repetitivas (cabeçalhos/rodapés de página) ─────────────────
function filtrarRepetitivas(linhas) {
  if (linhas.length < 6) return linhas
  const freq = {}
  for (const l of linhas) freq[l] = (freq[l] || 0) + 1
  const limiar = Math.max(2, linhas.length * 0.15)
  return linhas.filter(l => freq[l] < limiar)
}

// ── Detecção de formato ───────────────────────────────────────────────────────
function detectarFormato(text) {
  if (/ruptura\s*[-–]?\s*pr[eé]/i.test(text)) return 'ruptura'
  if (/RAZ[ÃA]O SOCIAL:CNPJ:/i.test(text))    return 'pedido'
  return 'generico'
}

// Sanitiza o nome do arquivo para uso seguro em nomes de arquivos gerados
function sanitizarNome(nome) {
  return (nome || 'documento')
    .replace(/\.pdf$/i, '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim() || 'documento'
}

// ── Parser A: Pedido de Compra ─────────────────────────────────────────────────
// Regex para o formato original (campos colados, ex: "RAZÃO SOCIAL:CNPJ:10.840.716/0009-08I.E.:...")
const RE_RAZAO   = /^(.+?)RAZ[ÃA]O SOCIAL:CNPJ:(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})I\.E\.:(\d+)/
const RE_LOJA    = /LOJA ENTREGA:(\w+)/
const RE_PEDIDO  = /PEDIDO[:\s]*(\d+)/i
const RE_COMPRA  = /DATA COMPRA:(\d{2}\/\d{2}\/\d{4})/
const RE_ENTREGA = /DATA ENTREGA:(\d{2}\/\d{2}\/\d{4})/
const RE_VENC    = /VENCIMENTOS:(\S+)/
const RE_ITEM    = /^(\d{6})(.+?)([A-Z]{2}\/\d{4})(\d{10,12})([\d.]+,\d{2})([\d.]+,\d{2})([\d.]+,\d{2})(\d+)(\d{13,14})([\d.]+,\d{2})/

function extrairPedidos(texto) {
  const pedidos = []
  let atual = null

  for (const linha of texto.split('\n')) {
    const mRazao = RE_RAZAO.exec(linha)
    if (mRazao) {
      if (atual && atual.itens.length) pedidos.push(atual)
      const fmt = mRazao[2].trim()
      atual = {
        razao_social:   mRazao[1].trim(),
        cnpj_formatado: fmt,
        cnpj_numerico:  cnpjNum(fmt),
        ie:             mRazao[3].trim(),
        loja:           '',
        num_pedido:     '',
        data_compra:    '',
        data_entrega:   '',
        vencimentos:    '',
        itens:          [],
      }
    }

    if (atual) {
      if (!atual.loja)         { const m = RE_LOJA.exec(linha);    if (m) atual.loja         = m[1] }
      if (!atual.num_pedido)   { const m = RE_PEDIDO.exec(linha);  if (m) atual.num_pedido   = m[1] }
      if (!atual.data_compra)  { const m = RE_COMPRA.exec(linha);  if (m) atual.data_compra  = m[1] }
      if (!atual.data_entrega) { const m = RE_ENTREGA.exec(linha); if (m) atual.data_entrega = m[1] }
      if (!atual.vencimentos)  { const m = RE_VENC.exec(linha);    if (m) atual.vencimentos  = m[1] }

      const mItem = RE_ITEM.exec(linha.replace(/R\$\s*/g, ''))
      if (mItem) {
        atual.itens.push({
          codproduto:       mItem[1],
          descricao:        mItem[2].trim(),
          emba:             mItem[3],
          externo:          mItem[4],
          precoVenda:       normalizarPreco(mItem[5]),
          frete:            normalizarPreco(mItem[6]),
          preco_tot:        normalizarPreco(mItem[7]),
          quantidade:       mItem[8],
          codembalagem:     mItem[9],
          desconto:         normalizarPreco(mItem[10]),
          qtUnit:           '',
          preco_emba:       '',
          preco_emba_st:    '',
          preco_unit:       normalizarPreco(mItem[5]),
          preco_tot_ion:    '',
          preco_tot_ion_st: '',
        })
      }
    }
  }

  if (atual && atual.itens.length) pedidos.push(atual)
  return pedidos
}

// ── Parser B: Ruptura Pré-Pedido ──────────────────────────────────────────────
// O pdf-parse extrai o texto das tabelas SEM espaços entre colunas. Formato real:
//   "016335 - 01 | CENCOSUD BRASIL ATACADO LTDA.09182947000488090002 - MIOJO...7891079000434000698 - HUMBERTO TEOTONIO DE FIGUEIREDO15009/06/202611/02/2026"
//
// Named capture groups eliminam qualquer risco de off-by-one no destructuring.
// Grupos: codloja(6d) filial(2d) razao cnpj(14d) codprod(6d) desc ean(13d)
//         codvend(6d) vend qtd datarup datault
// Grupo preco é opcional: PDFs sem coluna "Preço de Venda" não o incluem.
// Quando presente, aparece colado ao nome do vendedor (ex: "FIGUEIREDOR$ 2,80150").
const RE_RUPTURA = /^(?<codloja>\d{6}) - (?<filial>\d{2}) \| (?<razao>.+?)(?<cnpj>\d{14})(?<codprod>\d{6}) - (?<desc>.+?)(?<ean>\d{13})(?<codvend>\d{6}) - (?<vend>.+?)(?:R\$\s*(?<preco>[\d.]+,\d{2}))?(?<qtd>\d+)(?<datarup>\d{2}\/\d{2}\/\d{4})(?<datault>\d{2}\/\d{2}\/\d{4})$/

function extrairRupturaPrePedido(texto) {
  const mapa = new Map() // CNPJ → pedido

  for (const linha of texto.split('\n')) {
    const m = RE_RUPTURA.exec(linha.trim())
    if (!m) continue

    const { codloja, filial, razao, cnpj, codprod, desc, ean, preco, qtd, datarup } = m.groups

    if (!mapa.has(cnpj)) {
      mapa.set(cnpj, {
        loja:           `${codloja}-${filial}`,
        razao_social:   razao.trim(),
        cnpj_formatado: fmtCNPJ(cnpj),
        cnpj_numerico:  cnpj,
        num_pedido:     '',
        data_compra:    datarup,
        data_entrega:   '',
        vencimentos:    '',
        itens:          [],
      })
    }

    mapa.get(cnpj).itens.push({
      codproduto:       codprod,
      descricao:        desc.trim().replace(/\s{2,}/g, ' '),
      codembalagem:     ean,
      emba:             '',
      quantidade:       qtd,
      qtUnit:           '',
      precoVenda:       normalizarPreco(preco || ''),
      preco_emba:       '',
      preco_emba_st:    '',
      preco_unit:       '',
      preco_tot:        '',
      preco_tot_ion:    '',
      preco_tot_ion_st: '',
    })
  }

  return [...mapa.values()].filter(p => p.itens.length > 0)
}

// ── Geradores (compartilhados) ────────────────────────────────────────────────
function _montarXls(pedido, completo, colsVazias) {
  const wb = XLSX.utils.book_new()
  const wsData = [
    ['cnpj', pedido.cnpj_numerico],
    XLS_HEADERS,
    ...pedido.itens.map(item =>
      XLS_KEYS.map(k => (!completo && colsVazias.has(k)) ? '' : (item[k] || ''))
    ),
  ]
  const ws = XLSX.utils.aoa_to_sheet(wsData)
  XLSX.utils.book_append_sheet(wb, ws, 'Planilha1')
  return XLSX.write(wb, { type: 'buffer', bookType: 'xls' })
}

function gerarXml(pedido) {
  const now = new Date().toISOString().slice(0, 19)
  const itensXml = pedido.itens.map((item, i) => `
    <Item seq="${i + 1}">
      <CodProduto>${esc(item.codproduto)}</CodProduto>
      <CodEmbalagem>${esc(item.codembalagem)}</CodEmbalagem>
      <Quantidade>${esc(item.quantidade)}</Quantidade>
      <Descricao>${esc(item.descricao)}</Descricao>
      <Emba>${esc(item.emba)}</Emba>
      <QtUnit>${esc(item.qtUnit)}</QtUnit>
      <PrecoVenda>${esc(item.precoVenda)}</PrecoVenda>
      <PrecoEmba>${esc(item.preco_emba)}</PrecoEmba>
      <PrecoEmbaST>${esc(item.preco_emba_st)}</PrecoEmbaST>
      <PrecoUnit>${esc(item.preco_unit)}</PrecoUnit>
      <PrecoTot>${esc(item.preco_tot)}</PrecoTot>
      <PrecoTotION>${esc(item.preco_tot_ion)}</PrecoTotION>
      <PrecoTotIONST>${esc(item.preco_tot_ion_st)}</PrecoTotIONST>
    </Item>`).join('')

  return `<?xml version="1.0" encoding="UTF-8"?>
<ImportacaoProdutos geradoEm="${now}" versao="1.0">
  <Cabecalho>
    <CNPJ>${esc(pedido.cnpj_formatado)}</CNPJ>
    <CNPJNumerico>${esc(pedido.cnpj_numerico)}</CNPJNumerico>
    <RazaoSocial>${esc(pedido.razao_social)}</RazaoSocial>
    <Loja>${esc(pedido.loja)}</Loja>
    <NumeroPedido>${esc(pedido.num_pedido)}</NumeroPedido>
    <DataCompra>${esc(pedido.data_compra)}</DataCompra>
    <DataEntrega>${esc(pedido.data_entrega)}</DataEntrega>
    <Vencimentos>${esc(pedido.vencimentos)}</Vencimentos>
    <TotalItens>${pedido.itens.length}</TotalItens>
  </Cabecalho>
  <Itens>${itensXml}
  </Itens>
</ImportacaoProdutos>`
}

function gerarCsv(pedidos) {
  const q = v => `"${String(v || '').replace(/"/g, '""')}"`
  const header = [
    'loja', 'num_pedido', 'cnpj', 'razao_social', 'data_compra', 'data_entrega',
    'vencimentos', 'codproduto', 'descricao', 'codembalagem', 'emba',
    'quantidade', 'precoVenda', 'frete', 'desconto', 'preco_tot',
  ]
  const rows = [header.map(q).join(';')]
  for (const p of pedidos) {
    for (const item of p.itens) {
      rows.push([
        p.loja, p.num_pedido, p.cnpj_formatado, p.razao_social,
        p.data_compra, p.data_entrega, p.vencimentos,
        item.codproduto, item.descricao, item.codembalagem,
        item.emba, item.quantidade, item.precoVenda || '',
        item.frete || '', item.desconto || '', item.preco_tot || '',
      ].map(q).join(';'))
    }
  }
  return '﻿' + rows.join('\r\n')
}

// ── Helpers para detecção de estrutura genérica ───────────────────────────────

function padArr(arr, n) {
  const r = arr.slice(0, n)
  while (r.length < n) r.push('')
  return r
}

// Remove prefixo monetário "R$" de valores — mantém apenas o número
function limparCelula(v) {
  return String(v ?? '').replace(/R\$\s*/gi, '').trim()
}

// Detecta se uma linha é cabeçalho: campos textuais, sem valores puramente numéricos
function ehCabecalho(campos) {
  return campos.length > 1 &&
    campos.every(c => !/^[\d.,]+$/.test(c.trim())) &&
    campos.some(c => c.trim().length > 1)
}

// ── Estratégia A: divisão por múltiplos espaços ───────────────────────────────
function tentarMultiEspaco(linhasDados) {
  const splitadas = linhasDados.map(l =>
    l.split(/\s{2,}/).map(c => limparCelula(c)).filter(Boolean)
  )
  const freq = {}
  splitadas.forEach(r => { if (r.length >= 2) freq[r.length] = (freq[r.length] || 0) + 1 })
  if (!Object.keys(freq).length) return null

  const nCols     = parseInt(Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0])
  const cobertura = (freq[nCols] || 0) / linhasDados.length
  if (nCols < 2 || cobertura < 0.25) return null

  const comCols   = splitadas.filter(r => r.length >= nCols - 1)
  const iCab      = comCols.findIndex(ehCabecalho)
  const cabecalho = iCab >= 0
    ? padArr(comCols[iCab], nCols)
    : Array.from({ length: nCols }, (_, i) => `Campo ${i + 1}`)
  const dados     = comCols
    .filter((_, i) => i !== iCab)
    .map(r => padArr(r.map(limparCelula), nCols))

  return { cabecalho, dados }
}

// ── Estratégia B: detecção de colunas por posição de caracteres ───────────────
// Útil para relatórios de largura fixa onde colunas são alinhadas por espaços.
function tentarPosicaoFixa(linhasDados) {
  const comprMedio  = linhasDados.reduce((s, l) => s + l.length, 0) / (linhasDados.length || 1)
  const linhasAlvo  = linhasDados.filter(l => l.length >= comprMedio * 0.5)
  if (linhasAlvo.length < 3) return null

  const maxLen = Math.max(...linhasAlvo.map(l => l.length))
  const freq   = new Float32Array(maxLen)
  for (const l of linhasAlvo)
    for (let i = 0; i < l.length; i++)
      if (l[i] === ' ') freq[i]++
  for (let i = 0; i < maxLen; i++) freq[i] /= linhasAlvo.length

  // Bordas de colunas: fim de uma região com ≥70% de espaços e largura ≥2
  const bordas = []
  let inGap = false, gapStart = 0
  for (let i = 1; i < maxLen; i++) {
    if (freq[i] >= 0.7 && !inGap) { inGap = true; gapStart = i }
    else if (freq[i] < 0.7 && inGap) { if (i - gapStart >= 2) bordas.push(i); inGap = false }
  }
  if (bordas.length < 2) return null

  const cortar = linha => {
    const cols = []
    let ini = 0
    for (const b of bordas) { cols.push(limparCelula(linha.slice(ini, b))); ini = b }
    cols.push(limparCelula(linha.slice(ini)))
    return cols.filter(Boolean)
  }

  const nCols     = bordas.length + 1
  const splitadas = linhasDados.map(cortar)
  const iCab      = splitadas.findIndex(ehCabecalho)
  const cabecalho = iCab >= 0
    ? padArr(splitadas[iCab], nCols)
    : Array.from({ length: nCols }, (_, i) => `Campo ${i + 1}`)
  const dados     = splitadas
    .filter((_, i) => i !== iCab)
    .map(r => padArr(r, nCols))

  return { cabecalho, dados }
}

// ── Estratégia C: pares chave–valor (ex: "Nome: João", "Data: 01/01/2025") ───
function tentarChaveValor(linhasDados) {
  const RE_KV = /^([^:\n]{1,60}):\s*(.+)$/
  const pares = linhasDados.map(l => RE_KV.exec(l)).filter(Boolean)
  if (pares.length / linhasDados.length < 0.4) return null
  return {
    cabecalho: ['Campo', 'Valor'],
    dados: pares.map(m => [m[1].trim(), limparCelula(m[2].trim())]),
  }
}

// ── Estratégia baseada em posição X/Y real (melhor resultado para tabelas) ────
// paginas: array de páginas; cada página é array de {str, x, y}
function extrairTabelaPorPosicao(paginas) {
  const todasLinhas = []

  for (const itensBrutos of paginas) {
    // Deduplicar itens com mesma posição e texto (camadas sobrepostas no PDF)
    const vistos = new Set()
    const itens = itensBrutos.filter(it => {
      const k = `${it.x}|${it.y}|${it.str}`
      if (vistos.has(k)) return false
      vistos.add(k)
      return true
    })

    // Agrupar em linhas por Y (tolerância 4px)
    const linhas = []
    for (const it of itens) {
      const l = linhas.find(l => Math.abs(l.y - it.y) <= 4)
      if (l) l.itens.push(it)
      else linhas.push({ y: it.y, itens: [it] })
    }

    // De cima para baixo (Y decrescente em coordenadas PDF)
    linhas.sort((a, b) => b.y - a.y)
    for (const l of linhas) {
      l.itens.sort((a, b) => a.x - b.x)
      todasLinhas.push(l.itens)
    }
  }

  if (!todasLinhas.length) return null

  // Clusterizar posições X (tolerância 20px) para detectar colunas
  const clusters = []
  for (const x of todasLinhas.flatMap(l => l.map(it => it.x)).sort((a, b) => a - b)) {
    const c = clusters.find(cl => Math.abs(cl.ref - x) <= 20)
    if (c) { c.xs.push(x); c.ref = Math.round(c.xs.reduce((s, v) => s + v, 0) / c.xs.length) }
    else clusters.push({ ref: x, xs: [x] })
  }
  clusters.sort((a, b) => a.ref - b.ref)
  if (clusters.length < 2) return null

  const nCols = clusters.length

  // Construir grade: cada linha → array de nCols células
  const grade = todasLinhas.map(itensLinha => {
    const linha = new Array(nCols).fill('')
    for (const it of itensLinha) {
      const ci = clusters.findIndex(cl => Math.abs(cl.ref - it.x) <= 20)
      if (ci >= 0) linha[ci] = linha[ci] ? `${linha[ci]} ${it.str}` : it.str
    }
    return linha
  }).filter(r => r.some(c => c.trim()))

  if (grade.length < 2) return null

  // Detectar cabeçalho na primeira linha não-vazia com texto (não números)
  let cabecalho, dados
  if (ehCabecalho(grade[0].filter(Boolean))) {
    cabecalho = grade[0].map((c, i) => c.trim() || `Coluna ${i + 1}`)
    dados = grade.slice(1)
  } else {
    cabecalho = clusters.map((_, i) => `Coluna ${i + 1}`)
    dados = grade
  }

  return { cabecalho, dados }
}

// ── Orquestrador de detecção ──────────────────────────────────────────────────
function detectarEstrutura(linhas) {
  const comprMedio = linhas.reduce((s, l) => s + l.length, 0) / (linhas.length || 1)

  // Primeiras linhas curtas (< 55% da média) → metadados do relatório
  let corte = 0
  for (let i = 0; i < Math.min(8, linhas.length); i++) {
    if (linhas[i].length < comprMedio * 0.55) corte = i + 1
    else break
  }
  const metadados   = linhas.slice(0, Math.max(1, corte))
  const linhasDados = linhas.slice(corte)

  const resultado = tentarMultiEspaco(linhasDados) ?? tentarPosicaoFixa(linhasDados) ?? tentarChaveValor(linhasDados)
  if (resultado) return { metadados, ...resultado }

  // Fallback: coluna única preservando todo o conteúdo
  return {
    metadados,
    cabecalho: ['Conteúdo do Documento'],
    dados: linhasDados.map(l => [limparCelula(l)]),
  }
}

// ── XML genérico ──────────────────────────────────────────────────────────────
function gerarXmlGenerico(metadados, cabecalho, dados, cnpj) {
  const now      = new Date().toISOString().slice(0, 19)
  const itensXml = dados.map((linha, i) => {
    const cols = cabecalho.map((h, ci) =>
      `      <${sanitizarTag(h)}>${esc(String(linha[ci] ?? ''))}</${sanitizarTag(h)}>`
    ).join('\n')
    return `    <Item seq="${i + 1}">\n${cols}\n    </Item>`
  }).join('\n')

  const cabXml = [
    cnpj ? `    <CNPJ>${esc(cnpj.formatado)}</CNPJ>` : '',
    cnpj ? `    <CNPJNumerico>${cnpj.numerico}</CNPJNumerico>` : '',
    `    <Titulo>${esc(metadados[0] || '')}</Titulo>`,
    `    <TotalItens>${dados.length}</TotalItens>`,
  ].filter(Boolean).join('\n')

  return `<?xml version="1.0" encoding="UTF-8"?>
<DocumentoExtraido geradoEm="${now}" versao="1.0">
  <Cabecalho>
${cabXml}
  </Cabecalho>
  <Itens>
${itensXml}
  </Itens>
</DocumentoExtraido>`
}

// ── Saída genérica: qualquer PDF não reconhecido ──────────────────────────────
// Cascata de estratégias:
//   1. Posição X/Y real (preserva colunas de tabelas visuais)
//   2. Multi-espaço | posição fixa | chave-valor | coluna única (fallback texto)
async function gerarSaidaGenerica(text, paginas, nomeArquivo) {
  const nome  = sanitizarNome(nomeArquivo)
  const cnpj  = extrairPrimeiroCnpj(text)

  const todasLinhasTexto = text.split('\n').map(l => l.trim()).filter(Boolean)
  const metadados = todasLinhasTexto.slice(0, 3)

  // Estratégia 1: extração por posição X/Y (melhor para tabelas)
  let estrutura = extrairTabelaPorPosicao(paginas)

  // Estratégias 2-5: baseadas em texto puro (fallback)
  if (!estrutura) {
    const linhas = filtrarRepetitivas(todasLinhasTexto)
    if (!linhas.length) throw Object.assign(new Error('Nenhum conteúdo legível encontrado no PDF.'), { status: 422 })
    estrutura = detectarEstrutura(linhas)
  }

  const { cabecalho, dados } = estrutura
  const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`

  // ── XLS ───────────────────────────────────────────────────────────────────
  const infoLinha = cnpj
    ? ['cnpj', cnpj.numerico]
    : ['documento', metadados.join('  |  ')]
  const wb     = XLSX.utils.book_new()
  const wsData = [infoLinha, cabecalho, ...dados]
  const ws     = XLSX.utils.aoa_to_sheet(wsData)

  ws['!cols'] = cabecalho.map((h, ci) => ({
    wch: Math.min(60, Math.max(
      String(h ?? '').length + 2,
      ...dados.map(r => String(r[ci] ?? '').length)
    ))
  }))

  XLSX.utils.book_append_sheet(wb, ws, 'Dados')
  const xlsBuf = XLSX.write(wb, { type: 'buffer', bookType: 'xls' })
  const xlsB64 = Buffer.from(xlsBuf).toString('base64')

  // ── XML ───────────────────────────────────────────────────────────────────
  const xmlStr  = gerarXmlGenerico(metadados, cabecalho, dados, cnpj)
  const xmlNome = `${nome}.xml`

  // ── CSV ───────────────────────────────────────────────────────────────────
  const csvContent = '﻿' + [
    cabecalho.map(q).join(';'),
    ...dados.map(r => r.map(q).join(';')),
  ].join('\r\n')
  const csvB64  = Buffer.from(csvContent, 'utf-8').toString('base64')
  const csvNome = `${nome}.csv`

  // ── ZIP ───────────────────────────────────────────────────────────────────
  const zip = new JSZip()
  zip.file(`${nome}.xls`, xlsBuf)
  zip.file(xmlNome, xmlStr)
  zip.file(csvNome, Buffer.from(csvContent, 'utf-8'))
  const zipB64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' })

  const razao  = (metadados[0] || nome).substring(0, 80)
  const lojaId = cnpj
    ? cnpj.numerico.slice(-6)
    : (metadados[1] || metadados[0] || nome).substring(0, 20).trim() || 'GERAL'

  return {
    pedidos: [{
      loja:              lojaId,
      razao_social:      razao,
      num_pedido:        '',
      cnpj:              cnpj?.formatado || '',
      total_itens:       dados.length,
      data_compra:       '',
      data_entrega:      '',
      xls_nome:          `${nome}.xls`,
      xls_completo_nome: `${nome}.xls`,
      xml_nome:          xmlNome,
      xls:               xlsB64,
      xls_completo:      xlsB64,
      xml:               Buffer.from(xmlStr, 'utf-8').toString('base64'),
    }],
    zip:      zipB64,
    csv:      csvB64,
    csv_nome: csvNome,
  }
}

// ── Ponto de entrada ──────────────────────────────────────────────────────────
export async function runConverter(body) {
  const b64 = body?.pdf
  if (!b64) throw Object.assign(new Error('PDF não fornecido'), { status: 400 })

  const buffer  = Buffer.from(b64, 'base64')
  const paginas = []

  const { text } = await pdfParse(buffer, {
    pagerender: async pageData => {
      const content = await pageData.getTextContent()
      const itens = []
      for (const it of content.items) {
        if (it.str && it.str.trim()) {
          itens.push({
            str: it.str.trim(),
            x: Math.round(it.transform[4]),
            y: Math.round(it.transform[5]),
          })
        }
      }
      paginas.push(itens)
      return ''
    },
  })

  const formato  = detectarFormato(text)
  const nome     = sanitizarNome(body.nome)

  // Formato não reconhecido → exporta todo o conteúdo como planilha genérica
  if (formato === 'generico') return gerarSaidaGenerica(text, paginas, nome)

  const pedidos = formato === 'ruptura'
    ? extrairRupturaPrePedido(text)
    : extrairPedidos(text)

  if (!pedidos.length) throw Object.assign(new Error('Nenhum pedido encontrado no PDF'), { status: 422 })

  const colsVazias = COLS_VAZIAS_PEDIDO

  const zip = new JSZip()
  const pedidosComArquivos = pedidos.map(p => {
    const slug       = p.loja.replace(/[^\w]/g, '_')
    const id         = p.num_pedido || p.cnpj_numerico || slug
    const xlsBuf     = _montarXls(p, false, colsVazias)
    const xlsCompBuf = _montarXls(p, true,  colsVazias)
    const xmlStr     = gerarXml(p)

    zip.file(`pedido_${id}_${slug}.xls`,          xlsBuf)
    zip.file(`pedido_${id}_${slug}_completo.xls`, xlsCompBuf)
    zip.file(`pedido_${id}_${slug}.xml`,          xmlStr)

    return {
      loja:              p.loja,
      razao_social:      p.razao_social,
      num_pedido:        p.num_pedido,
      cnpj:              p.cnpj_formatado,
      total_itens:       p.itens.length,
      data_compra:       p.data_compra,
      data_entrega:      p.data_entrega,
      xls_nome:          `pedido_${id}_${slug}.xls`,
      xls_completo_nome: `pedido_${id}_${slug}_completo.xls`,
      xml_nome:          `pedido_${id}_${slug}.xml`,
      xls:               Buffer.from(xlsBuf).toString('base64'),
      xls_completo:      Buffer.from(xlsCompBuf).toString('base64'),
      xml:               Buffer.from(xmlStr, 'utf-8').toString('base64'),
    }
  })

  const csvStr = gerarCsv(pedidos)
  zip.file('pedidos_combinados.csv', csvStr)

  const zipB64 = await zip.generateAsync({ type: 'base64' })
  const csvB64 = Buffer.from(csvStr, 'utf-8').toString('base64')

  return {
    pedidos:  pedidosComArquivos,
    zip:      zipB64,
    csv:      csvB64,
    csv_nome: 'pedidos_combinados.csv',
  }
}
