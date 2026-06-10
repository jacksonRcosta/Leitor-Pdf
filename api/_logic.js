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

// Colunas deixadas em branco na versão XLS padrão, por tipo de documento
const COLS_VAZIAS_PEDIDO  = new Set(['codproduto', 'descricao', 'emba', 'preco_unit', 'preco_tot'])
const COLS_VAZIAS_RUPTURA = new Set(['emba', 'qtUnit', 'precoVenda', 'preco_emba', 'preco_emba_st', 'preco_unit', 'preco_tot', 'preco_tot_ion', 'preco_tot_ion_st'])

// ── Utilitários ───────────────────────────────────────────────────────────────
const esc     = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')
const cnpjNum = s => s.replace(/\D/g, '')
const fmtCNPJ = c => /^\d{14}$/.test(c)
  ? c.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5')
  : c

// ── Detecção de formato ───────────────────────────────────────────────────────
function detectarFormato(text) {
  if (/ruptura\s*[-–]?\s*pr[eé]/i.test(text)) return 'ruptura'
  if (/RAZ[ÃA]O SOCIAL:CNPJ:/i.test(text))    return 'pedido'
  return 'pedido'
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

      const mItem = RE_ITEM.exec(linha)
      if (mItem) {
        atual.itens.push({
          codproduto:       mItem[1],
          descricao:        mItem[2].trim(),
          emba:             mItem[3],
          externo:          mItem[4],
          precoVenda:       mItem[5],
          frete:            mItem[6],
          preco_tot:        mItem[7],
          quantidade:       mItem[8],
          codembalagem:     mItem[9],
          desconto:         mItem[10],
          qtUnit:           '',
          preco_emba:       '',
          preco_emba_st:    '',
          preco_unit:       mItem[5],
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
// Âncoras fixas: CNPJ(14d) sem espaço após razão social; EAN(13d) sem espaço
// após descrição; datas coladas à quantidade e entre si no final da linha.
//
// Agrupamento: um "pedido" por CNPJ único — múltiplas linhas com o mesmo CNPJ
// são consolidadas em um único conjunto de arquivos (XLS/XML).
const RE_RUPTURA = /^(\d{6}) - (\d{2}) \| (.+?)(\d{14})(\d{6}) - (.+?)(\d{13})(\d{6}) - (.+?)(\d+)(\d{2}\/\d{2}\/\d{4})(\d{2}\/\d{2}\/\d{4})$/

function extrairRupturaPrePedido(texto) {
  const mapa = new Map() // CNPJ → pedido

  for (const linha of texto.split('\n')) {
    const m = RE_RUPTURA.exec(linha.trim())
    if (!m) continue

    const [, codloja, filial, razaoSocial, cnpj, codprod, descProd, ean, qtd, dataRuptura] = m

    if (!mapa.has(cnpj)) {
      mapa.set(cnpj, {
        loja:           `${codloja}-${filial}`,
        razao_social:   razaoSocial.trim(),
        cnpj_formatado: fmtCNPJ(cnpj),
        cnpj_numerico:  cnpj,
        num_pedido:     '',
        data_compra:    dataRuptura,
        data_entrega:   '',
        vencimentos:    '',
        itens:          [],
      })
    }

    mapa.get(cnpj).itens.push({
      codproduto:       codprod,
      descricao:        descProd.trim().replace(/\s{2,}/g, ' '),
      codembalagem:     ean,
      emba:             '',
      quantidade:       qtd,
      qtUnit:           '',
      precoVenda:       '',
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

// ── Ponto de entrada ──────────────────────────────────────────────────────────
export async function runConverter(body) {
  const b64 = body?.pdf
  if (!b64) throw Object.assign(new Error('PDF não fornecido'), { status: 400 })

  const buffer   = Buffer.from(b64, 'base64')
  const { text } = await pdfParse(buffer)
  const formato  = detectarFormato(text)

  const pedidos = formato === 'ruptura'
    ? extrairRupturaPrePedido(text)
    : extrairPedidos(text)

  if (!pedidos.length) throw Object.assign(new Error('Nenhum pedido encontrado no PDF'), { status: 422 })

  const colsVazias = formato === 'ruptura' ? COLS_VAZIAS_RUPTURA : COLS_VAZIAS_PEDIDO

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
