const pdfParse = require('pdf-parse');
const XLSX     = require('xlsx');
const JSZip    = require('jszip');

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

// ── Regex (formato pdf-parse — campos sem espaço) ─────────────────────────────
// Header: "BARROS COMERCIO LTDA 0009RAZÃO SOCIAL:CNPJ:10.840.716/0009-08I.E.:240332490"
const RE_RAZAO   = /^(.+?)RAZ[ÃA]O SOCIAL:CNPJ:(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})I\.E\.:(\d+)/;
// Loja em linha separada: "LOJA ENTREGA:LJ09 - SUP.SAO"
const RE_LOJA    = /LOJA ENTREGA:(\w+)/;
// Pedido na linha do fornecedor: "...N° PEDIDO:172409"
const RE_PEDIDO  = /PEDIDO[:\s]*(\d+)/i;
const RE_COMPRA  = /DATA COMPRA:(\d{2}\/\d{2}\/\d{4})/;
const RE_ENTREGA = /DATA ENTREGA:(\d{2}\/\d{2}\/\d{4})/;
const RE_VENC    = /VENCIMENTOS:(\S+)/;
// Item: cod(6) + desc + emb(2L/4D) + externo(10-12D) + custo + frete + total + qtde + EAN(13D) + disc
const RE_ITEM    = /^(\d{6})(.+?)([A-Z]{2}\/\d{4})(\d{10,12})([\d.]+,\d{2})([\d.]+,\d{2})([\d.]+,\d{2})(\d+)(\d{13,14})([\d.]+,\d{2})/;

const XLS_HEADERS = [
  'codproduto','codembalagem','quantidade','descricao',
  'emba','qtUnit','precoVenda','preço emba','preço emba st',
  'preço unit','preço tot','preco tot ion','preco tot ion st',
];
const XLS_KEYS = [
  'codproduto','codembalagem','quantidade','descricao',
  'emba','qtUnit','precoVenda','preco_emba','preco_emba_st',
  'preco_unit','preco_tot','preco_tot_ion','preco_tot_ion_st',
];

const esc = s => String(s || '')
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

const cnpjNum = s => s.replace(/\D/g, '');

// ── Extração ──────────────────────────────────────────────────────────────────
function extrairPedidos(texto) {
  const pedidos = [];
  let atual = null;

  for (const linha of texto.split('\n')) {
    // Detecta novo pedido pelo cabeçalho RAZÃO SOCIAL
    const mRazao = RE_RAZAO.exec(linha);
    if (mRazao) {
      if (atual && atual.itens.length) pedidos.push(atual);
      const fmt = mRazao[2].trim();
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
      };
    }

    if (atual) {
      if (!atual.loja)        { const m = RE_LOJA.exec(linha);    if (m) atual.loja         = m[1]; }
      if (!atual.num_pedido)  { const m = RE_PEDIDO.exec(linha);  if (m) atual.num_pedido   = m[1]; }
      if (!atual.data_compra) { const m = RE_COMPRA.exec(linha);  if (m) atual.data_compra  = m[1]; }
      if (!atual.data_entrega){ const m = RE_ENTREGA.exec(linha); if (m) atual.data_entrega = m[1]; }
      if (!atual.vencimentos) { const m = RE_VENC.exec(linha);    if (m) atual.vencimentos  = m[1]; }

      // Item: cod(6) desc emb externo custo frete total qtde EAN disc
      const mItem = RE_ITEM.exec(linha);
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
        });
      }
    }
  }

  if (atual && atual.itens.length) pedidos.push(atual);
  return pedidos;
}

// ── Geradores ─────────────────────────────────────────────────────────────────
function gerarXls(pedido) {
  const wb = XLSX.utils.book_new();
  const wsData = [
    ['cnpj', pedido.cnpj_numerico],
    XLS_HEADERS,
    ...pedido.itens.map(item => XLS_KEYS.map(k => item[k] || '')),
  ];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, 'Planilha1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xls' });
}

function gerarXml(pedido) {
  const now = new Date().toISOString().slice(0, 19);
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
    </Item>`).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>\n<ImportacaoProdutos geradoEm="${now}" versao="1.0">\n  <Cabecalho>\n    <CNPJ>${esc(pedido.cnpj_formatado)}</CNPJ>\n    <CNPJNumerico>${esc(pedido.cnpj_numerico)}</CNPJNumerico>\n    <RazaoSocial>${esc(pedido.razao_social)}</RazaoSocial>\n    <Loja>${esc(pedido.loja)}</Loja>\n    <NumeroPedido>${esc(pedido.num_pedido)}</NumeroPedido>\n    <DataCompra>${esc(pedido.data_compra)}</DataCompra>\n    <DataEntrega>${esc(pedido.data_entrega)}</DataEntrega>\n    <Vencimentos>${esc(pedido.vencimentos)}</Vencimentos>\n    <TotalItens>${pedido.itens.length}</TotalItens>\n  </Cabecalho>\n  <Itens>${itensXml}\n  </Itens>\n</ImportacaoProdutos>`;
}

function gerarCsv(pedidos) {
  const q = v => `"${String(v || '').replace(/"/g, '""')}"`;
  const header = ['loja','num_pedido','cnpj','razao_social','data_compra','data_entrega',
    'vencimentos','codproduto','descricao','codembalagem','emba',
    'quantidade','precoVenda','frete','desconto','preco_tot'];
  const rows = [header.map(q).join(';')];
  for (const p of pedidos) {
    for (const item of p.itens) {
      rows.push([
        p.loja, p.num_pedido, p.cnpj_formatado, p.razao_social,
        p.data_compra, p.data_entrega, p.vencimentos,
        item.codproduto, item.descricao, item.codembalagem,
        item.emba, item.quantidade, item.precoVenda,
        item.frete, item.desconto, item.preco_tot,
      ].map(q).join(';'));
    }
  }
  return '﻿' + rows.join('\r\n');
}

// ── Handler ───────────────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: CORS, body: '' };
  }

  try {
    const body = JSON.parse(event.body || '{}');
    if (!body.pdf) {
      return { statusCode: 400, headers: CORS, body: JSON.stringify({ erro: 'PDF não fornecido' }) };
    }

    const pdfBuffer = Buffer.from(body.pdf, 'base64');
    const parsed    = await pdfParse(pdfBuffer);
    const pedidos   = extrairPedidos(parsed.text);

    if (!pedidos.length) {
      return { statusCode: 422, headers: CORS, body: JSON.stringify({ erro: 'Nenhum pedido encontrado no PDF' }) };
    }

    const zip = new JSZip();
    const pedidosComArquivos = pedidos.map(p => {
      const slug   = p.loja.replace(/[^\w]/g, '_');
      const id     = p.num_pedido || slug;
      const xlsBuf = gerarXls(p);
      const xmlStr = gerarXml(p);
      zip.file(`pedido_${id}_${slug}.xls`, xlsBuf);
      zip.file(`pedido_${id}_${slug}.xml`, xmlStr);
      return {
        loja:          p.loja,
        razao_social:  p.razao_social,
        num_pedido:    p.num_pedido,
        cnpj:          p.cnpj_formatado,
        total_itens:   p.itens.length,
        data_compra:   p.data_compra,
        data_entrega:  p.data_entrega,
        xls_nome:      `pedido_${id}_${slug}.xls`,
        xml_nome:      `pedido_${id}_${slug}.xml`,
        xls:           Buffer.from(xlsBuf).toString('base64'),
        xml:           Buffer.from(xmlStr, 'utf-8').toString('base64'),
      };
    });

    const csvStr  = gerarCsv(pedidos);
    zip.file('pedidos_combinados.csv', csvStr);

    const zipB64 = await zip.generateAsync({ type: 'base64' });
    const csvB64 = Buffer.from(csvStr, 'utf-8').toString('base64');

    return {
      statusCode: 200,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify({ zip: zipB64, csv: csvB64, csv_nome: 'pedidos_combinados.csv', pedidos: pedidosComArquivos }),
    };

  } catch (err) {
    return {
      statusCode: 500,
      headers: CORS,
      body: JSON.stringify({ erro: err.message }),
    };
  }
};
