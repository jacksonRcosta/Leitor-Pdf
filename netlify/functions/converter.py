import json
import base64
import io
import re
import csv
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime


# ── Regex ─────────────────────────────────────────────────────────────────────
RE_RAZAO = re.compile(
    r'RAZ[AÃ]O SOCIAL:\s*(.+?)\s+CNPJ:\s*([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})'
    r'\s+I\.E\.:\s*([\d]+)\s+LOJA ENTREGA:\s*(\S+)',
    re.IGNORECASE,
)
RE_PEDIDO       = re.compile(r'N[ºO]\s*PEDIDO\s*[:\s]*(\d+)', re.IGNORECASE)
RE_DATA_COMPRA  = re.compile(r'DATA COMPRA:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
RE_DATA_ENTREGA = re.compile(r'DATA ENTREGA:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
RE_VENCIMENTOS  = re.compile(r'VENCIMENTOS:\s*(\S+)', re.IGNORECASE)
RE_ITEM = re.compile(
    r'^(\d{6})\s+(.+?)\s+(\d{13,14})\s+(\w+/\d{4})\s+(\d{10,12})\s+(\d+)\s+'
    r'([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})',
    re.MULTILINE,
)

CABECALHOS_XLS = [
    'codproduto', 'codembalagem', 'quantidade', 'descricao',
    'emba', 'qtUnit', 'precoVenda', 'preço emba', 'preço emba st',
    'preço unit', 'preço tot', 'preco tot ion', 'preco tot ion st',
]
MAPA_XLS = [
    'codproduto', 'codembalagem', 'quantidade', 'descricao',
    'emba', 'qtUnit', 'precoVenda', 'preco_emba', 'preco_emba_st',
    'preco_unit', 'preco_tot', 'preco_tot_ion', 'preco_tot_ion_st',
]
LARGURAS_XLS = [15, 18, 12, 40, 8, 10, 14, 14, 16, 12, 12, 14, 16]


def _cnpj_num(s):
    return re.sub(r'\D', '', s)


def extrair_pedidos(pdf_bytes):
    import pdfplumber

    pedidos = []
    atual = None

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pag in pdf.pages:
            texto = pag.extract_text() or ''
            m_razao = RE_RAZAO.search(texto)

            if m_razao:
                if atual and atual['itens']:
                    pedidos.append(atual)
                cnpj_fmt = m_razao.group(2).strip()
                num  = RE_PEDIDO.search(texto)
                dc   = RE_DATA_COMPRA.search(texto)
                de   = RE_DATA_ENTREGA.search(texto)
                venc = RE_VENCIMENTOS.search(texto)
                atual = {
                    'razao_social':   m_razao.group(1).strip(),
                    'cnpj_formatado': cnpj_fmt,
                    'cnpj_numerico':  _cnpj_num(cnpj_fmt),
                    'ie':             m_razao.group(3).strip(),
                    'loja':           m_razao.group(4).strip(),
                    'num_pedido':     num.group(1)  if num  else '',
                    'data_compra':    dc.group(1)   if dc   else '',
                    'data_entrega':   de.group(1)   if de   else '',
                    'vencimentos':    venc.group(1) if venc else '',
                    'itens':          [],
                }

            if atual is not None:
                for m in RE_ITEM.finditer(texto):
                    atual['itens'].append({
                        'codproduto':      m.group(1),
                        'descricao':       m.group(2).strip(),
                        'codembalagem':    m.group(3),
                        'emba':            m.group(4),
                        'externo':         m.group(5),
                        'quantidade':      m.group(6),
                        'precoVenda':      m.group(7),
                        'frete':           m.group(8),
                        'desconto':        m.group(9),
                        'preco_tot':       m.group(10),
                        'qtUnit':          '',
                        'preco_emba':      '',
                        'preco_emba_st':   '',
                        'preco_unit':      m.group(7),
                        'preco_tot_ion':   '',
                        'preco_tot_ion_st': '',
                    })

    if atual and atual['itens']:
        pedidos.append(atual)

    return pedidos


def _gerar_xls_bytes(pedido):
    import xlwt
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Planilha1', cell_overwrite_ok=True)
    ws.write(0, 0, 'cnpj')
    ws.write(0, 1, pedido['cnpj_numerico'])
    for col, cab in enumerate(CABECALHOS_XLS):
        ws.write(1, col, cab)
    for row, item in enumerate(pedido['itens'], start=2):
        for col, chave in enumerate(MAPA_XLS):
            ws.write(row, col, item.get(chave, ''))
    for col, larg in enumerate(LARGURAS_XLS):
        ws.col(col).width = larg * 256
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _gerar_xml_bytes(pedido):
    raiz = ET.Element('ImportacaoProdutos')
    raiz.set('geradoEm', datetime.now().strftime('%Y-%m-%dT%H:%M:%S'))
    raiz.set('versao', '1.0')
    cab = ET.SubElement(raiz, 'Cabecalho')
    ET.SubElement(cab, 'CNPJ').text         = pedido['cnpj_formatado']
    ET.SubElement(cab, 'CNPJNumerico').text = pedido['cnpj_numerico']
    ET.SubElement(cab, 'RazaoSocial').text  = pedido['razao_social']
    ET.SubElement(cab, 'Loja').text         = pedido['loja']
    ET.SubElement(cab, 'NumeroPedido').text = pedido['num_pedido']
    ET.SubElement(cab, 'DataCompra').text   = pedido['data_compra']
    ET.SubElement(cab, 'DataEntrega').text  = pedido['data_entrega']
    ET.SubElement(cab, 'Vencimentos').text  = pedido['vencimentos']
    ET.SubElement(cab, 'TotalItens').text   = str(len(pedido['itens']))
    itens_el = ET.SubElement(raiz, 'Itens')
    for seq, item in enumerate(pedido['itens'], start=1):
        el = ET.SubElement(itens_el, 'Item')
        el.set('seq', str(seq))
        ET.SubElement(el, 'CodProduto').text    = item['codproduto']
        ET.SubElement(el, 'CodEmbalagem').text  = item['codembalagem']
        ET.SubElement(el, 'Quantidade').text    = item['quantidade']
        ET.SubElement(el, 'Descricao').text     = item['descricao']
        ET.SubElement(el, 'Emba').text          = item['emba']
        ET.SubElement(el, 'QtUnit').text        = item['qtUnit']
        ET.SubElement(el, 'PrecoVenda').text    = item['precoVenda']
        ET.SubElement(el, 'PrecoEmba').text     = item['preco_emba']
        ET.SubElement(el, 'PrecoEmbaST').text   = item['preco_emba_st']
        ET.SubElement(el, 'PrecoUnit').text     = item['preco_unit']
        ET.SubElement(el, 'PrecoTot').text      = item['preco_tot']
        ET.SubElement(el, 'PrecoTotION').text   = item['preco_tot_ion']
        ET.SubElement(el, 'PrecoTotIONST').text = item['preco_tot_ion_st']
    xml_str = minidom.parseString(
        ET.tostring(raiz, encoding='unicode')
    ).toprettyxml(indent='  ')
    return xml_str.encode('utf-8')


def _gerar_csv_bytes(pedidos):
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=';', quoting=csv.QUOTE_ALL)
    w.writerow([
        'loja', 'num_pedido', 'cnpj', 'razao_social',
        'data_compra', 'data_entrega', 'vencimentos',
        'codproduto', 'descricao', 'codembalagem', 'emba',
        'quantidade', 'precoVenda', 'frete', 'desconto', 'preco_tot',
    ])
    for p in pedidos:
        for item in p['itens']:
            w.writerow([
                p['loja'], p['num_pedido'], p['cnpj_formatado'], p['razao_social'],
                p['data_compra'], p['data_entrega'], p['vencimentos'],
                item['codproduto'], item['descricao'], item['codembalagem'],
                item['emba'], item['quantidade'], item['precoVenda'],
                item['frete'], item['desconto'], item['preco_tot'],
            ])
    return ('﻿' + buf.getvalue()).encode('utf-8')


def _montar_zip(pedidos):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for p in pedidos:
            slug = re.sub(r'[^\w]', '_', p['loja'])
            pedido_id = p['num_pedido'] or slug
            zf.writestr(f'pedido_{pedido_id}_{slug}.xls', _gerar_xls_bytes(p))
            zf.writestr(f'pedido_{pedido_id}_{slug}.xml', _gerar_xml_bytes(p))
        zf.writestr('pedidos_combinados.csv', _gerar_csv_bytes(pedidos))
    return buf.getvalue()


def handler(event, context):
    headers_cors = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
    }

    if event.get('httpMethod') == 'OPTIONS':
        return {'statusCode': 200, 'headers': headers_cors, 'body': ''}

    try:
        body = json.loads(event.get('body') or '{}')
        pdf_b64 = body.get('pdf')
        if not pdf_b64:
            return {
                'statusCode': 400,
                'headers': headers_cors,
                'body': json.dumps({'erro': 'PDF não fornecido'}),
            }

        pdf_bytes = base64.b64decode(pdf_b64)
        pedidos = extrair_pedidos(pdf_bytes)

        if not pedidos:
            return {
                'statusCode': 422,
                'headers': headers_cors,
                'body': json.dumps({'erro': 'Nenhum pedido encontrado no PDF'}),
            }

        zip_bytes = _montar_zip(pedidos)
        zip_b64   = base64.b64encode(zip_bytes).decode('utf-8')

        resumo = [
            {
                'loja':        p['loja'],
                'razao_social': p['razao_social'],
                'num_pedido':  p['num_pedido'],
                'cnpj':        p['cnpj_formatado'],
                'total_itens': len(p['itens']),
                'data_compra': p['data_compra'],
                'data_entrega': p['data_entrega'],
            }
            for p in pedidos
        ]

        return {
            'statusCode': 200,
            'headers': {**headers_cors, 'Content-Type': 'application/json'},
            'body': json.dumps({'zip': zip_b64, 'pedidos': resumo}),
        }

    except Exception as exc:
        return {
            'statusCode': 500,
            'headers': headers_cors,
            'body': json.dumps({'erro': str(exc)}),
        }
