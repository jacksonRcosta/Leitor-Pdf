# Conversor de Pedidos PDF → XLS / XML / CSV

Aplicação web para conversão de pedidos de compra em PDF para os formatos **XLS**, **XML** e **CSV**, no layout de importação padrão.

## Funcionalidades

- Upload de PDF via drag & drop ou seletor de arquivo
- Extração automática de pedidos por loja (CNPJ, razão social, itens)
- Geração de `.xls` e `.xml` individuais por loja + `.csv` consolidado
- Download de todos os arquivos em um único `.zip`

## Tecnologias

- **Frontend**: HTML / CSS / JavaScript (estático — hospedado no Netlify)
- **Backend**: Python via Netlify Functions (`pdfplumber`, `xlwt`)

## Estrutura do Projeto

```
pdf-conversor/
├── public/                  # Frontend estático (Netlify publish)
│   ├── index.html
│   ├── css/style.css
│   └── js/app.js
├── netlify/
│   └── functions/
│       ├── converter.py     # Netlify Function (Python)
│       └── requirements.txt
├── netlify.toml
├── .gitignore
└── README.md
```

## Deploy no Netlify

1. Faça o fork ou clone deste repositório no GitHub
2. Acesse [app.netlify.com](https://app.netlify.com) → **Add new site** → **Import from Git**
3. Selecione o repositório
4. As configurações de build já estão definidas no `netlify.toml`:
   - **Publish directory**: `public`
   - **Functions directory**: `netlify/functions`
5. Clique em **Deploy site**

## Desenvolvimento Local

### Pré-requisitos

- [Node.js](https://nodejs.org) (para o Netlify CLI)
- Python 3.12+
- [Netlify CLI](https://docs.netlify.com/cli/get-started/)

```bash
npm install -g netlify-cli
pip install pdfplumber xlwt
netlify dev
```

Acesse em `http://localhost:8888`.

## Limitações

- Tamanho máximo do PDF: **4 MB** (limite do Netlify Functions)
- Timeout da função: **10 segundos**
- Formato suportado: PDFs de pedidos de compra com texto nativo (não escaneados)

## Layout de Saída (XLS / XML)

O arquivo gerado segue o **Layout 5 — ArquivoImportado**:

| Coluna | Campo |
|--------|-------|
| A | codproduto |
| B | codembalagem |
| C | quantidade |
| D | descricao |
| E | emba |
| F | qtUnit |
| G | precoVenda |
| H | preço emba |
| I | preço emba st |
| J | preço unit |
| K | preço tot |
| L | preco tot ion |
| M | preco tot ion st |

> A linha 1 do XLS contém o **CNPJ do cliente** (razão social compradora).
