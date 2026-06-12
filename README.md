# Conversor de Pedidos PDF → XLS / XML / CSV

Plataforma web corporativa da **Asa Branca Distribuidora** para conversão de pedidos de compra em PDF para os formatos de importação do sistema ION Força de Vendas: **XLS (Layout 5)**, **XML** e **CSV**.

---

## Funcionalidades

- Upload de PDF via drag & drop ou seletor de arquivo (máx. 10 MB)
- Detecção automática do formato do documento (4 parsers + fallback genérico)
- Suporte a PDFs de imagem (escaneados) via OCR — Tesseract.js + canvas
- Extração por loja: CNPJ, razão social, número do pedido, datas, itens
- Geração de `.xls` padrão e `.xls` completo por loja
- Geração de `.xml` estruturado por loja
- CSV consolidado com todos os pedidos
- Download individual por loja ou ZIP com todos os arquivos

---

## Stack Tecnológica

| Camada | Tecnologia |
|--------|-----------|
| Frontend | React 18 + Vite 5 + Tailwind CSS + shadcn/ui |
| Backend | Node.js 24 — Vercel Serverless Functions |
| Parser PDF | pdf-parse (pdfjs v2 interno) |
| OCR | Tesseract.js v7 + canvas |
| Excel | SheetJS (xlsx v0.18) |
| Compressão | JSZip |
| Deploy | Vercel (Fluid Compute, região iad1) |

---

## Formatos de PDF Suportados

| Formato | Detecção | Parser |
|---------|----------|--------|
| **Pedido de Compra padrão** | `RAZÃO SOCIAL:CNPJ:` | Regex por linha — extrai cabeçalho + itens |
| **Emitir Pedido de Compra** | `emitir pedido de compra` | Coordenadas X/Y + clustering por coluna |
| **Pedidos por Período (V. CISS)** | `pedidos de compra por periodo` | Coordenadas X/Y sem clustering global |
| **Ruptura Pré-Pedido** | `ruptura - pré pedido` | Named capture groups — colunas coladas |
| **Genérico / Fallback** | qualquer PDF não reconhecido | Multi-estratégia: X/Y → multi-espaço → posição fixa → chave-valor |
| **PDF de imagem** | sem camada de texto | OCR via Tesseract.js (português) |

---

## Saída — Layout 5 (ArquivoImportado ION)

### Estrutura do XLS

```
Linha 1:  cnpj  |  [CNPJ numérico 14 dígitos]
Linha 2:  codproduto | codembalagem | quantidade | descricao | emba | qtUnit |
          precoVenda | preço emba | preço emba st | preço unit | preço tot |
          preco tot ion | preco tot ion st
Linha 3+: dados dos itens
```

**Regras de tipos de célula** (exigidas pelo ION):
- `codembalagem`, `descricao`, `emba` → texto (`t="s"`)
- `codproduto`, `quantidade`, `qtUnit`, preços → número real (`t="n"`)
- Campos vazios → célula ausente (sem `t="s" v=""`)

O arquivo é gerado a partir do template `api/template.xls` (ArquivoImportado aceito pelo ION) para preservar os metadados OLE2 necessários (`AppVersion: 16.0000`, `CodePage`).

### Botões de download na interface

| Botão | Arquivo | Descrição |
|-------|---------|-----------|
| **XLS** | `pedido_{id}_{loja}.xls` | Campos padrão (alguns vazios para preenchimento manual) |
| **XLS+** | `pedido_{id}_{loja}_completo.xls` | Todos os campos preenchidos |
| **CSV** | `pedidos_combinados.csv` | Todos os pedidos, separador `;`, BOM UTF-8 |
| **ZIP** | `{nome}_convertido.zip` | XLS + XLS+ + XML de todos os pedidos |

---

## Estrutura do Projeto

```
pdf-conversor/
├── api/
│   ├── converter.js      # Handler HTTP da Vercel Function (CORS, body parser)
│   ├── _logic.js         # Motor de extração, parsers, geradores XLS/XML/CSV
│   └── template.xls      # Template OLE2/BIFF8 aceito pelo ION (base para geração)
├── src/
│   ├── App.jsx            # Componente raiz — máquina de estados upload→loading→result|error
│   ├── main.jsx
│   ├── index.css
│   ├── components/
│   │   ├── AppShell.jsx   # Shell (sidebar desktop + drawer mobile)
│   │   ├── UploadZone.jsx # Drag-and-drop com validação de tipo e tamanho
│   │   ├── ResultTable.jsx# Tabela de resultados com botões de download
│   │   └── ui/            # Componentes shadcn/ui (Badge, Button, Card, Table)
│   └── lib/utils.js       # cn() — clsx + tailwind-merge
├── public/
│   ├── logo.png           # Logo Asa Branca Distribuidora
│   ├── favicon.png
│   ├── icon-192.png
│   ├── icon-512.png
│   └── manifest.json      # PWA manifest
├── netlify/
│   └── functions/
│       └── converter.js   # Wrapper compatibilidade Netlify → /api/converter
├── vercel.json            # Config deploy: maxDuration, includeFiles (WASM + template)
├── netlify.toml           # Config compatibilidade Netlify
├── vite.config.js
├── tailwind.config.js     # Design tokens corporativos (azul #1F5FAE, etc.)
├── postcss.config.js
└── package.json
```

---

## Fluxo de Processamento

```
1. Browser: FileReader → base64
2. POST /.netlify/functions/converter (→ rewrite → /api/converter)
3. Buffer.from(base64)
4. pdfParse(buffer, { pagerender: getTextContent + coords })
   ├─ text: linhas reconstruídas por coordenada Y (ordem do stream preservada)
   └─ paginas[]: [{str, x, y}] por página
5. totalItens === 0 ou text.length < 20?
   └─ Sim → OCR: Tesseract.createWorker('por') + canvas + pdfjs render → texto
6. detectarFormato(text) → 'pedido' | 'emitir_pedido' | 'ciss_pedido' | 'ruptura' | 'generico'
7. Parser específico → pedidos[]
8. _montarXls() × 2 (padrão + completo) + gerarXml() + gerarCsv() + JSZip
9. Response: { pedidos, zip (base64), csv (base64) }
10. Browser: Blob → URL.createObjectURL → <a download>
```

---

## Desenvolvimento Local

### Pré-requisitos

- Node.js 20+
- Vercel CLI: `npm i -g vercel`

### Instalação

```bash
git clone https://github.com/jacksonRcosta/Leitor-Pdf.git
cd Leitor-Pdf
npm install
```

### Executar

```bash
# Frontend (Vite dev server)
npm run dev
# Acesse: http://localhost:5173

# Para testar as Vercel Functions localmente:
vercel dev
# Acesse: http://localhost:3000
```

---

## Deploy no Vercel

### Primeiro deploy

```bash
vercel login
vercel link --scope asa-branca-s-projects --project pdf-conversor
vercel --prod
```

### Deploy de atualização

```bash
vercel --prod --scope asa-branca-s-projects
```

### URLs de produção

| URL | Descrição |
|-----|-----------|
| `https://pdf-conversor-rho.vercel.app/` | Alias estável de produção |
| `https://pdf-conversor.app/` | Domínio customizado |

> **Atenção:** URLs no formato `pdf-conversor-XXXX-asa-branca-s-projects.vercel.app` requerem autenticação Vercel (Deployment Protection). Use sempre os aliases acima.

---

## Configurações Vercel (`vercel.json`)

```json
{
  "buildCommand": "npm run build",
  "outputDirectory": "dist",
  "functions": {
    "api/converter.js": {
      "maxDuration": 300,
      "includeFiles": "{node_modules/tesseract.js-core/**,api/template.xls}"
    }
  },
  "rewrites": [
    { "source": "/.netlify/functions/converter", "destination": "/api/converter" },
    { "source": "/((?!api/).*)", "destination": "/index.html" }
  ]
}
```

- `maxDuration: 300` — 5 min (suporta OCR de PDFs densos)
- `includeFiles` — inclui WASM do Tesseract e o template XLS no pacote Lambda (necessário pois o Vercel Node File Tracing não detecta imports dinâmicos de `.wasm`)
- Rewrite `/netlify/…` → compatibilidade dupla Netlify/Vercel sem alterar o frontend

---

## Limitações Conhecidas

| Item | Limite |
|------|--------|
| Tamanho máximo do PDF | 10 MB (15 MB no body parser) |
| Timeout total | 4 min 40 s (cliente) / 5 min (função) |
| Timeout OCR | 4 min |
| Páginas OCR | máx. 10 páginas por PDF |
| OCR cold start | ~30–60 s (download do modelo `por.traineddata` ~22 MB) |
| Idioma OCR | Português (`por`) |

---

## Design System

Tokens de cor definidos em `tailwind.config.js`:

| Token | Cor | Uso |
|-------|-----|-----|
| `primary` | `#1F5FAE` | Botões, links, active state |
| `success` | `#1E8E5A` | Badges de loja, botão XLS |
| `danger` | `#C0392B` | Erros, botão destrutivo |
| `warning` | `#C77D17` | Alertas |
| `bg` | `#F7F8FA` | Fundo da página |
| `surface` | `#FFFFFF` | Cards, sidebar |
| `text-muted` | `#6B7280` | Labels secundários |

Fontes: **Inter** (sans) + **JetBrains Mono** (mono, CNPJs e valores).

---

## Repositório

- **GitHub**: [jacksonRcosta/Leitor-Pdf](https://github.com/jacksonRcosta/Leitor-Pdf)
- **Branch principal**: `main`
- **Branch local**: `master` → push via `git push origin HEAD:main`
