# 📄 Gerador de Proposta PDF — Ademicon Crédito Estruturado

Aplicativo web em Python + Streamlit que lê planilhas de simulação de crédito estruturado via consórcio imobiliário (formato Ademicon) e gera propostas profissionais em PDF com seções configuráveis.

---

## 🗂 Estrutura do Projeto

```
.
├── app.py              # Interface Streamlit + geração de PDF (ReportLab)
├── leitor.py           # Leitura e extração de dados das planilhas (.xlsx)
└── requirements.txt    # Dependências Python
```

---

## ⚙️ Requisitos

- Python 3.11+
- Pip

Instale as dependências com:

```bash
pip install -r requirements.txt
```

**Pacotes utilizados:**

| Pacote | Uso |
|---|---|
| `streamlit` | Interface web |
| `pandas` | Manipulação de dados tabulares |
| `openpyxl` | Leitura das planilhas `.xlsx` |
| `reportlab` | Geração do PDF |

---

## 🚀 Como Rodar

```bash
streamlit run app.py
```

O app abrirá automaticamente no navegador em `http://localhost:8501`.

---

## 📥 Entrada — Planilha de Simulação

O app aceita planilhas `.xlsx` no padrão Ademicon. São esperadas as seguintes abas:

### Aba `RESUMO`
- Total de crédito levantado
- Quantidade e valor de parcelas
- TIR mensal e anual
- Taxa estática
- Tabela de fluxo mensal resumido (mês, parcela, crédito, crédito acumulado)

### Aba `CARTEIRA`
- Grupos de consórcio selecionados (número do grupo, crédito contratado, prazo, lances, cotas, crédito líquido)
- Totais por seção (Lance Livre, Lance Fixo/Limitado)
- Prazo médio da carteira e percentual de lance fixo

### Aba `FLUXO` (nome pode variar, ex.: `FLUXO lance livre COM FIDC`)
- Premissas: crédito total, parcela, total de cotas, taxa FIDC, TIR
- Fluxo mensal detalhado: cotas contempladas, valor pago, lances, crédito liberado, crédito líquido acumulado

> **Detecção automática:** o leitor identifica campos por nome de coluna (não por posição fixa), tolerando variações de capitalização e espaços extras. Quando há múltiplas abas de Fluxo, a versão `COM FIDC` é priorizada.

---

## 🖥 Interface do App

### Sidebar — Configuração da Proposta

| Campo | Padrão |
|---|---|
| Nome do Cliente | *(obrigatório)* |
| Gerente Responsável | Julio Cesar Santos |
| Cargo | Gerente de Crédito Estruturado |
| Unidade | Ademicon Faria Lima |
| Data de Referência | Data atual |

**Seções da proposta (toggles on/off):**
- Resumo Executivo
- Custo FIDC
- Fluxo dos Primeiros 12 Meses
- Carteira de Cotas
- Detalhamento de Prazos por Grupo
- Disclaimer Padrão Ademicon

**Versão da TIR:** `Com FIDC` / `Sem FIDC` / `Ambas`

### Preview de Dados

Antes de gerar o PDF, o app exibe os dados extraídos em 3 abas:
- **Resumo** — métricas principais e tabela de fluxo mensal
- **Carteira** — tabela completa de grupos com totais
- **Fluxo** — fluxo detalhado com destaque verde nos meses de contemplação

---

## 📄 Saída — Proposta em PDF

Layout profissional gerado com ReportLab, contendo:

| Elemento | Descrição |
|---|---|
| **Cabeçalho** | Banner azul escuro com título da proposta |
| **Bloco de Info** | Cliente, gerente, data de referência, tipo de financiamento |
| **Seções** | Título em azul médio, tabelas formatadas com zebra |
| **Fluxo** | Meses de contemplação destacados em verde |
| **Rodapé** | Assinatura do gerente, cargo, unidade, número de página e data de geração |
| **Disclaimer** | Texto padrão Ademicon em itálico ao final (quando ativado) |

### Paleta de Cores

| Variável | Hex | Uso |
|---|---|---|
| Azul Escuro | `#1F3864` | Cabeçalho, títulos, rodapé |
| Azul Médio | `#2F75B6` | Banners de seção, bordas |
| Azul Claro | `#D6E4F0` | Fundo de tabelas e blocos de info |
| Verde Claro | `#E2EFDA` | Destaque de meses com contemplação |
| Cinza Claro | `#F5F5F5` | Linhas alternadas de tabelas |

---

## 🧩 Arquitetura do Código

### `leitor.py`

Contém três funções principais de extração, mais um ponto de entrada:

```
ler_planilha(file)   →  abre o workbook e chama as funções abaixo
  ├── ler_resumo(ws)    →  indicadores financeiros + fluxo resumido
  ├── ler_carteira(ws)  →  grupos, totais, prazo médio, % lance fixo
  └── ler_fluxo(ws)     →  premissas FIDC, TIR, fluxo mensal detalhado
```

Retorna um dicionário com chaves: `sheets`, `resumo`, `carteira`, `fluxo`, `fluxo_sheet_name`, `erros`.

### `app.py`

Dividido em três camadas:

1. **Helpers** — formatação de moeda, percentual, datas
2. **Builders de seção** — uma função por seção do PDF (`build_resumo_executivo`, `build_custo_fidc`, `build_fluxo_12m`, `build_carteira`, `build_prazos`, `build_disclaimer`)
3. **Interface Streamlit** — upload, preview, sidebar de configuração, botão de geração e download

---

## ⚠️ Observações

- O app não requer banco de dados nem autenticação.
- O PDF é gerado em memória e disponibilizado para download direto, sem salvar arquivos no servidor.
- Campos não encontrados na planilha são exibidos como `—` no PDF (sem travar a geração).
- Se o campo **Nome do Cliente** não estiver preenchido na sidebar, o botão de geração do PDF não é exibido.
