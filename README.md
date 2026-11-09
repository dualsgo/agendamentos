# 📦 Sistema de Gestão Logística, Agendamentos & Portal de Promotores — Ri Happy

Sistema integrado de **Gestão de Recebimento de Cargas, Agendamento de Fornecedores, Controle de Divergências e Portal de Promotores** desenvolvido para operações de loja e estoque da **Ri Happy**.

---

## 🚀 Principais Funcionalidades

### 1. 📅 Painel de Controle de Agendamentos (`index.html`)
- **Fluxo Completo de Recebimento**: Acompanhamento de entregas nos status `PENDENTE`, `EM TRATATIVA`, `CONFIRMADO`, `CONCLUÍDO` e `CANCELADO`.
- **Trava de Capacidade & Bloqueio de Inventário**: Limite configurável de volumes diários (`limite_vols`) e bloqueio automático de agendamentos na véspera, no dia e no dia seguinte ao inventário geral.
- **Baixa Física em Lote e Check-in com Divergência**: Registro de volumes recebidos, alerta de entrada fiscal no SAP e apontamento detalhado de faltas/sobras por Código SAP e Nota Fiscal.
- **Integração com E-mail (Gmail)**: Resposta direta ao remetente/transportadora a partir do painel via Google Apps Script (`responderEmail`).

### 2. 🤝 Portal de Promotores (`promotores.html` / `portal-promotores/`)
- Interface dedicada para registro, consulta e acompanhamento de atendimento de promotores e abastecimento em loja.

### 3. ⚙️ Backend Serverless (`Code.gs`)
- API REST JSON construída no **Google Apps Script** (`doGet` / `doPost`) conectada ao **Google Sheets**.
- Controle de concorrência com `LockService`, deduplicação automática por ID e log de auditoria em `ScriptProperties`.

---

## 🛠️ Arquitetura e Tecnologias

- **Frontend**: HTML5, Vanilla JavaScript (ES6+), Tailwind CSS, Google Fonts (*Manrope* & *Inter*) e Material Symbols.
- **Backend**: Google Apps Script (`Code.gs`).
- **Banco de Dados**: Google Sheets (abas `Agenda` e `Divergências`) + `PropertiesService`.

---

## 📂 Estrutura do Repositório

```text
agendamentos/
├── .agents/
│   └── AGENTS.md               # Especificação técnica de arquitetura e schemas
├── portal-promotores/
│   └── index.html              # Módulo standalone do Portal de Promotores
├── Code.gs                     # Backend Google Apps Script (API e regras de negócio)
├── index.html                  # SPA principal de Gestão de Agendamentos e Recebimento
├── promotores.html             # Interface do Portal de Promotores
└── LICENSE
```

---

## 💻 Como Configurar e Executar

1. **Backend (Google Apps Script)**:
   - Crie uma planilha no Google Sheets com as abas `Agenda` e `Divergências`.
   - Em **Extensões > Apps Script**, cole o conteúdo de `Code.gs` e publique como **App da Web** (*Executar como: Eu*, *Quem pode acessar: Qualquer pessoa*).
2. **Frontend**:
   - Configure a constante `API_URL_GAS` no `index.html` com a URL gerada pelo Web App.
   - Hospede os arquivos estáticos no Vercel, GitHub Pages ou abra localmente no navegador.

---

## 🔒 Privacidade e Dados
Planilhas e arquivos `.csv` contendo dados reais de notas fiscais, transportadoras ou e-mails de fornecedores não são versionados neste repositório.
