# Ri Happy | Gestão e Agendamento de Entregas e Recebimento de Cargas

Sistema de gestão logística para controle de agendamento de entregas de fornecedores e recebimento de cargas no Centro de Distribuição / Lojas da **Ri Happy**.

## 🚀 Tecnologias

- **Frontend**: HTML5, Tailwind CSS, Google Fonts, Material Symbols, Vanilla JavaScript
- **Backend / API**: Google Apps Script (`Code.gs`)
- **Banco de Dados**: Google Sheets (Planilha de Agendamentos e Divergências)
- **Hospedagem Frontend**: Vercel / GitHub Pages

## 📁 Estrutura do Projeto

- `index.html`: Interface web (SPA) responsiva com suporte a temas, painel de notificações, calendário de capacidade e formulário de check-in.
- `Code.gs`: Código do Google Apps Script com funções de banco de dados, regras de negócio e endpoints para API REST JSON (`doGet` / `doPost`).

## 🛠️ Configuração e Implantação

### 1. Backend (Google Apps Script)
1. Crie uma planilha no Google Sheets com as abas `Agenda` e `Divergências`.
2. Vá em **Extensões** > **Apps Script** e adicione o conteúdo do arquivo `Code.gs`.
3. Clique em **Implantar** > **Nova Implantação** > **App da Web**.
4. Configure:
   - **Executar como**: Eu
   - **Quem tem acesso**: Qualquer pessoa
5. Copie o URL do App da Web gerado.

### 2. Frontend (Vercel / GitHub)
1. No arquivo `index.html`, insira o URL do App da Web na constante `API_URL_GAS`.
2. Conecte este repositório do GitHub à **Vercel** para implantação automática.
