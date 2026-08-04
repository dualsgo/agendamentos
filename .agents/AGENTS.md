# Contexto e Regras do Projeto - Ri Happy Logistics & Agendamentos

## 📌 Visão Geral do Projeto
Sistema de Gestão Logística e Agendamento de Entregas/Recebimento de Cargas para os Centros de Distribuição e Lojas da **Ri Happy**.
O sistema gerencia todo o fluxo desde a entrada de e-mails/formulários de agendamento de fornecedores, confirmação de datas com trava de capacidade e bloqueio de inventário, até a baixa física no estoque com controle de divergências de volumes/peças.

---

## 🛠️ Arquitetura & Tecnologias
- **Frontend**: SPA responsiva em `index.html` construída com HTML5, Vanilla JavaScript, Tailwind CSS (via CDN), Google Fonts (*Manrope* e *Inter*) e Material Symbols.
- **Hospedagem Frontend**: Vercel / GitHub Pages.
- **Backend / API**: Google Apps Script (`Code.gs`) atuando como endpoint REST JSON via funções `doGet` e `doPost`.
- **Banco de Dados / Armazenamento**: Google Sheets (Planilha de Agendamentos e Divergências) + `PropertiesService.getScriptProperties()` para configurações e auditoria.

---

## 📊 Estrutura dos Dados (Schemas)

### Aba `Agenda` (Planilha Principal)
1. **Data Real** (`Col 1`): Data/Hora da conclusão física/recebimento (`dd/MM/yyyy HH:mm:ss`).
2. **Fornecedor** (`Col 2`): Nome do fornecedor/transportadora.
3. **Notas Fiscais** (`Col 3`): Número(s) da(s) nota(s) fiscal(is).
4. **Volumes** (`Col 4`): Quantidade total de volumes agendados.
5. **Data Agendada** (`Col 5`): Data prevista para entrega (`dd/MM/yyyy`).
6. **Observações** (`Col 6`): Notas ou detalhes adicionais.
7. **Status** (`Col 7`): Status atual (`PENDENTE`, `EM TRATATIVA`, `CONFIRMADO`, `CONCLUÍDO`, `CANCELADO`).
8. **Origem** (`Col 8`): Origem da solicitação (ex: *Formulário Web*, *Agendamento Manual*, *Automação / E-mail*).
9. **ID** (`Col 9`): Identificador único do agendamento (UUID ou prefixado ex: `MANUAL_...`).
10. **Remetente** (`Col 10`): E-mail do solicitante.
11. **Assunto** (`Col 11`): Assunto da mensagem/e-mail original.
12. **Volumes Recebidos** (`Col 12`): Quantidade real de volumes recebidos na conferência.
13. **Resumo Direto** (`Col 13`): Resumo sintético extraído da solicitação.

### Aba `Divergências`
1. **Data Registro** (`Col 1`): Data em que a divergência foi apontada.
2. **Código SAP** (`Col 2`): Código do produto no SAP.
3. **Volumes Faltantes** (`Col 3`): Quantidade de volumes faltantes.
4. **Peças Faltantes** (`Col 4`): Quantidade de peças faltantes.
5. **NF** (`Col 5`): Nota Fiscal associada.
6. **Fornecedor** (`Col 6`): Nome do fornecedor.

### ScriptProperties (`CONFIG`)
- `limite_vols`: Limite de capacidade diária de volumes (padrão: 1000).
- `data_inventario`: Data do próximo inventário geral (formato ISO `YYYY-MM-DD`).
- `historico_alteracoes`: Log dos últimos 500 registros de alteração por campo/usuário.
- `log`: Log dos últimos 100 eventos e ações da API.

---

## ⚙️ API Endpoints & Métodos (`Code.gs`)
As requisições frontend são disparadas contra a URL do Web App do Google Apps Script (`API_URL_GAS`):

| Ação (`action`) | Método | Parâmetros | Descrição |
|---|---|---|---|
| `getTodosOsDados` | GET / POST | N/A | Retorna `{ alertas, agendamentos, divergencias }` consolidados por ID. |
| `getSettings` | GET / POST | N/A | Retorna `{ limite_vols, data_inventario }`. |
| `salvarSettings` | POST | `[ { limite_vols, data_inventario } ]` | Atualiza limites e inventário no ScriptProperties. |
| `salvarAgendamento` | POST | `[ dadosAgendamento ]` | Cria ou edita um agendamento. |
| `atualizarStatus` | POST | `[ id, novoStatus ]` | Atualiza o status do agendamento. |
| `atualizarDataSugerida` | POST | `[ id, novaData ]` | Altera a data agendada de um registro. |
| `registrarCheckinComDivergencia` | POST | `[ id, dadosCheckin ]` | Registra baixa com apontamento de divergências de itens. |
| `marcarRecebidosLote` | POST | `[ idsJSON/Str, dataRealStr ]` | Dá baixa concluída em lote. |
| `aprovarLote` | POST | `[ idsJSON/Str ]` | Aprova e altera status para CONFIRMADO em lote. |
| `excluirAlertaFlow` | POST | `[ id ]` | Exclui linha do agendamento. |
| `excluirDivergencia` | POST | `[ index ]` | Exclui item da aba Divergências. |
| `responderEmail` | POST | `[ id, mensagem ]` | Envia resposta ao remetente via `GmailApp`. |

---

## 🧠 Regras de Negócio Importantes
1. **Consolidação por ID**: Em `getTodosOsDados()`, registros com o mesmo ID são agrupados e o registro da linha mais recente é mantido como oficial.
2. **Concorrência (`LockService`)**: Todas as mutações na planilha utilizam `LockService.getScriptLock().waitLock(15000)` para evitar escrita simultânea e inconsistência.
3. **Bloqueio por Inventário**: Entregas são automaticamente sinalizadas/bloqueadas no dia do inventário, no dia anterior e no dia posterior.
4. **Alerta de Lembrete SAP**: Após baixa física, o sistema exibe um lembrete para dar entrada fiscal no SAP.
5. **Formatação Flexível de Datas**: O parser de datas (`parseSafeDate` e `formatarDataStudio`) aceita múltiplos formatos de entrada, incluindo datas do Looker Studio e nomes de meses em português/inglês.
