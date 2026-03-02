# Finanças do Casal — Contexto do Projeto

## Objetivo
App mobile (PWA) para controle financeiro do casal, hospedado no GitHub Pages.
Permite lançar receitas/despesas rapidamente e sincroniza automaticamente com uma planilha Excel no OneDrive.

---

## Stack e Bibliotecas

| Item | Detalhe |
|------|---------|
| Hospedagem | GitHub Pages (sem backend) |
| Auth | MSAL.js v2.39.0 (Microsoft OAuth 2.0) |
| Graph API | Microsoft Graph v1.0 (OneDrive + Excel) |
| Excel local | SheetJS (xlsx v0.18.5) — para exportação local |
| Azure App | Client ID: `d1770348-0816-49b0-b2c0-e22b2c1220c3` |
| Dados locais | `localStorage` chave `financas_casal_v2` |

---

## Estrutura de Arquivos

```
/
├── financas-casal.html   # App completo (UI + lógica + MSAL)
├── excelSync.js          # Módulo de sincronização com tabela Excel
└── CONTEXT.md            # Este arquivo
```

---

## Decisões Técnicas

### Por que Table API em vez de PATCH de células?
- A Table API (`/workbook/tables/{name}/rows/add`) adiciona uma linha nova a cada lançamento, criando um log de transações.
- O PATCH de células individuais acumula totais mas perde o histórico de lançamentos.
- A Table API é mais robusta: não depende de posições fixas de célula.

### Por que share link em vez de path?
- O arquivo Excel está em uma conta OneDrive diferente da conta Azure do app.
- O share link resolve para o `driveId` + `itemId` corretos via `/shares/{shareId}/driveItem`, independente de qual conta é dona do arquivo.
- O `shareId` é gerado como `"u!" + base64url(shareUrl)`.

### Por que cache em localStorage?
- Evita resolver o share link e listar tabelas a cada lançamento (cada resolução = 2 chamadas de API).
- Cache pode ser limpo chamando `resetExcelSyncCache()` no console do navegador.

### Fila offline
- Se não há conexão ou token, o lançamento vai para a fila (`excelSync_queue` no localStorage).
- Quando o browser volta online, `flushExcelQueue()` é chamado automaticamente.

---

## Modelo de Dados (transação)

```javascript
{
  id:          "uuid",       // identificador único
  date:        "2026-03-15", // ISO date
  month:       "2026-03",    // YYYY-MM
  person:      "marido",     // "marido" | "esposa" | "casa"
  type:        "despesa",    // "despesa" | "receita"
  category:    "Alimentação",
  description: "Mercado",
  amount:      150.00
}
```

---

## Tabela Excel Esperada

O `excelSync.js` auto-detecta qualquer tabela no workbook.
Se não existir nenhuma, cria automaticamente a aba **"Lançamentos"** com a tabela **"TbLancamentos"** e as colunas:

```
Data | Pessoa | Tipo | Categoria | Descrição | Valor
```

O mapeamento de colunas é feito por nome (case-insensitive, ignora acentos).
Colunas desconhecidas recebem string vazia.

---

## Como Rodar / Testar / Publicar

### Rodar localmente
```bash
# Qualquer servidor HTTP simples serve — sem build step
npx serve .
# ou
python3 -m http.server 8080
```

### Testar sincronização manualmente
1. Abra o app no celular (URL do GitHub Pages)
2. Digite o PIN
3. Se necessário: clique em "Entrar com Microsoft" e use a conta com o OneDrive
   - **Nota:** se o botão não aparecer, o MSAL já fez login silencioso (comportamento correto)
4. Salve um lançamento
5. Aguarde ~2 segundos e abra a planilha Excel no OneDrive
6. Verifique que apareceu uma nova linha na tabela "TbLancamentos" (aba "Lançamentos")

### Publicar no GitHub Pages
1. Faça upload dos arquivos `financas-casal.html`, `excelSync.js` e `CONTEXT.md` para o repositório
2. GitHub Pages publica automaticamente em 1-2 minutos

---

## O Que Já Foi Feito

- [x] App mobile com PIN, UI em PT-BR, 4 telas (Início, Lançamentos, Resumo, Metas)
- [x] Login Microsoft via MSAL.js
- [x] Backup JSON no OneDrive (via PUT de arquivo)
- [x] Escrita de totais nas células da planilha existente (via PATCH)
- [x] **Integração Table API: append de linha a cada lançamento (excelSync.js)**
- [x] Fila offline com flush automático ao voltar online
- [x] Cache de itemId/tableName para eficiência

---

## Como Reverter a Integração Excel

Se algo quebrar, basta remover do `financas-casal.html`:
1. A linha `<script src="excelSync.js"></script>`
2. O bloco `syncExcelRow(txSnapshot).then(...)` dentro de `saveLancamento()`

O app continua funcionando normalmente com localStorage + backup JSON.

---

## Troubleshooting

| Problema | Solução |
|----------|---------|
| Linha não aparece na planilha | Verifique console do navegador por erros `[excelSync]` |
| "Não foi possível acessar o arquivo compartilhado" | Link do share expirou — atualize `EXCEL_SHARE_URL` no `excelSync.js` |
| Tabela errada sendo usada | Defina `TABLE_NAME_CONFIG = 'NomeDaTabela'` no `excelSync.js` |
| Quero forçar re-detecção | Execute `resetExcelSyncCache()` no console do navegador |
| Lançamentos pendentes acumulando | Verifique `localStorage.getItem('excelSync_queue')` no console |

---

## Próxima Tarefa Sugerida

- [ ] Mostrar badge/contador de lançamentos pendentes na UI
- [ ] Botão "Sincronizar agora" na tela de Resumo
- [ ] Testar com tabela já existente na planilha (renomear `TABLE_NAME_CONFIG`)
