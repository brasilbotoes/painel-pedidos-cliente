# Painel de Pedidos — Brasil Botões (versão HTML)

Dashboard responsivo (notebook + celular) que reproduz os KPIs da aba **Dashboard**
da sua planilha, com filtros de Mês / Representante / Cliente e tabelas detalhadas
clicáveis. Testado: os números batem exatamente com os da planilha original.

## Arquivos

| Arquivo | Para quê serve |
|---|---|
| `index.html` | O dashboard em si. Não precisa editar. |
| `data.json` | Os dados (gerado automaticamente). Já vem com uma cópia atual. |
| `atualizar_dashboard.py` | Abre o Excel, atualiza os ODBCs e regera o `data.json`. |
| `publicar.bat` | Roda o script acima e envia a atualização para o site. |

---

## Passo 1 — Testar localmente (2 min)

1. Abra a pasta `dashboard` no notebook.
2. Dê duplo clique no `index.html`.
3. **Importante:** navegadores bloqueiam a leitura do `data.json` quando o arquivo
   é aberto direto (`file://`). Para testar localmente, abra um terminal na pasta e rode:
   ```
   python -m http.server 8000
   ```
   e acesse `http://localhost:8000` no navegador. Isso é só para teste — uma vez publicado
   no GitHub/Cloudflare Pages, funciona normalmente ao abrir o link.

## Passo 2 — Publicar (link único, sem SharePoint)

**Opção recomendada: Cloudflare Pages + Cloudflare Access**
Como nem todo mundo que vai acessar é da empresa, isso permite restringir quem abre
o link (por e-mail, com código de uso único) sem precisar de senha nem de conta
`@brasilbotoes`. É gratuito até 50 usuários.

1. Crie uma conta gratuita em https://dash.cloudflare.com
2. Em **Workers & Pages → Create → Pages → Upload assets**, envie a pasta `dashboard`
   (ou conecte um repositório GitHub para publicar automaticamente a cada atualização —
   veja Passo 3).
3. Anote o link gerado (ex: `brasilbotoes-pedidos.pages.dev`).
4. Em **Zero Trust → Access → Applications → Add an application → Self-hosted**,
   aponte para esse domínio e crie uma política liberando apenas a lista de e-mails
   das pessoas que devem ver o painel (funcionários e não funcionários). Ative o
   método de login **One-time PIN**.

**Opção mais simples: GitHub Pages**
Mais rápido de configurar, porém o link fica público na internet para qualquer
pessoa que o receba (não aparece em buscas, mas não tem controle de quem acessa).
Se os dados não forem sensíveis a esse nível, funciona bem:

1. Crie um repositório no GitHub (ex: `dashboard-brasil-botoes`), pode ser público.
2. Suba os arquivos desta pasta para o repositório.
3. Em **Settings → Pages**, selecione a branch `main` e salve.
4. Em alguns minutos o site estará em `https://SEU_USUARIO.github.io/dashboard-brasil-botoes/`.

## Passo 3 — Automatizar a atualização (sem você mexer na planilha)

Como você me confirmou que não há um servidor sempre ligado — só o seu notebook em
horário comercial — a atualização automática funciona assim: o Windows dispara o
script sozinho enquanto seu notebook está ligado, sem você precisar abrir a
planilha ou clicar em "Atualizar". Fora do horário em que o notebook está ligado,
o site simplesmente mantém os últimos dados publicados.

1. Instale as dependências (uma vez só), no Prompt de Comando:
   ```
   pip install xlwings openpyxl
   ```
2. Instale o **Git** (https://git-scm.com/downloads) e configure o repositório local:
   ```
   cd C:\Users\SEU_USUARIO\Documents
   git clone https://github.com/SEU_USUARIO/dashboard-brasil-botoes.git
   ```
   Copie `index.html`, `data.json`, `atualizar_dashboard.py` e `publicar.bat` para
   dentro dessa pasta clonada.
3. Abra `atualizar_dashboard.py` e ajuste as duas linhas em **CONFIGURACOES**:
   - `CAMINHO_PLANILHA`: onde está o `.xlsm` original no seu notebook.
   - `PASTA_DASHBOARD`: a pasta do repositório clonado no passo 2.
4. Abra `publicar.bat` e ajuste o `cd /d` para a mesma pasta do repositório.
5. Teste manualmente: dê duplo clique em `publicar.bat`. Ele deve abrir o Excel
   escondido, atualizar os dados e enviar para o GitHub.
6. Agende no **Agendador de Tarefas do Windows**:
   - Abra "Agendador de Tarefas" → **Criar Tarefa Básica**
   - Gatilho: repetir a cada 2 horas, das 8h às 18h (ajuste como preferir)
   - Ação: iniciar o programa `publicar.bat`
   - Pronto — a partir daí roda sozinho enquanto o notebook estiver ligado.

## Sobre os números

Os 12 indicadores (Vendas e Amostras) foram extraídos das fórmulas exatas da sua
aba Dashboard e reproduzidos em JavaScript — incluindo o detalhe de que, na seção
Amostras, "Qtde. Faturados", "Valor Faturado" e os indicadores de entrega usam
**Data Faturamento** para o filtro de mês, enquanto o restante usa **Data Emissão
Pedido** — exatamente como na planilha original. Os valores foram conferidos e
batem 100% com o arquivo enviado.

## Próximos passos possíveis (se quiser depois)
- Adicionar gráfico de evolução mensal.
- Exportar tabela filtrada para Excel/CSV direto do navegador.
- Enviar um aviso (e-mail/Slack) automaticamente se "Pedidos em atraso" passar de um limite.
