# Escala Médica Online — Hospital Cruzeiro do Sul (Hapvida)

Sistema automatizado que lê a planilha oficial do OneDrive local, gera o painel web (`index.html`) no Design System da Torre de Inteligência Hospitalar (Hapvida) e publica automaticamente via **GitHub Pages** com link fixo para os médicos.

## Arquivos incluídos
- `index.html`: Página final pronta para publicação (já gerada com os dados atuais da planilha).
- `atualizar_escala.py`: Script Python que lê a planilha `.xlsx`, filtra dados sensíveis (LGPD), gera o HTML e faz `git push`.
- `iniciar_monitor.bat`: Inicia o monitoramento contínuo em segundo plano (`--watch --git-push`).
- `atualizar_uma_vez.bat`: Executa uma atualização manual imediata + push.
- `hapvida_horizontal_colorido_texto-branco.svg`: Logo oficial utilizada no cabeçalho.

## Configuração Inicial (3 minutos)
1. Instale a dependência única:
   ```bash
   pip install openpyxl
   ```
2. Crie um repositório no GitHub (ex.: `escala-cruzeiro-do-sul`) e clone na sua máquina.
3. Copie todos os arquivos desta pasta para dentro do repositório clonado.
4. No GitHub, vá em **Settings > Pages > Build and deployment > Branch: `main` / `/ (root)` > Save**.
5. Faça o primeiro envio:
   ```bash
   git add .
   git commit -m "feat: primeira publicacao escala Hospital Cruzeiro do Sul"
   git push -u origin main
   ```
6. Seu link fixo estará no ar em ~1 minuto: `https://SEU_USUARIO.github.io/escala-cruzeiro-do-sul/`

## Uso Diário Automático
- Dê duplo clique em `iniciar_monitor.bat` (ou coloque um atalho dele em `Win + R` -> `shell:startup` para iniciar com o Windows).
- Sempre que você ou alguém editar a planilha no SharePoint/OneDrive e ela sincronizar no seu PC, o script detecta a alteração, gera o novo `index.html` e publica automaticamente.

## Conformidade LGPD
O script extrai apenas Nome, CRM/UF, datas, horários e carga horária. **CPF, CNPJ e valores financeiros são ignorados na leitura e nunca vão para o HTML ou GitHub.**
