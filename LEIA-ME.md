# App de Ordens de Serviço — Ponto dos Fogões

## O que foi criado
- **Nova O.S.**: cadastro com número automático, cliente, telefone, setor, atendente, até 6 itens (item/garantia/serviço/histórico/valor), entrada/taxa, status, valor total automático e fotos (câmera ou galeria).
- **Dashboard**: aba nova com contagem de O.S. por status e lista das últimas ordens criadas.
- **Status com fluxo controlado**: a O.S. nasce em "ABERTO" e só pode avançar seguindo a ordem do processo (recebimento → avaliação → orçamento → aprovação → execução → finalização → entrega). Não dá mais para pular etapas por engano. Cancelar ou descartar continua liberado em qualquer etapa não finalizada.
- **Histórico**: toda mudança de status, item adicionado ou foto adicionada fica registrada com data/hora e quem fez, numa aba nova (HISTORICO_OS) e visível como linha do tempo dentro de cada O.S.
- **Consultar**: busca por nº da O.S., nome do cliente **ou telefone**, tela de edição (status, adicionar novo item, adicionar fotos).
- **Prazo de garantia**: quando um item tem garantia "Sim", aparece um campo para informar o prazo (ex: "90 dias", "1 ano").
- **Offline**: se não houver internet, a O.S. fica salva no tablet e é enviada sozinha quando a conexão voltar (mostra "X pendente(s)" no topo).
- **Planilha**: nenhuma coluna ou aba existente é alterada. São criadas apenas colunas novas no final da aba principal (**FOTOS**, **PRAZO DE GARANTIA** por item) e uma aba nova (**HISTORICO_OS**), gerada automaticamente na primeira vez que o app for usado.

## Passo 1 — Backend (Google Apps Script)
1. Abra a planilha "Ordem de serviço - PDF - NOVA".
2. Confirme, na parte de baixo, o nome exato da aba com os dados (ex: "Ordem de Serviços").
3. No menu **Extensões > Apps Script**, apague o conteúdo padrão e cole o arquivo `apps-script-Code.gs`.
4. No topo do código, ajuste:
   - `SHEET_NAME` — nome exato da aba (passo 2).
   - `DRIVE_FOLDER_ID` — crie uma pasta no Google Drive para as fotos e cole o ID dela (está na URL da pasta).
5. Clique em **Implantar > Nova implantação**:
   - Tipo: **App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
6. Copie a **URL do app da web** gerada (termina em `/exec`).

## Passo 2 — Frontend (o app em si)
1. Abra `index.html` e substitua a linha:
   ```js
   const GAS_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```
   pela URL copiada no passo anterior.
2. Hospede os arquivos (`index.html`, `manifest.json`, `service-worker.js`, `icon-192.png`, `icon-512.png`) no GitHub Pages — mesmo esquema já usado no app de ponto eletrônico, para evitar bloqueio de CORS.
3. No tablet, abra o link do GitHub Pages no Chrome e use **Adicionar à tela inicial** para instalar como app.

## Observações
- O valor total é calculado automaticamente somando os valores dos itens 1 a 6.
- Cada O.S. aceita até 6 itens, igual à planilha atual.
- As fotos são comprimidas no navegador antes de enviar (para não pesar no tablet/conexão).
- Se o fluxo de status precisar de ajuste (ex: pular uma etapa, adicionar um novo status), edite o `STATUS_FLOW` no topo do `Code.gs` — e o mesmo objeto, duplicado no `index.html`, para manter os dois sincronizados.
- Se quiser, depois dá pra evoluir: notificação por WhatsApp, geração de PDF da O.S., controle de estoque/peças, financeiro, permissões por usuário.

## Importante ao atualizar o Code.gs
Depois de colar uma versão nova do `Code.gs`, é preciso **atualizar a implantação** para a mudança valer na URL já em uso:
**Implantar > Gerenciar implantações > (ícone de lápis) > Nova versão > Implantar**.
Só colar o código novo e salvar não é suficiente — sem esse passo, o app continua usando a versão antiga.
