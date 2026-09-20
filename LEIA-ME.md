# App de Ordens de Serviço — Ponto dos Fogões

## O que foi criado
- **Nova O.S.**: cadastro com número automático, cliente, telefone, setor, atendente, até 6 itens (item/garantia/serviço/histórico/valor), entrada/taxa, status, valor total automático e fotos (câmera ou galeria).
- **Cliente pelo telefone**: ao digitar o telefone na Nova O.S., o app busca sozinho na aba `Cadastro` (só leitura, nunca escreve lá) e preenche o nome do cliente automaticamente se ele já existir. Se não encontrar, mostra "Cliente novo" e o atendente digita normalmente — nada é criado na aba `Cadastro` pelo app, então não há risco de duplicar cadastro.
- **Dashboard completo**: gráfico de barras colorido com a quantidade de O.S. por status, filtros por setor/status/prazo, e uma lista de acompanhamento das ordens em aberto ordenada por urgência.
- **Alertas de prazo**: dois cartões grandes no topo do Dashboard — 🔴 vencidos e 🟡 perto do prazo — clicáveis para filtrar a lista na hora. Prazos usados:
  - **Envio de orçamento** (a partir da Data de Entrada): Panelas 5 dias · Liquidificadores 15 dias · Elétricos 15 dias.
  - **Finalização do serviço** (a partir da Data de Aprovação): Panelas 5 dias · Liquidificadores 10 dias · Elétricos 10 dias.
  - Se precisar encomendar peça, mude o status para **AGUARDANDO PEÇA** — o prazo de finalização pausa automaticamente até voltar para PRONTO.
- **Status com fluxo controlado**: ABERTA → AGUARDA APROVAÇÃO → APROVADO → (opcional: AGUARDANDO PEÇA) → PRONTO → RETIRADO, com o desvio NÃO APROVADO → RETIRADO. Cancelar tem um botão próprio, com confirmação e motivo.
- **Histórico**: toda mudança de status, item adicionado, foto adicionada ou dado cadastral editado fica registrada com data/hora e quem fez, numa aba nova (HISTORICO_OS) e visível como linha do tempo dentro de cada O.S.
- **Área do Técnico** (`tecnico.html`, tela separada): acesso por PIN, restrito a Wilgler (setor PANELAS) e Isackson (setor ELÉTRICOS). Cada um só vê as O.S. do próprio setor, **sem telefone e sem nenhum valor**. Pode editar item/serviço/histórico dos itens já cadastrados, e só pode mudar o status quando ele estiver em **APROVADO**, passando direto para **FINALIZADO**.
- **Editar dados cadastrais**: ícone de lápis ✏️ na tela de detalhe da O.S. permite corrigir Cliente, Telefone, Setor, Atendente, **Responsável Técnico e Entrada/Taxa** — esses dois últimos agora aparecem no resumo da O.S. e podem ser definidos ou trocados a qualquer momento (antes só existiam na criação). A **Data de Aprovação** também aparece no resumo (é preenchida sozinha quando o status vira APROVADO).
- **Enviar WhatsApp**: escolha um dos 4 modelos (Reserva, Orçamento, Pronto, Não aprovado), **edite o texto livremente** (inclusive com botões de emoji rápido) antes de mandar, ouça em voz alta pra conferir, e envie — abre o WhatsApp com a mensagem pronta pro telefone da O.S. Toda mensagem enviada fica registrada com data/hora numa aba nova (**WHATSAPP_ENVIOS**) e também aparece na linha do tempo da O.S.
  - *Sobre mensagens de áudio*: o link do WhatsApp só permite pré-preencher texto — não existe forma (em nenhum sistema) de anexar um áudio pronto automaticamente. O botão "🔊 Ouvir" lê a mensagem em voz alta no aparelho, útil pra conferir antes de enviar ou pra passar por telefone.
- **Atendente**: agora é uma lista fixa (Thais Alves, Thais Rangel, Márcia Rangel, Wilgler Martins, Thaíssia Rangel) em vez de texto livre, tanto na criação quanto na edição.
- **Telefone com +55**: o telefone é gravado automaticamente com o código do Brasil, pra funcionar direto se você copiar pro WhatsApp.
- **Valor Total**: essa coluna já tinha uma fórmula própria na planilha (soma dos 6 itens). O app **não escreve mais nela** — só lê o que a fórmula calcula, pra não correr o risco de apagar a fórmula.
- **Consultar**: busca por nº da O.S., nome do cliente **ou telefone**, tela de edição (status, adicionar novo item, adicionar fotos).
- **Prazo de garantia**: quando um item tem garantia "Sim", aparece um campo para informar o prazo (ex: "90 dias", "1 ano").
- **Offline**: se não houver internet, a O.S. fica salva no tablet e é enviada sozinha quando a conexão voltar (mostra "X pendente(s)" no topo).
- **Planilha**: nenhuma coluna ou aba existente é alterada. São criadas apenas colunas novas no final da aba principal (**FOTOS**, **PRAZO DE GARANTIA** por item) e uma aba nova (**HISTORICO_OS**).

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
2. Hospede os arquivos (`index.html`, `tecnico.html`, `manifest.json`, `service-worker.js`, `icon-192.png`, `icon-512.png`) no GitHub Pages — mesmo esquema já usado no app de ponto eletrônico, para evitar bloqueio de CORS.
3. No tablet/celular de cada técnico, abra o link `.../tecnico.html` (não o `index.html`) no Chrome e use **Adicionar à tela inicial**.
   Para Thais Rangel e Thais Alves, o link normal é o `index.html` (acesso completo).

## Observações
- O valor total é calculado automaticamente somando os valores dos itens 1 a 6.
- Cada O.S. aceita até 6 itens, igual à planilha atual.
- As fotos são comprimidas no navegador antes de enviar (para não pesar no tablet/conexão).
- Se o fluxo de status precisar de ajuste (ex: pular uma etapa, adicionar um novo status), edite o `STATUS_FLOW` no topo do `Code.gs` — e o mesmo objeto, duplicado no `index.html`, para manter os dois sincronizados.
- Se quiser, depois dá pra evoluir: notificação por WhatsApp, geração de PDF da O.S., controle de estoque/peças, financeiro, permissões por usuário.

## Importante ao atualizar o Code.gs
⚠️ **Substitua o arquivo INTEIRO, sempre.** Detectamos que uma atualização anterior ficou parcial (só um trecho foi colado por cima do código antigo), o que fez ordens novas caírem em colunas erradas por dias sem ninguém perceber. Para evitar isso:
1. No editor do Apps Script, clique dentro do código e aperte **Ctrl+A** (selecionar tudo) e **Delete**.
2. Cole o conteúdo completo do `Code.gs` novo.
3. Role até perto do topo e confira se a linha `ATENDENTE: 6,` aparece dentro do bloco `const COL = {` — se aparecer, colou certo.
4. **Implantar → Gerenciar implantações → (lápis) → Nova versão → Implantar**.

E o mesmo vale pro `index.html`: substitua o arquivo inteiro no GitHub, não só parte dele.
