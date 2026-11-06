# Controle de saída e empréstimo de itens

Aplicação web do Google Apps Script para registrar entregas individuais e coletivas, empréstimos e devoluções. Uma submissão aceita vários itens e grava uma linha por item.

## Recursos

- Destinatário colaborador ou setor para uso coletivo.
- Vínculo próprio ou terceirizado; a matrícula própria é informada manualmente e não é consultada na planilha.
- Setores e funções carregados da aba `MATRIZ POR SETOR`, com opção para setores assistenciais pontuais.
- EPIs individuais limitados aos itens previstos para setor e função. Para setor/uso coletivo, a lista inclui avental plumbífero, protetor de tireóide, cinta ergonômica e protetor auditivo de espuma.
- Entrega inicial, substituição, empréstimo com previsão de devolução e registro de devolução por item emprestado.
- Itens adicionais selecionados do catálogo `CÓD. ALMOX`.
- Tamanho padronizado e quantidade validada como inteiro positivo. CA e observações são guardados por item.
- Grava o histórico completo na aba `MOVIMENTAÇÕES EPI` e prepara uma linha por item na aba `CONTROLE ALMOX`, com caixa de seleção para marcar o lançamento.

## Configuração

1. Abra o projeto Apps Script vinculado à planilha de controle.
2. Adicione `Code.gs`, `Index.html` e este `appsscript.json` ao projeto.
3. Em **Configurações do projeto > Propriedades do script**, crie `SPREADSHEET_ID` com o ID da planilha de controle. O ID não deve ser commitado no repositório.
4. Execute `setupEpiModule()` uma vez. A função prepara `MOVIMENTAÇÕES EPI`, `CONTROLE ALMOX` e `MAPEAMENTO ALMOX`, e instala os acionadores para ocultar linhas marcadas como lançadas e adicionar o menu de exibição.
5. Na aba `MAPEAMENTO ALMOX`, relacione cada EPI padronizado ao produto correto do almox. Preencha `TIPO` com `EPI`, o nome usado na matriz em `ITEM PADRONIZADO`, `TAMANHO` (use `NA` quando não se aplica), e marque `ATIVO?` como `SIM`. A descrição e o código do produto devem corresponder ao catálogo do almox.
6. Implantar como **Aplicativo da Web**, executando como a pessoa responsável pela implantação e restringindo o acesso ao domínio institucional.

## Cruzamento dos códigos

Materiais selecionados pela descrição exata do catálogo recebem o código correspondente automaticamente. Para EPI, o formulário consulta `MAPEAMENTO ALMOX` por tipo, nome padronizado e tamanho. Se não houver um único cruzamento válido, a linha é gravada com `CONFERIR` no código; isso evita associar silenciosamente um item a um código incerto. Complete a aba de mapeamento antes de usar a saída do almox como lançamento definitivo.

## Dados e privacidade

O código-fonte não contém o ID da planilha, credenciais ou dados de colaboradores. O manifesto restringe a aplicação ao domínio do Workspace. Este repositório está público; mantenha nele somente código e documentação sem dados operacionais ou pessoais.

## Estrutura

- `Code.gs` — leitura da matriz e catálogo, validações, movimentações, devoluções e fila de lançamentos no almox.
- `Index.html` — interface responsiva do formulário.
- `appsscript.json` — configuração do projeto e acesso do aplicativo web.


## Controle do almox

Cada item enviado pelo formulário entra como uma linha em `CONTROLE ALMOX`, com data/hora, setor, tipo, item, tamanho, quantidade, código do almox e a caixa `LANÇADO NO ALMOX?`. Marcar a caixa oculta a linha sem apagá-la. Use o menu **Controle Almox > Exibir todos os registros (inclui lançados)** para consultar os itens concluídos. Após atualizar o script, execute `setupEpiModule()` novamente para criar os acionadores necessários e implante uma nova versão do aplicativo web.
