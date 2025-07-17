# Auto List ADR

## Introdução

Este é um programa que automatiza o preenchimento de planilhas Google com dados de ADR (Additional Dialogue Recording) para um determinado projeto, assim como a geração de textos em rtf para a gravação dos diálogos adicionais.
A aplicação lê markers exportados em formato txt de sessões do Pro Tools, que sigam determinada estrutra, e preenche a planilha especificada pelo usuário com os dados do documento, gerando os textos para cada personagem no fim.

Antes de usar, é necessário fazer o [quickstart](https://developers.google.com/workspace/sheets/api/quickstart/python?hl=pt-br) do Google para usar a API do Google Sheets e salvar o JSON com id de cliente na mesma pasta que o programa. 

## Estrutura de Markers

adr personagem “Texto a ser falado.” Motivo # OBS # TC OUT # idioma

Não é necessário colocar letras maiúsculas em ADR e personagem, porque elas
ficarão todas maiúsculas automaticamente. Porém texto, motivo e observação
precisam estar do jeito que o usuário quer que apareça na planilha (e no textos do
atores, no caso do texto).
Não é necessário preencher todos os campos indicados no Marker. Os únicos
campos obrigatórios são personagem e texto, todos os outros são opcionais.
Porém, caso o usuário queira pular um dos campos e ir direto para outro, é
necessário colocar no Marker a marcação de fim do(s) campo(s) anterior(es),
mesmo que não haja nada nesse(s) campo(s). Por exemplo:

adr personagem “Texto a ser falado.” # OBS

O marker acima irá gerar uma entrada do personagem PERSONAGEM com o
texto “Texto a ser falado.” e a observação OBS, sem nenhum motivo e nenhum
TC OUT (o idioma ficará o padrão, “Português”). Note como foi necessário usar o
“#” para pular o campo de motivo. Caso ele não fosse utilizado, a entrada teria
OBS como motivo e nenhuma observação.

adr personagem “Texto a ser falado.” # # 01:23:04:02 # Idioma

No caso acima, a entrada de PERSONAGEM vai ter o mesmo texto, nenhum
motivo e nenhuma observação, TC OUT 01:23:04:02 e idioma Idioma. Note que
foi necessário colocar no Marker 2 marcações “#” antes de informar o TC OUT,
para o programa pular os campos de motivo e observação.
Sempre que o idioma for omitido ele será computado como o padrão, “Português”.
Para informar o idioma “Español” basta colocar um “e” na posição do idioma.
Caso queira-se marcar qualquer outro idioma, é necessário escrever por
completo.

## Como Exportar os Markers no Pro Tools

File -> Export -> Session Info as Text -> Deixar selecionado apenas Include Markers, e mudar File Format para UTF-8 ‘TEXT’
