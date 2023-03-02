# Organizador de arquivos XML NF-E

Este script lê arquivos XML de NF-E 4.00 em uma pasta especificada, busca pelo elemento "cAut" em cada arquivo e verifica se o valor desse elemento está presente em uma determinada coluna de uma planilha do Excel. Em seguida, move o arquivo para uma das três pastas, dependendo se o valor foi encontrado na planilha, se a tag "cAut" não foi encontrada ou se o arquivo não é um arquivo XML.

## Por que usar este script?

Este script é útil para quem precisa organizar arquivos NF-E em XML em uma pasta com base em um valor presente na tag "cAut". Por exemplo, se você recebe muitos arquivos XML e precisa verificar se estão em uma planilha, pode usar este script para ler o valor de uma tag de cada arquivo XML e mover o arquivo para uma pasta correspondente, com base em uma planilha do Excel que contém uma coluna com todos os valores da tag que você porcura.

## Setup


1. Instale as bibliotecas necessárias do Python, que são os, 
    - `xml.etree.ElementTree` 
    - `shutil`
    - `openpyxl`

2. Coloque todos os arquivos XML que você deseja organizar em uma pasta especificada e atualize a variável PASTA_ORIGEM no script para refletir o caminho para esta pasta.

3. Atualize as variáveis PASTA_DESTINO_ENCONTRADO, PASTA_DESTINO_NAO_ENCONTRADO e PASTA_DESTINO_SEM_TAG para refletir os caminhos para as pastas para onde você deseja que o script mova os arquivos.

4. Abra sua planilha do Excel e atualize a variável planilha no script para refletir o nome da planilha na qual você deseja pesquisar pelo valor "cAut".

## Como usar

1. Coloque todos os arquivos XML que deseja organizar na pasta `inputs`.
2. Abra o arquivo Excel que contém a lista de valores a serem comparados com a tag "cAut".
3. Na planilha do Excel, insira os valores de referência na coluna apropriada (a coluna pode ser alterada no script).
4. Execute o script. Os arquivos XML serão movidos para as pastas `Encontrado`, `NaoEncontrado` ou `semTag`, dependendo do resultado da comparação.
5. Verifique as pastas `Encontrado`, `NaoEncontrado` e `semTag` para verificar o resultado.

## Autor

Este script foi criado por [Thiago Gentil](https://github.com/Tgentil).

