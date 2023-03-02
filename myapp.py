'''

O script realiza a leitura de arquivos XML 
de uma pasta especificada, procura pelo elemento "cAut" 
em cada arquivo e verifica se o valor desse elemento está presente 
em uma determinada coluna de uma planilha do Excel. Em seguida, 
move o arquivo para uma das três pastas, dependendo se o valor foi 
encontrado na planilha, se a tag "cAut" não foi encontrada ou se o 
arquivo não é um arquivo XML.

'''

# Importando bibliotecas necessárias
import os
import xml.etree.ElementTree as ET
import shutil
import openpyxl

# Definindo variáveis para pastas de origem e destino
PASTA_ORIGEM = r"C:\Users\pasta\exemplo\inputs"
PASTA_DESTINO_ENCONTRADO = r"C:\Users\pasta\exemplo\Encontrado"
PASTA_DESTINO_NAO_ENCONTRADO = r"C:\Users\pasta\exemplo\NaoEncontrado"
PASTA_DESTINO_SEM_TAG = r"C:\Users\pasta\exemplo\semTag"

# Carregando planilha Excel
arquivo_excel = openpyxl.load_workbook(
    r"C:\Users\path\exemplo\relatorio\nomeDoRelatório.xlsx")
planilha = arquivo_excel['nome_da_planilha']

# Iterando por todos os arquivos na pasta de origem
for arquivo in os.listdir(PASTA_ORIGEM):
    if arquivo.endswith('.xml'):
        caminho_arquivo = os.path.join(PASTA_ORIGEM, arquivo)

        try:
            # Lendo o valor do elemento cAut do arquivo XML
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            cAut = root.find(
                './/{http://www.portalfiscal.inf.br/nfe}cAut').text

        except AttributeError:
            # elemento cAut não encontrado no arquivo XML
            shutil.move(caminho_arquivo, os.path.join(
                PASTA_DESTINO_SEM_TAG, arquivo))
            continue

        # procure o valor na coluna (altere para a sua coluna)
        VALOR_ENCONTRADO = False
        for linha in planilha.iter_rows(min_row=2, min_col=2, max_col=2):
            if linha[0].value == cAut:
                VALOR_ENCONTRADO = True
                break

        # Movendo o arquivo para a pasta de destino correspondente
        if VALOR_ENCONTRADO:
            DESTINO = PASTA_DESTINO_ENCONTRADO
        else:
            DESTINO = PASTA_DESTINO_NAO_ENCONTRADO

        shutil.move(caminho_arquivo, os.path.join(DESTINO, arquivo))    
# mensagem de conclusão da tarefa
print("Processamento concluído.")
