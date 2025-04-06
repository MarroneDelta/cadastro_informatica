import os
from time import sleep
import pandas as pd

class Computador:
    def __init__(self, marca, memoria, placa_de_video, sistema_operacional, tipo_processador, valor):
        self.marca = marca.upper()
        self.memoria = memoria.upper()
        self.placa_de_video = placa_de_video.upper()
        self.sistema_operacional = sistema_operacional.upper()
        self.tipo_processador = tipo_processador.upper()
        self.valor = valor
        print(f'A marca do Notebook é: {marca} \nQuantidade de memória é: {memoria}GB \nA placa de vídeo é: {placa_de_video} \nSistema Operacional é: {sistema_operacional} \nTipo de Processador é: {tipo_processador} \nPreço em promoção: {valor}')
        
        # Dados para o DataFrame
        dados = { 
            'Marca': [self.marca], 
            'Memória': [self.memoria], 
            'Placa de Vídeo': [self.placa_de_video], 
            'Sistema Operacional': [self.sistema_operacional], 
            'Tipo de Processador': [self.tipo_processador], 
            'Valor': [self.valor]
        }
        df = pd.DataFrame(dados)
        
        # Caminho do arquivo .xlsx
        caminho_arquivo = 'C:/Área de Trabalho/Nova pasta/computadores.xlsx'

        try:
            # Verificar se o arquivo já existe
            if os.path.exists(caminho_arquivo):
                df_existente = pd.read_excel(caminho_arquivo)
                df_final = pd.concat([df_existente, df], ignore_index=True)
            else:
                df_final = df

            # Gravar o DataFrame no arquivo .xlsx
            df_final.to_excel(caminho_arquivo, index=False)
            print(f'Dados salvos em {caminho_arquivo}')
        except Exception as e:
            print(f'Ocorreu um erro ao salvar os dados: {e}')

while True:
    marca = input('Digite a marca do seu Notebook (ou digite "sair" para terminar): ')
    if marca.lower() == 'sair':
        print('Encerrando o programa, aguarde...')
        sleep(3)
        print('Programa finalizado')
        break
    memoria = input('Digite a quantidade de memória: ')
    placa_de_video = input('Digite a placa de vídeo: ')
    sistema_operacional = input('Digite o sistema operacional: ')
    tipo_processador = input('Digite o modelo do processador: ')
    valor = input('Digite o preço: ')
    
    computador = Computador(marca, memoria, placa_de_video, sistema_operacional, tipo_processador, valor)
