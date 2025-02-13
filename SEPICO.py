import os.path
import os
from os import chdir, listdir
from os.path import isdir

import pandas as pd
import openpyxl

import shutil
from datetime import datetime


# Verfica se o que foi digitado é uma letra
def is_alpha(word):
    return word.isalpha()

#Define um contador que será utilizado posteriormente no cógido
Contador= 0


#Define o nome do arquivo.
ARQUIVO_LOG = 'log_SEPICO.xlsx'


#Busca o diretório atual
dir_atual = os.getcwd()


#Caminho do patch definido diretamente
caminho_pasta = ("PASTA")


#Altera o diretório de trabalho atual para path definido.
chdir(caminho_pasta)

#Titulos do programa
print(f'\tLIMPADOR PASTA DE TRANFERÊNCIA\n\r')
print("Listando diretórios Encontrados:\n\r")



#lista todos os diretórios encontrados no path
buscar_caminho = [caminho_pasta+"/"+fd for fd in os.listdir(caminho_pasta)]

# Faz varredura dos dados dentro da variável buscar_caminho e armazena dentro de buscar
for buscar in buscar_caminho:

  #Pega as informações sobre cada diretório encontrado  
  informacao_diretorio = os.stat(buscar)

  #Extrai a data da ultima modificação do diretório
  data_modificacao = datetime.fromtimestamp(informacao_diretorio.st_mtime).strftime('%Y-%m-%d')

  #Exibe o nome do diretório e a data da ultima modificaçãop
  print(f'Diretório: {buscar} ------- Data da ultima modificação {data_modificacao}')

  #incrementa o contador
  Contador +=1


#Busca a data atual do sistema 
data_atual = datetime.today().date()

#Faz a formatação da Data
data_formatada = data_atual.strftime('%Y-%m-%d')

#Exibe a data atual obtida do sistema e formatada no modelo Ano-mes-dia Hora:minuto:segundo
print(f'\t\t\t\t\n\r Data Atual: {data_formatada}\n\r')

#Faz a verificação de há algum diretório na listagem
if Contador > 0:

    #Verifica se a data de modificação do arquivo é diferente da atual para poder criar a planilha de log.
    if data_formatada != data_modificacao:
  
        #Cria a planilha de Log
        criar_log_xls = openpyxl.Workbook()
        sheet = criar_log_xls.active
        criar_log_xls.save(dir_atual+ARQUIVO_LOG)
        

    while True:

        #Solicita uma opção do usuário
        verificacao = input('Você deseja Limpar as pastas? [S]im  [N]ão: ')

        #Faz a validação em uma função se o que foi digitado é uma letra      
        validacao_string = is_alpha(verificacao)

        #Verifica o valor retornado da função:
        if validacao_string == False:
            print('\n\rOops! Parece que o que você digitou não é uma opção válida! Tente novamente')
            continue

        else:
            #Verifica se o que foi digitado foi S
            if verificacao.upper() == ('S'):

                #faz a busca dos diretórios a serem excluidos
                for caminho in listdir(caminho_pasta):

                    #verifica se o objeto analizado é um diretório
                    if isdir(caminho):

                        #verifica se há algo em caminho
                        if len(caminho) >0:

                            #busca as informações do diretório
                            informacao_diretorio = os.stat(caminho)

                            #busca a data da ultima modificação
                            data_modificacao = datetime.fromtimestamp(informacao_diretorio.st_mtime).strftime('%Y-%m-%d')

                            #Verifica se a data da ultima modificação é diferente da data atual.
                            if data_modificacao != data_formatada:

                                #Busca o arquivo no diretório atual e faz a leirua
                                df = pd.read_excel(dir_atual+ARQUIVO_LOG)

                                #Faz a criação de um dicionário com as informações das pastas que serão excluidas
                                log_dados = {'Diretório Principal':[caminho_pasta],
                                             'Diretório Excluido': [caminho],
                                             'Data Ultima Modificação': [data_modificacao],
                                             'Data Exclusão': [data_formatada] }

                                #cria um data frame com o dicionário log_dados
                                df_log_dados = pd.DataFrame(log_dados)

                                #concactena o noso data frame com o data frame que está contido no arquivo.xls lido
                                df_concatenado = pd.concat([df, df_log_dados], ignore_index=True)

                                #Faz a inserção dos dados no data frame e posteriormente a criação de um arquico log_SEPICO.xls no mesmo diretório do programa
                                df_concatenado.to_excel(dir_atual+'\\'+ARQUIVO_LOG, index=False)

                                try:
                                    print('Excluido: '+caminho_pasta + '\\' + caminho +' '+data_modificacao)
                                    print()

                                    #faz a exclusão apenas dos diretórios
                                    shutil.rmtree(caminho, ignore_errors=False, onerror=None)


                                except:
                                    print('Não foi possível excluir o diretório: '+caminho_pasta + '\\' + caminho)
                                    print("Tente novamente depois!")
                            
                            else:
                                continue
                        else:
                            print("Oops! Parece que não há nenhum diretório")
                            break
            elif verificacao.upper() == ('N'):
                
                print("Entendido! Irei parar a execução para que não ocorra problemas!")
                break
            else:
                print("Ooops! Não entendi! Poderia digitar novamente?")
                continue
            
        os.remove(dir_atual+ARQUIVO_LOG)
        break

else:
    print('Nenhum diretório encontrado!')

print('Programa Finalizado!')
    

