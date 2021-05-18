import pandas as pd
import numpy as np
from functools import reduce
import xlsxwriter

#VERIFICANDO BASE ATIVA
#Coletando as informaçãoes dos Links Ativos  e aplicando o filtro da VIVO (15 e VI)
#  n° de pedido (G) - SRANPED
#  velocidade (AH) - STEVELCON
# operadora - SRACODOPR
print("Iniciando, este processo pode demorar alguns minutos")
dfAtivos  = pd.read_excel("Links Ativos.xlsx")
dfAtivosFiltro1 = dfAtivos[dfAtivos['SRACODOPR'] == '15']
dfAtivosFiltro2 = dfAtivos[dfAtivos['SRACODOPR'] == 'VI']
dfAtivosFiltrado = pd.concat([dfAtivosFiltro1, dfAtivosFiltro2],sort=False)
dfAtivosFiltrado.rename(columns = {'SRANPED':'Nº PEDIDO'},inplace=True)

#Removendo os valores não numericos da lista dos Ativos Filtrados
dfAtivosFiltrado['Numerico']=dfAtivosFiltrado['Nº PEDIDO'].str.isnumeric()
dfAtivosFiltrado.drop(dfAtivosFiltrado.loc[dfAtivosFiltrado['Numerico']==False].index, inplace=True)

#print(dfAtivosFiltrado['Nº PEDIDO'])
dfAtivosFiltrado = dfAtivosFiltrado.dropna(subset=['Nº PEDIDO'])

#Convertendo para string os valores das colunas dos Ativos Filtrados
dfAtivosFiltrado['Nº PEDIDO'] = dfAtivosFiltrado['Nº PEDIDO'].astype('int64')

#Coletando as informações da Base do Itau
dfBase  = pd.read_excel("Base Completa.xlsx")
dfBase['Nº PEDIDO'] = dfBase['Nº PEDIDO'].astype('int64')

#Juntando os dataframes em um unico dataframe
lista_dfs = [dfAtivosFiltrado, dfBase] # Lista vazia. Consolida os resultados
df_consolidado = reduce(lambda left,right: pd.merge(left,right,on=['Nº PEDIDO'],how='outer'), lista_dfs)

#resetando o seu index
df_consolidado.reset_index(inplace = True, drop = True)

#Verifica se o numero de Pedido da Base esta nos Ativos
dfBaseBatimento = dfBase
dfBaseBatimento['Batimento'] = dfBase['Nº PEDIDO'].isin(dfAtivosFiltrado['Nº PEDIDO'])
dfBaseBatimento = pd.DataFrame(dfBaseBatimento, columns=['Nº PEDIDO','Batimento'])

#Convertendo as  Velocidade
df_consolidado.loc[df_consolidado.STEVELCON == "        64 KBPS", "STEVELCON"] = 64
df_consolidado.loc[df_consolidado.STEVELCON == "      2.00 MBPS", "STEVELCON"] = 2048
df_consolidado.loc[df_consolidado.STEVELCON == "         1 GBPS", "STEVELCON"] = 100000
df_consolidado.loc[df_consolidado.STEVELCON == "         1 MBPS", "STEVELCON"] = 1024
df_consolidado.loc[df_consolidado.STEVELCON == "        10 MBPS", "STEVELCON"] = 10000
df_consolidado.loc[df_consolidado.STEVELCON == "        10 MBPS", "STEVELCON"] = 10000
df_consolidado.loc[df_consolidado.STEVELCON == "       100 MBPS", "STEVELCON"] = 102400
df_consolidado.loc[df_consolidado.STEVELCON == "   1024.00 KBPS", "STEVELCON"] = 1024
df_consolidado.loc[df_consolidado.STEVELCON == "       128 KBPS", "STEVELCON"] = 128
df_consolidado.loc[df_consolidado.STEVELCON == "     15.00 MBPS", "STEVELCON"] = 15360
df_consolidado.loc[df_consolidado.STEVELCON == "        16 GBPS", "STEVELCON"] = 160000
df_consolidado.loc[df_consolidado.STEVELCON == "     16.00 KBPS", "STEVELCON"] = 16
df_consolidado.loc[df_consolidado.STEVELCON == "      19.2 KBPS", "STEVELCON"] = 19.2
df_consolidado.loc[df_consolidado.STEVELCON == "         2 MBPS", "STEVELCON"] = 2048
df_consolidado.loc[df_consolidado.STEVELCON == "        20 MBPS", "STEVELCON"] = 20480
df_consolidado.loc[df_consolidado.STEVELCON == "       200 MBPS", "STEVELCON"] = 204800
df_consolidado.loc[df_consolidado.STEVELCON == "   2048.00 MBPS", "STEVELCON"] = 2048
df_consolidado.loc[df_consolidado.STEVELCON == "     25.00 MBPS", "STEVELCON"] = 25600
df_consolidado.loc[df_consolidado.STEVELCON == "       256 KBPS", "STEVELCON"] = 256
df_consolidado.loc[df_consolidado.STEVELCON == "    256.00 KBPS", "STEVELCON"] = 256
df_consolidado.loc[df_consolidado.STEVELCON == "       300 MBPS", "STEVELCON"] = 307200
df_consolidado.loc[df_consolidado.STEVELCON == "     32.00 KBPS", "STEVELCON"] = 32
df_consolidado.loc[df_consolidado.STEVELCON == "      4.00 MBPS", "STEVELCON"] = 4096
df_consolidado.loc[df_consolidado.STEVELCON == "         4 MBPS", "STEVELCON"] = 4096
df_consolidado.loc[df_consolidado.STEVELCON == "         4 MBPS", "STEVELCON"] = 4096
df_consolidado.loc[df_consolidado.STEVELCON == "     50.00 MBPS", "STEVELCON"] = 51200
df_consolidado.loc[df_consolidado.STEVELCON == "       512 KBPS", "STEVELCON"] = 512
df_consolidado.loc[df_consolidado.STEVELCON == "    512.00 KBPS", "STEVELCON"] = 512
df_consolidado.loc[df_consolidado.STEVELCON == "         6 MBPS", "STEVELCON"] = 6144
df_consolidado.loc[df_consolidado.STEVELCON == "      8.00 GBPS", "STEVELCON"] = 80000
df_consolidado.loc[df_consolidado.STEVELCON == "      8.00 MBPS", "STEVELCON"] = 8192
df_consolidado.loc[df_consolidado.STEVELCON == "        10 KBPS", "STEVELCON"] = 10
df_consolidado.loc[df_consolidado.STEVELCON == "      8.00 MBPS", "STEVELCON"] = 8192

# Comparando a Velocidade em STEVELCON e em VELOCIDADE ACESSO PONTA A e Criando uma coluna nova, retornando
# sim quando bate e não quando não bate
df_consolidado['Batimento'] = np.where(df_consolidado['STEVELCON']== df_consolidado['VELOCIDADE ACESSO PONTA A'], 'Sim', 'Não')

#Completando Dados Faltantes
batimentodf1 = df_consolidado.loc[(df_consolidado['Batimento']=='Sim'),['Nº PEDIDO','Batimento']]
batimentodf2 = df_consolidado.loc[(df_consolidado['Batimento']=='Não'),['Nº PEDIDO','Batimento','STEVELCON','VELOCIDADE ACESSO PONTA A']]

#Jutando os dataframes em um unico data frame e renomeando suas colunas.
dfBatimentos = pd.concat([batimentodf1, batimentodf2],sort=False)
dfBatimentos.rename(columns = {'STEVELCON':'Velocidade Base Itau'},inplace=True)
dfBatimentos.rename(columns = {'VELOCIDADE ACESSO PONTA A':'Velocidade Base Operedora'},inplace=True)

#VERIFICANDO BASE CANCELADA
#Coletando as informaçãoes dos Links Ativos  e aplicando o filtro da VIVO (15 e VI)
#Separando por Numero de pedido e data do cancelamento
#  n° de pedido (G) - SRCNPED
#  data cancelmanto  - SRCDTCAN
# operadora - SRCCODOPR

dfCancelado  = pd.read_excel("Links Cancelados.xlsx")
dfCancelado = pd.DataFrame(dfCancelado, columns=['SRCNPED','SRCDTCAN','SRCCODOPR'])
dfCanceladoFiltro1 = dfCancelado[dfCancelado['SRCCODOPR'] == '15']
dfCanceladoFiltro2 = dfCancelado[dfCancelado['SRCCODOPR'] == 'VI']
dfCanceladoFiltrado = pd.concat([dfCanceladoFiltro1, dfCanceladoFiltro2],sort=False)
dfCanceladoFiltrado.rename(columns = {'SRCNPED':'Nº PEDIDO'},inplace=True)

# Removendo os valores não numericos da tabela dos pedidos cancelados
dfCanceladoFiltrado['Numerico']=dfCanceladoFiltrado['Nº PEDIDO'].str.isnumeric()
dfCanceladoFiltrado.drop(dfCanceladoFiltrado.loc[dfCanceladoFiltrado['Numerico']==False].index, inplace=True)
dfCanceladoFiltrado = dfCanceladoFiltrado.dropna(subset=['Nº PEDIDO'])

#Conventerdno para Int os valores do numero de pedido
dfCanceladoFiltrado['Nº PEDIDO'] = dfCanceladoFiltrado['Nº PEDIDO'].astype('int64')

#Verificando se o pedido cancelado esta na base do Banco, e se estiver ira retornar o numero do pedido
# e a data do cancelamento.
dfCanceladoFiltrado['Batimento'] = dfCanceladoFiltrado['Nº PEDIDO'].isin(dfBase['Nº PEDIDO'])
dfCanceladoFinal = dfCanceladoFiltrado.loc[(dfCanceladoFiltrado['Batimento']== True),['Nº PEDIDO','SRCDTCAN']]
dfCanceladoFinal.rename(columns = {'SRCDTCAN':'Data Cancelamento'},inplace=True)

#Verificando se o pedido da base esta no Links ativos ou no Links Cancelados
dfBaseBatimento['Batimento2'] = dfBaseBatimento['Nº PEDIDO'].isin(dfCanceladoFiltrado['Nº PEDIDO'])
dfBaseBatimento['Situação'] = np.where(dfBaseBatimento['Batimento']== dfBaseBatimento['Batimento2'], 'Não Cadastrado', 'Cadastrado')
dfBaseBatimento = dfBaseBatimento.drop(dfBaseBatimento[(dfBaseBatimento['Batimento'] == True) & (dfBaseBatimento['Batimento2'] == True)].index)
dfBaseBatimento.drop(dfBaseBatimento.loc[dfBaseBatimento['Situação']=='Cadastrado'].index, inplace=True)
dfBaseBatimento.reset_index(inplace = True, drop = True)
dfBaseBatimento = pd.DataFrame(dfBaseBatimento, columns=['Nº PEDIDO','Situação'])

#Removendo o pedido não cadastrado que esta na base dos ativos
dfBatimentos['Cadastrado'] = dfBatimentos['Nº PEDIDO'].isin(dfBaseBatimento['Nº PEDIDO'])
dfBatimentos.drop(dfBatimentos.loc[dfBatimentos['Cadastrado']==True].index, inplace=True)
del dfBatimentos['Cadastrado']

# Salvando no Excel e formatando suas colunas
writer = pd.ExcelWriter("Batimento.xlsx", engine="xlsxwriter")
dfBatimentos.to_excel(writer, sheet_name='Ativo',index=False)
dfCanceladoFinal.to_excel(writer, sheet_name='Cancelado',index=False)
dfBaseBatimento.to_excel(writer, sheet_name='Não Cadastrado',index=False)
workbook  = writer.book
worksheet = writer.sheets['Ativo']
worksheet.set_column(0, 50, 25)
worksheet2 = writer.sheets['Cancelado']
worksheet2.set_column(0, 50, 20)
worksheet3 = writer.sheets['Não Cadastrado']
worksheet3.set_column(0, 50, 20)
format = workbook.add_format({'align':'center'})
worksheet.set_column('A:Z',25,format)
worksheet2.set_column('A:Z',25,format)
worksheet3.set_column('A:Z',25,format)

writer.save()
print("Processo Finalizado")
