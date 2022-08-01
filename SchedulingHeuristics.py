#		UNIVERSIDADE FEDERAL DE SANTA CATARINA
#			   Engenharia Mecatronica
#						PO3
#
#	  Heuristicas para Sequenciamento de Tarefas
#
#
#				 Guilherme Turatto
#
#			 Joinville, julho de 2022 

from ast import Return
from copy import deepcopy
from IPython.display import display
import PySimpleGUI as sg
import pandas as pd


class Interface:
	def __init__(self):

		layout = [
			[sg.Text('Nome do Arquivo:'), sg.Input(default_text='DadosTarefas.xlsx', size=(38,5), key='FileName')],
			[sg.Text(' ',size=(35,0))],
			[sg.Text('Heurísticas a serem calculadas:',size=(30,0))],
			[sg.Checkbox('SPT',key='SPT'),sg.Checkbox('LPT',key='LPT'),sg.Checkbox('WSPT',key='WSPT'),sg.Checkbox('WLPT',key='WLPT'),sg.Checkbox('EDD',key='EDD'),sg.Checkbox('MST',key='MST')],
			[sg.Text(' ',size=(35,0))],
			[sg.Text('Se já existir uma folha com a mesma heurística:')],
			[sg.Radio('Sobrescrever','sobrescrever', key='overWrite',default=True), sg.Radio('Criar nova folha','sobrescrever', key='NOToverWrite')],
			[sg.Text(' ',size=(35,0))],
			[sg.Button('Iniciar', expand_x=True)],
			[sg.Text('- - -', expand_x=True, justification='center', key='msg_usr')],
		]

		self.janela = sg.Window("Heurísticas para Sequênciamento de Tarefas", layout=layout)

	def Read(self):
		global file_name
		global runSPT, runLPT, runWSPT, runWLPT, runEDD, runMST, OverWriteSheet
		
		try:
			self.button, self.values = self.janela.Read()

			file_name 	= self.values['FileName']
			runSPT 		= self.values['SPT']
			runLPT		= self.values['LPT']
			runWSPT		= self.values['WSPT']
			runWLPT		= self.values['WLPT']
			runEDD		= self.values['EDD']
			runMST		= self.values['MST']

			OverWriteSheet = self.values['overWrite']
		
		except:
			raise Exception("Window Closed")

	def AtualizaMSG(self,msg):
		self.janela['msg_usr'].update(msg)


def Dataframe_to_Excel(file_name, sheet_name, data):
	#Cria manipulador do arquivo excel
	xlwriter = 0
	if OverWriteSheet:
		xlwriter = pd.ExcelWriter(file_name, mode='a', if_sheet_exists='overlay')
	else:
		xlwriter = pd.ExcelWriter(file_name, mode='a', if_sheet_exists='new')

	#Escreve os dados do dataframe no excel
	data.to_excel(excel_writer=xlwriter, float_format="%.2f", sheet_name=sheet_name, index=False, startrow=0, startcol=0)
	
	#Destroi o manipulador do arquivo excel
	xlwriter.close()

def costAccounting(sortedData, Heuristic):
	#Cria o dataframe de resposta que sera salvo no arquivo excel
	col = ['Inicio', 'Fim', 'Atraso(+) ou Adiantamento(-)', 'Ocorre Adiantamento', 'Ocorre Atraso', 'Custo do Adiantamento', 'Custo do Atraso', 'Sequência ' + Heuristic]
	outData = pd.DataFrame(columns=col)

	#inicializa as variaveis para contagem de tempo e custo
	t_inicio	= 0 #Tempo de inicio de cada tarefa
	t_fim 		= 0 #Tempo de finalizacao de cada tarefa
	c_atTotal 	= 0	#Custo total de atraso
	c_adTotal 	= 0 #Custo total de adiantamento
	t_atTotal 	= 0 #Tempo total de atraso
	t_adTotal 	= 0	#Tempo total de adiantamento

	for index, row in sortedData.iterrows():
		t_inicio = t_fim													#Atualiza o tempo inicial						
		t_fim = t_inicio + row['Tempo de Processamento']					#Calcula o tempo de finalizacao
		S = t_fim - row['Data de Entrega']									#Calcula o atraso ou adiantamento
		

		if(S <= 0):															#Se S for negativo, houve adiantamento
			adiantamento	= 1												
			atraso			= 0												
			c_adiantamento	= abs(S) * row['Penalidade por Adiantamento']	#Calcula custo por adiantamento
			c_atraso		= 0												#Nao ha custo por atraso
			t_adTotal		+= abs(S)										#Acumula tempo de adiantamento

		else:																#Se S for positivo, houve atraso			
			adiantamento 	= 0
			atraso			= 1
			c_adiantamento	= 0												#Nao ha custo por adiantamento
			c_atraso		= abs(S) * row['Penalidade por Atraso']			#Calcula custo por atraso
			t_atTotal		+= abs(S)										#Acumula tempo de atraso

		#Acumula os custos de atraso e adiantamento
		c_atTotal += c_atraso
		c_adTotal += c_adiantamento	

		#Adiciona os dados calculados ao dataframe
		outData.loc[index] = [t_inicio, t_fim, S, adiantamento, atraso, c_adiantamento, c_atraso, row['Produto']]

	#Adiciona uma linha em branco e os custos totais na ultima linha do dataframe
	outData.loc[outData.shape[0]+1] = [' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ']
	outData.loc[outData.shape[0]+2] = [' ', ' ', 'TOTAL:', t_adTotal, t_atTotal, c_adTotal, c_atTotal, (c_adTotal + c_atTotal)]
	
	return outData


def SPT(inData):
	#Ordena os dados de acordo com a heuristica
	#-->Tempo de processamento em ordem crescente
	sortedData = inData.sort_values('Tempo de Processamento')

	#Calcula os custos atrelados a esta ordem de tarefas 
	outData = costAccounting(sortedData, 'SPT')

	#Chama a funcao para escrever o dataframe dos custos no excel
	Dataframe_to_Excel(file_name, 'SPT', outData)
	
	#Exibe o dataframe na tela
	#display(outData)

def LPT(inData):
	#Ordena os dados de acordo com a heuristica
	#-->Tempo de processamento em ordem decrescente
	sortedData = inData.sort_values('Tempo de Processamento', ascending=False)  

	#Calcula os custos atrelados a esta ordem de tarefas 
	outData = costAccounting(sortedData, 'LPT')

	#Chama a funcao para escrever o dataframe dos custos no excel
	Dataframe_to_Excel(file_name, 'LPT', outData)
	
	#Exibe o dataframe na tela
	#display(outData)

def WSPT(inData):
	#Ordena os dados de acordo com a heuristica
	sortedData = deepcopy(inData)

	#Adiciona uma coluna ao dataframe calculando a metrica desejada para todas as linhas
	sortedData['WSPT'] = 'NaN'
	for index, row in sortedData.iterrows():
		sortedData.loc[index, 'WSPT'] = row['Penalidade por Atraso'] / row['Tempo de Processamento']


	#--> Ordena as razoes de forma decrescente para que: (w[1] / p[1]) ≥ (w[2] / p[2]) ≥ ... ≥ (w[n] / p[n]);
	sortedData = sortedData.sort_values('WSPT', ascending=False)
	#display(sortedData) 


	#Calcula os custos atrelados a esta ordem de tarefas 
	outData = costAccounting(sortedData, 'WSPT')

	#Chama a funcao para escrever o dataframe dos custos no excel
	Dataframe_to_Excel(file_name, 'WSPT', outData)
	
	#Exibe o dataframe na tela
	#display(outData)

def WLPT(inData):
	#Ordena os dados de acordo com a heuristica
	sortedData = deepcopy(inData)

	#Adiciona uma coluna ao dataframe calculando a metrica desejada para todas as linhas
	sortedData['WLPT'] = 'NaN'
	for index, row in sortedData.iterrows():
		sortedData.loc[index, 'WLPT'] = row['Penalidade por Adiantamento'] / row['Tempo de Processamento']


	#--> Ordena as razoes de forma crescente para que: (h[1] / p[1]) ≤ (h[2] / p[2]) ≤ ... ≤ (h[n] / p[n])
	sortedData = sortedData.sort_values('WLPT')
	#display(sortedData) 


	#Calcula os custos atrelados a esta ordem de tarefas 
	outData = costAccounting(sortedData, 'WLPT')

	#Chama a funcao para escrever o dataframe dos custos no excel
	Dataframe_to_Excel(file_name, 'WLPT', outData)
	
	#Exibe o dataframe na tela
	#display(outData)

def MST(inData):
	#Ordena os dados de acordo com a heuristica
	sortedData = deepcopy(inData)

	#Adiciona uma coluna ao dataframe calculando a metrica desejada para todas as linhas
	sortedData['MST'] = 'NaN'
	for index, row in sortedData.iterrows():
		sortedData.loc[index, 'MST'] = row['Data de Entrega'] - row['Tempo de Processamento']


	#--> Ordena as razoes de forma crescente para que: (d[1] – p[1]) ≤ ( d[2] – p[2]) ≤ ... ≤ (d[n] – p[n])
	sortedData = sortedData.sort_values('MST')
	#display(sortedData) 


	#Calcula os custos atrelados a esta ordem de tarefas 
	outData = costAccounting(sortedData, 'MST')

	#Chama a funcao para escrever o dataframe dos custos no excel
	Dataframe_to_Excel(file_name, 'MST', outData)
	
	#Exibe o dataframe na tela
	#display(outData)

def EDD(inData):
	#Ordena os dados de acordo com a heuristica
	#-->Data de Entrega em ordem crescente
	sortedData = inData.sort_values('Data de Entrega')  

	#Calcula os custos atrelados a esta ordem de tarefas 
	outData = costAccounting(sortedData, 'EDD')

	#Chama a funcao para escrever o dataframe dos custos no excel
	Dataframe_to_Excel(file_name, 'EDD', outData)
	
	#Exibe o dataframe na tela
	#display(outData)


def main():
	
	tela = Interface()

	while(True):
		#Tenta ler os comandos da janela de interface
		try:
			tela.Read()
			tela.AtualizaMSG('- - -')
		
		#Se a janela for fechada encerra o programa 
		except:
			break
		
		#Tenta ler a planilha
		try:
			data = pd.read_excel(file_name, sheet_name=0)
			#display(data)

			if(not(runSPT or runLPT or runWSPT or runWLPT or runEDD or runMST)):
				tela.AtualizaMSG('Nenhuma Heurística Selecionada')

			else:
				if runSPT:
					SPT(data)
				if runLPT:
					LPT(data)
				if runWSPT:
					WSPT(data)
				if runWLPT:
					WLPT(data)
				if runEDD:
					EDD(data)
				if runMST:
					MST(data)
					
				tela.AtualizaMSG('Cálculos Finalizados')
		
		#Mensagem de erro caso a planilha esteja aberta
		except:
			tela.AtualizaMSG('Erro ao ler o arquivo de dados')


if __name__ == "__main__":
    main()

