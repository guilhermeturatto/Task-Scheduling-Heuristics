#   UNIVERSIDADE FEDERAL DE SANTA CATARINA
#          Mechatronics Engineering              
#	       Operational Research III
#
#	      Task Scheduing Heuristics
#
#
#			  Guilherme Turatto
#
#	       Joinville, julho de 2022 

from copy import deepcopy
import PySimpleGUI as sg
import pandas as pd
#from IPython.display import display


class Interface:
	def __init__(self):
		
		layout = [
		[sg.Text('File Name:'), sg.Input(default_text='TaskData.xlsx', size=(38,5), key='FileName')],
		[sg.Text(' ',size=(35,0))],
		[sg.Text('Heuristics to be calculated:',size=(30,0))],
		[sg.Checkbox('SPT',key='SPT'),sg.Checkbox('LPT',key='LPT'),sg.Checkbox('WSPT',key='WSPT'),sg.Checkbox('WLPT',key='WLPT'),sg.Checkbox('EDD',key='EDD'),sg.Checkbox('MST',key='MST')],
		[sg.Text(' ',size=(35,0))],
		[sg.Text('If a sheet already exists with the same heuristic:')],
		[sg.Radio('Overwrite','overwrite', key='overWrite',default=True), sg.Radio('Create new sheet','overwrite', key='NOToverWrite')],
		[sg.Text(' ',size=(35,0))],
		[sg.Button('Start', expand_x=True)],
		[sg.Text('- - -', expand_x=True, justification='center', key='msg_usr')],
		]

		self.window = sg.Window("Task Scheduling Heuristics", layout=layout)

	def Read(self):
		global file_name
		global runSPT, runLPT, runWSPT, runWLPT, runEDD, runMST, OverWriteSheet
		
		try:
			self.button, self.values = self.window.Read()

			file_name   = self.values['FileName']
			runSPT      = self.values['SPT']
			runLPT      = self.values['LPT']
			runWSPT     = self.values['WSPT']
			runWLPT     = self.values['WLPT']
			runEDD      = self.values['EDD']
			runMST      = self.values['MST']

			OverWriteSheet = self.values['overWrite']
		
		except:
			raise Exception("Window Closed")

	def UpdateMSG(self,msg):
		self.window['msg_usr'].update(msg)

def Dataframe_to_Excel(file_name, sheet_name, data):
    #Create excel file handler
    xlwriter = 0
    if OverWriteSheet:
        xlwriter = pd.ExcelWriter(file_name, mode='a', if_sheet_exists='overlay')
    else:
        xlwriter = pd.ExcelWriter(file_name, mode='a', if_sheet_exists='new')

    #Write dataframe to excel file
    data.to_excel(excel_writer=xlwriter, float_format="%.2f", sheet_name=sheet_name, index=False, startrow=0, startcol=0)
    #xlwriter.save()
    #Close the excel file handler
    xlwriter.close()

def costAccounting(sortedData, Heuristic):
    # Create the output dataframe that will be saved in the excel file
    col = ['Start', 'End', 'Delay(+) or Advancement(-)', 'Advancement Occurs', 'Delay Occurs', 'Advancement Cost', 'Delay Cost', Heuristic + ' Sequence']
    outData = pd.DataFrame(columns=col)

    # Initialize variables for time and cost count
    t_start    = 0 #Start time of each task
    t_end      = 0 #End time of each task
    c_dTotal   = 0 #Total delay cost
    c_aTotal   = 0 #Total advancement cost
    t_dTotal   = 0 #Total delay time
    t_aTotal   = 0 #Total advancement time

    for index, row in sortedData.iterrows():
        t_start = t_end                                                  # Update start time
        t_end = t_start + row['Processing Time']                         # Calculate end time
        S = t_end - row['Deadline']                                      # Calculate delay or advancement

        if(S == 0):                                                      # Just in time
            advancement    = 0
            delay          = 0
            c_advancement  = 0                                           # No advancement cost
            c_delay        = 0                                           # No delay cost

        elif(S < 0):                                                     # If S is negative, there was advancement
            advancement    = 1
            delay          = 0
            c_advancement  = abs(S) * row['Advancement Penalty']         # Calculate advancement cost
            c_delay        = 0                                           # No delay cost
            t_aTotal       += abs(S)                                     # Accumulate advancement time

        else:                                                            # If S is positive, there was delay
            advancement    = 0
            delay          = 1
            c_advancement  = 0                                           # No advancement cost
            c_delay        = abs(S) * row['Delay Penalty']               # Calculate delay cost
            t_dTotal       += abs(S)                                     # Accumulate delay time

        # Accumulate delay and advancement costs
        c_dTotal += c_delay
        c_aTotal += c_advancement

        # Add calculated data to dataframe
        outData.loc[index] = [t_start, t_end, S, advancement, delay, c_advancement, c_delay, row['Task']]

    # Add a blank line and total costs in the last line of the dataframe
    outData.loc[outData.shape[0]+1] = [' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ']
    outData.loc[outData.shape[0]+2] = [' ', ' ', ' ', 'Total Advancement Time',\
                                                      'Total Delay Time'      ,\
                                                      'Total Advancement Cost',\
                                                      'Total Delay Cost'      ,\
                                                      'Total Costs'
                                    ]
    outData.loc[outData.shape[0]+3] = [' ', ' ', ' ', t_aTotal  , t_dTotal  , c_aTotal  , c_dTotal  , (c_aTotal + c_dTotal)]

    return outData

def SPT(inData):
    #Sort data according to the heuristic
    #-->Processing time in ascending order
    sortedData = inData.sort_values('Processing Time')

    #Calculate costs associated with this task order
    outData = costAccounting(sortedData, 'SPT')

    #Call function to write the cost dataframe to excel
    Dataframe_to_Excel(file_name, 'SPT', outData)

    #Display the dataframe on the screen
    #display(outData)

def LPT(inData):
    #Sort data according to the heuristic
    #-->Processing time in descending order
    sortedData = inData.sort_values('Processing Time', ascending=False)  

    #Calculate costs associated with this task order
    outData = costAccounting(sortedData, 'LPT')

    #Call function to write the cost dataframe to excel
    Dataframe_to_Excel(file_name, 'LPT', outData)

    #Display the dataframe on the screen
    #display(outData)

def WSPT(inData):
    #Sort data according to the heuristic
    sortedData = deepcopy(inData)

    #Add a column to the dataframe calculating the desired metric for all rows
    sortedData['WSPT'] = 'NaN'
    for index, row in sortedData.iterrows():
        sortedData.loc[index, 'WSPT'] = row['Delay Penalty'] / row['Processing Time']


    #-->Sort the ratios in descending order so that: (w[1] / p[1]) ≥ (w[2] / p[2]) ≥ ... ≥ (w[n] / p[n]);
    sortedData = sortedData.sort_values('WSPT', ascending=False)
    #display(sortedData) 


    #Calculate costs associated with this task order
    outData = costAccounting(sortedData, 'WSPT')

    #Call function to write the cost dataframe to excel
    Dataframe_to_Excel(file_name, 'WSPT', outData)

    #Display the dataframe on the screen
    #display(outData)

def WLPT(inData):
    #Sort data according to heuristic
    sortedData = deepcopy(inData)

    #Add a column to the dataframe calculating the desired metric for all rows
    sortedData['WLPT'] = 'NaN'
    for index, row in sortedData.iterrows():
        sortedData.loc[index, 'WLPT'] = row['Advancement Penalty'] / row['Processing Time']

    #--> Sort the ratios in ascending order so that: (h[1] / p[1]) ≤ (h[2] / p[2]) ≤ ... ≤ (h[n] / p[n])
    sortedData = sortedData.sort_values('WLPT')
    #display(sortedData)

    #Calculates the costs associated with this order of tasks
    outData = costAccounting(sortedData, 'WLPT')

    #Call the function to write the costs dataframe to excel
    Dataframe_to_Excel(file_name, 'WLPT', outData)

    #Displays the dataframe on the screen
    #display(outData)


def MST(inData):
	#Sort data according to heuristic
	sortedData = deepcopy(inData)

	#Add a column to the dataframe by calculating the desired metric for all rows
	sortedData['MST'] = 'NaN'
	for index, row in sortedData.iterrows():
		sortedData.loc[index, 'MST'] = row['Deadline'] - row['Processing Time']

	#--> Sort the ratios in ascending order so that: (d[1] – p[1]) ≤ ( d[2] – p[2]) ≤ ... ≤ (d[n] – p[n])
	sortedData = sortedData.sort_values('MST')
	#display(sortedData) 

	#Calculates the costs associated with this task order
	outData = costAccounting(sortedData, 'MST')

	#Calls the function to write the cost dataframe to excel
	Dataframe_to_Excel(file_name, 'MST', outData)

	#Displays the dataframe on the screen
	#display(outData)


def EDD(inData):
	#Sort data according to heuristic
	#-->Deadline in ascending order
	sortedData = inData.sort_values('Deadline')

	#Calculate costs associated with this task order
	outData = costAccounting(sortedData, 'EDD')

	#Call function to write cost dataframe to excel
	Dataframe_to_Excel(file_name, 'EDD', outData)

	#Display dataframe on screen
	#display(outData)


def main():
	screen = Interface()

	while(True):
		#Tries to read the commands from the interface window
		try:
			screen.Read()
			screen.UpdateMSG('- - -')
		
		#If the window is closed it ends the program
		except:
			break
		
		#Tries to read the spreadsheet
		try:
			
			data = pd.read_excel(file_name, sheet_name=0, engine='openpyxl')
			#display(data)

			if(not(runSPT or runLPT or runWSPT or runWLPT or runEDD or runMST)):
				screen.UpdateMSG('No Heuristic Selected')

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
					
				screen.UpdateMSG('Calculations Completed')
		
		#Error message in case the spreadsheet is open
		except Exception as e:
			print(str(e))
			screen.UpdateMSG(f'Error reading data file\n{str(e)}')

if __name__ == "__main__":
    main()

