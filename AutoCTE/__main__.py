# -*- coding: utf-8 -*-
import tkinter as tk
import xlwings as xw
import ctypes
import re
import xmltodict
import os
import pygubu
import threading
from pygubu.builder import ttkstdwidgets
from win32api import GetKeyState
from win32con import VK_CAPITAL
from AutoCTE.Classes.cte import CTe
from AutoCTE.Classes.nfe import NFe
from AutoCTE.Classes.conhecimento import Conhecimento
from AutoCTE.Routines.EmitirCTE import CTEauto
from AutoCTE.Routines.AgendViag import AutoAgendamento
from AutoCTE.Routines.Calculo import Calcular
from AutoCTE.Routines.EncerramentoViag import Alterar
from AutoCTE.Routines.EncerramentoViag import Encerrar
from AutoCTE.Routines.Confimar import Confirmar
from AutoCTE.Routines.EncerrarEntrega import EncerrarKlabin
from AutoCTE.SubRoutines.AutoActions import ClickOn
from AutoCTE.SubRoutines.AutoActions import WriteOn
from AutoCTE.SubRoutines.AutoActions import Sleep
from AutoCTE.SubRoutines.AutoActions import CheckFor
from AutoCTE.SubRoutines.AutoActions import PressKey
from AutoCTE.SubRoutines.AutoActions import FormatTXT
import xlwings as xw
import pyautogui
from tkinter import messagebox

CTes = []
Espelhos = []
NFes = []
flagEvent = threading.Event()

def main():
				
	FormatTXT()
	
	if GetKeyState(VK_CAPITAL):
		pyautogui.press('capslock')
	
	ShowMainWindow()


def LoadNFEs():
	
	path = 'path'
	
	if (os.path.isdir('U:/AutoCTE/')):
		path = 'U:/AutoCTE/'
	else:
		path = 'S:/AutoCTE/'
	
	XMLs = os.listdir(path + 'Novos')
	
	for x in XMLs:
		tempNFE = ReadXML(x)
		print('Nota: ' + tempNFE.numeroNFE + ' foi encontrada.')
		NFes.append(tempNFE)


def LoadConhecimentos():
	
	path = 'path'
	
	if (os.path.isdir('U:/AutoCTE/')):
		path = 'U:/AutoCTE/'
	else:
		path = 'S:/AutoCTE/'
	
	Conhecimentos = os.listdir(path + 'Conhecimentos')
	
	for x in Conhecimentos:
		tempEspelho = Conhecimento()
		with open(path + 'Conhecimentos/' + x, "r+") as f:
			
			txtEspelho = f.read()
			
			for NF in NFes:
				
				if NF.numeroNFE in txtEspelho:
					
					isDuplicate = False
					for i in range(len(tempEspelho.notas)):
						if NF.numeroNFE == tempEspelho.notas[i].numeroNFE:
							isDuplicate == True

					if isDuplicate == False:
						tempEspelho.notas.append(NF)
		
		tempEspelho.numero = str(x.rstrip('.txt'))
		print('Espelho: ' + tempEspelho.numero + ', com ' + str(len(tempEspelho.notas)) + ' notas foi encontrado.')
		Espelhos.append(tempEspelho)


def CTEKlabin(flag):
	
	emptyCTE = False
	LoadNFEs()
	LoadConhecimentos()
	
	workbook = xw.Book('./AutoCTE/MODELO.xlsm')
	sheet = workbook.sheets['CTE-Klabin']
	total = sheet.range('A1').current_region.last_cell.row

	for i in range(total):
		
		if sheet['A' + str(2+i)].value != 'ABERTO':
			continue
		
		tempCTE = CTe()
		tempCTE.motorista = sheet.range('C' + str(2+i)).value
		tempCTE.placas = sheet.range('D' + str(2+i)).value
		tempCTE.booking = sheet.range('B' + str(2+i)).value
		tempCTE.cntr = sheet.range('G' + str(2+i)).value
		tempCTE.cntr.replace(" ", "")
		tempCTE.seqEnd = sheet.range('F' + str(2+i)).value
		tempCTE.coleta = '000' + sheet.range('E' + str(2+i)).value
		tempCTE.espelho = sheet.range('H' + str(2+i)).value
		tempCTE.servico = sheet.range('k' + str(2+i)).value
		sheet.range('I' + str(2+i)).value = None
		
		tempEspelho = None
		
		for espelho in Espelhos:
			
			for nf in espelho.notas:
				
				if tempCTE.cntr in nf.extraINFO.upper():
					sheet.range('H' + str(2+i)).value = espelho.numero
					tempEspelho = espelho
		
		if tempCTE.espelho != None:
			
			for espelho in Espelhos:
				
				if tempCTE.espelho == espelho.numero:
					tempEspelho = espelho
					break
		
		if tempEspelho == None:
			emptyCTE = True
			continue
			
		for nf in tempEspelho.notas:
								
			tempCTE.notas.append(nf)
			
			sheet.range('J' + str(2+i)).value = nf.destinatario + ' - ' + nf.cnpj				
			
			if sheet.range('I' + str(2+i)).value == None:
				sheet.range('I' + str(2+i)).value = nf.numeroNFE + ' - '
			else:
				sheet.range('I' + str(2+i)).value = sheet.range('I' + str(2+i)).value + nf.numeroNFE + ' - '
		
		CTes.append(tempCTE)
	
	if emptyCTE:
		ctypes.windll.user32.MessageBoxW(0, "Não foi possível localizar um ou mais contâineres, insira manualmente o ESPELHO e clique em CTE-Klabin novamente", "ERRO!!!", 1)
		print('Aborted')
		return
	
	print('Ended loading...')
	CTEauto(CTes, "Klabin", flag)


def CTEGlobal(flag):

	workbook = xw.Book('./AutoCTE/MODELO.xlsm')
	sheet = workbook.sheets['Agendamento']
	total = sheet.range('P1').current_region.last_cell.row
	print (total)
	sheetAux = workbook.sheets['BancoConjunto']
	totalAux = sheetAux.range('A1').current_region.last_cell.row

	for i in range(total):
		
		if sheet.range('P' + str(2+i)).value == None:
			continue
		
		tempCTE = CTe()
		
		tempCTE.motorista = str(sheet.range('L' + str(2+i)).value)
		tempCTE.booking = str(sheet.range('U' + str(2+i)).value)
		tempCTE.cntr = str(sheet.range('M' + str(2+i)).value)
		tempCTE.coleta = '000' + str(sheet.range('P' + str(2+i)).value)
		tempCTE.servico = str(sheet.range('T' + str(2+i)).value)
		tempCTE.tipoCTE = str(sheet.range('R' + str(2+i)).value)
		codConj = str(sheet.range('K' + str(2+i)).value)
		
		for j in range(totalAux):
			
			if codConj == sheetAux.range('F' + str(2+j)).value:
				codCav = str(sheetAux.range('D' + str(2+j)).value)
				codCar = str(sheetAux.range('E' + str(2+j)).value)
		
		tempCTE.placas = codCav + ' / ' + codCar
		
		tempNFE = NFe()
		tempNFE.pesoBruto = str(sheet.range('N' + str(2+i)).value)
		tempNFE.valor = str(sheet.range('O' + str(2+i)).value)
		tempNFE.qtde = str(sheet.range('S' + str(2+i)).value)
		tempNFE.codigoXML = str(sheet.range('V' + str(2+i)).value)
		tempNFE.numeroNFE = str(sheet.range('Q' + str(2+i)).value)
		
		tempCTE.notas.append(tempNFE)
		CTes.append(tempCTE)
		
		
	CTEauto(CTes, "Global", flag)


def ReadXML(xmlNumber):

	path = 'path'
	
	if (os.path.isdir('U:/AutoCTE/')):
		path = 'U:/AutoCTE/'
	else:
		path = 'S:/AutoCTE/'

	xml = None
	newNFE = NFe()

	with open(path + 'Novos/' + xmlNumber, 'r', encoding='utf-8') as fd:
		xmlString = fd.read()
		xml = xmltodict.parse(xmlString)

	xmlPattern = '\d{44}'
	newNFE.codigoXML = str(re.findall(xmlPattern , xmlNumber))[2:46]
	print ("Reading: " + str(newNFE.codigoXML))
	newNFE.data = xml['nfeProc']['NFe']['infNFe']['ide']['dhEmi'][0:10]
	newNFE.destinatario = xml['nfeProc']['NFe']['infNFe']['dest']['xNome']
	newNFE.remetente = xml['nfeProc']['NFe']['infNFe']['emit']['xFant']
	newNFE.numeroNFE = xml['nfeProc']['NFe']['infNFe']['ide']['nNF']
	checkCFOP = newNFE.cfop = xml['nfeProc']['NFe']['infNFe']['det']
	
	if type(checkCFOP) == list:
		newNFE.cfop = xml['nfeProc']['NFe']['infNFe']['det'][0]['prod']['CFOP']
	else:	
		newNFE.cfop = xml['nfeProc']['NFe']['infNFe']['det']['prod']['CFOP']

	newNFE.qtde = xml['nfeProc']['NFe']['infNFe']['transp']['vol']['qVol']
	newNFE.valor = xml['nfeProc']['NFe']['infNFe']['total']['ICMSTot']['vNF']
	newNFE.pesoBruto = xml['nfeProc']['NFe']['infNFe']['transp']['vol']['pesoB']
	newNFE.extraINFO = xml['nfeProc']['NFe']['infNFe']['infAdic']['infCpl']
	newNFE.extraINFO.replace(" ", "")

	if (xml['nfeProc']['NFe']['infNFe']['dest']['enderDest']['xMun'] == "EXTERIOR"):

		newNFE.cnpj = "EXTERIOR"
	else:

		newNFE.cnpj = xml['nfeProc']['NFe']['infNFe']['dest']['CNPJ']

	return newNFE

###############################################################################
#                            UI Controller                                    #
###############################################################################

def MsgYES_NO(msg):
	
	import ctypes
	return ctypes.windll.user32.MessageBoxW(0, msg, "title", 1)
	
def ShowMainWindow():
	root = tk.Tk()
	root.title("Auxílio TMS")
	app = MyApplication(root)
	root.mainloop()

class MyApplication(pygubu.TkApplication):

	def _create_ui(self):

		self.builder = builder = pygubu.Builder()
		builder.add_from_file('./AutoCTE/GUI/Auto TMS - main.ui')
		self.mainwindow = builder.get_object('mainWindow', self.master)
		builder.connect_callbacks(self)

	def autoCTE_Clicked(self):
		flagEvent.set()
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		x = threading.Thread(target=CTEKlabin, args=(flagEvent,), daemon=True)
		x.start()

	def CTE_Clicked(self):
		flagEvent.set()		
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		x = threading.Thread(target=CTEGlobal, args=(flagEvent,), daemon=True)
		x.start()

	def autoAgendamento_Clicked(self):
		flagEvent.set()
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		x = threading.Thread(target=AutoAgendamento, args=(flagEvent,), daemon=True)
		x.start()
	
	def autoConfirma_Clicked(self):
		flagEvent.set()
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		qtdeViag = self.builder.tkvariables.__getitem__('qtdeViag').get()
		codViag = self.builder.tkvariables.__getitem__('codViag').get()
		x = threading.Thread(target=Confirmar, args=(qtdeViag, codViag, flagEvent), daemon=True)
		x.start()
		
	def autoEncerrar_AlterarRepom(self):
		flagEvent.set()		
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		viagREPOM = self.builder.tkvariables.__getitem__('viagREPOM').get()
		x = threading.Thread(target=Alterar, args=(int(viagREPOM), flagEvent,), daemon=True)
		x.start()

	def autoEncerrar_Clicked(self):
		flagEvent.set()		
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		viagREPOM = self.builder.tkvariables.__getitem__('viagREPOM').get()
		x = threading.Thread(target=Encerrar, args=(int(viagREPOM), flagEvent,), daemon=True)
		x.start()
		
	def autoEncerrarEntrega_Clicked(self):
		flagEvent.set()		
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		viagREPOM = self.builder.tkvariables.__getitem__('viagREPOM').get()
		x = threading.Thread(target=EncerrarKlabin, args=(int(viagREPOM), flagEvent,), daemon=True)
		x.start()

	def autoCalcular_Clicked(self):
		flagEvent.set()		
		if GetKeyState(VK_CAPITAL):
			pyautogui.press('capslock')
		viag = self.builder.tkvariables.__getitem__('viag').get()
		x = threading.Thread(target=Calcular, args=(int(viag), flagEvent,), daemon=True)
		x.start()
		
	def abrirExcel(self):
		ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
		print ('Opening EXCEL')
		os.system('start excel.exe ' + ROOT_DIR + '\MODELO.xlsm')