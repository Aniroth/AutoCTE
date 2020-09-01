# -*- coding: utf-8 -*-
from AutoCTE.SubRoutines.AutoActions import ClickOn
from AutoCTE.SubRoutines.AutoActions import WriteOn
from AutoCTE.SubRoutines.AutoActions import Sleep
from AutoCTE.SubRoutines.AutoActions import CheckFor
from AutoCTE.SubRoutines.AutoActions import PressKey
from AutoCTE.SubRoutines.AutoActions import WarningMSG
from AutoCTE.SubRoutines.AutoActions import FindDest
from AutoCTE.SubRoutines.AutoActions import AutoOBS
import xlwings as xw
import pyautogui
from datetime import date

def CTEauto (CTEs, tipo, flag):
	
	for CTE in CTEs:

		ClickOn('./AutoCTE/Buttons/LUPA.bmp')
		ClickOn('./AutoCTE/Buttons/INCLUIR.bmp')
		WriteOn(str(len(CTE.notas)))
		PressKey('tab', 3)
		PressKey('3')
		ClickOn('./AutoCTE/Buttons/SALVAR.bmp')
		Sleep(0.3)
		ClickOn('./AutoCTE/Buttons/OK.bmp')
		PressKey('tab', 2)
		WriteOn(CTE.coleta)
		PressKey('tab')

		if CheckFor('./AutoCTE/Buttons/CHECKLojaRem.bmp', 20) == False:
			WarningMSG('Erro loja remetente')

		WriteOn('01')
		Sleep(0.3)
		
		if tipo != "Global":
			if CTE.notas[0].cnpj == 'EXTERIOR':
				WriteOn(str(FindDest(CTE.notas[0].destinatario)))
				CTE.tipoCTE = str(4)
			else:
				WriteOn(str(FindDest(CTE.notas[0].cnpj, ext = False)))
				CTE.tipoCTE = str(0)
		else:
			PressKey('tab')

		WriteOn('01')
		Sleep(0.3)
		PressKey('delete')
		PressKey('tab')
		PressKey('delete')
		PressKey('tab', 3)
		Sleep(0.5)
		PressKey('1')
		ClickOn('./AutoCTE/Buttons/SERV_TRANSP.bmp', precision = 0.98)
		PressKey('3')
		PressKey('tab')
		PressKey('1')
		Sleep(0.5)

		if CheckFor('./AutoCTE/Buttons/CHECKescSer.bmp', 20):
			WriteOn(CTE.servico)
			PressKey('enter', 2)

		if (tipo == "Global"):
			PressKey('tab')
		else:
			PressKey('2')
			
		WriteOn(AutoOBS(CTE))
		PressKey('tab', 2)
		
		if CTE.seqEnd != None:
			WriteOn(str(CTE.seqEnd))
			if (len(str(CTE.seqEnd)) != 9):
				PressKey('tab')
		else:
			PressKey('tab')
		
		if tipo != "Global":
			WriteOn(str(CTE.tipoCTE))
			PressKey('tab', 3)
		else:
			WriteOn(str(CTE.tipoCTE))
			PressKey('tab', 3)
		
		PressKey('n')

		for i in range(len(CTE.notas)):

			if i == 0:
				ClickOn('./AutoCTE/Buttons/DOCTO_ENTRADA.bmp')

			PressKey('enter')
			WriteOn(str(CTE.notas[i].numeroNFE))
			PressKey('enter')
			PressKey('1')
			PressKey('enter')

			if CheckFor('./AutoCTE/Buttons/ERRO_SerieNFE.bmp', 20):
				WarningMSG('NFE já cadastrada')

			PressKey('enter')

			today = date.today()
			day = today.strftime("%d%m%y")

			WriteOn(str(day))
			PressKey('enter')
			
			if tipo == "Global":
				WriteOn('TERCEIROS0006OF')
			else:
				WriteOn('8000002000008')
				
			PressKey('enter')
			Sleep(0.3)
			PressKey('right')
			PressKey('enter')
			WriteOn(CTE.cntr)
			PressKey('right', 3)
			PressKey('enter')
			WriteOn(str(CTE.notas[i].pesoBruto))
			PressKey('enter')
			Sleep(0.3)
			PressKey('right')
			PressKey('enter')
			WriteOn(str(CTE.notas[i].valor))
			Sleep(0.3)

			if i == 0:

				PressKey('right', 5)
			else:

				PressKey('right')
				PressKey('0')
				PressKey('enter')
				PressKey('right', 3)

			PressKey('enter')
			
			if tipo == "Global":
				WriteOn('EXPO')
			else:
				WriteOn('CONHECIMENTO 000' + str(CTE.espelho))
				
			PressKey('enter')
			PressKey('enter')
			WriteOn(str(CTE.notas[i].qtde))
			PressKey('enter')
			PressKey('right', 23)
			PressKey('enter')
			WriteOn(str(CTE.notas[i].codigoXML))

			if i != len(CTE.notas)-1:

				PressKey('down')
		
		if tipo == "Global":
			ClickOn('./AutoCTE/Buttons/SALVAR.bmp')
			ClickOn('./AutoCTE/Buttons/SALVAR3.bmp')
		else:
			ClickOn('./AutoCTE/Buttons/SALVAR.bmp')
			ClickOn('./AutoCTE/Buttons/Cancelar3.bmp')