# -*- coding: utf-8 -*-
from AutoCTE.SubRoutines.AutoActions import ClickOn
from AutoCTE.SubRoutines.AutoActions import WriteOn
from AutoCTE.SubRoutines.AutoActions import Sleep
from AutoCTE.SubRoutines.AutoActions import CheckFor
from AutoCTE.SubRoutines.AutoActions import PressKey
import xlwings as xw

def AutoAgendamento(flag):
	
	workbook = xw.Book('./AutoCTE/MODELO.xlsm')
	sheet = workbook.sheets['Agendamento']
	total = sheet.range('A1').current_region.last_cell.row
	isFirstLine = True
	
	for i in range(total):
		
		if sheet.range('A' + str(2+i)).value != 'ABERTO':
			continue
		
		solicitante = str(sheet.range('B' + str(2+i)).value)
		remetente = str(sheet.range('D' + str(2+i)).value)
		seqEnd = str(sheet.range('F' + str(2+i)).value)
		destinatario = str(sheet.range('G' + str(2+i)).value)
		consignatario = str(sheet.range('I' + str(2+i)).value)
		codConj = str(sheet.range('K' + str(2+i)).value)
		cntr = str(sheet.range('M' + str(2+i)).value)
		peso = str(sheet.range('N' + str(2+i)).value)
		valor = str(sheet.range('O' + str(2+i)).value)
		codMot = None
		codCav = None
		codCar = None
		sheetAux = workbook.sheets['BancoConjunto']
		totalAux = sheetAux.range('A1').current_region.last_cell.row
		
		for j in range(totalAux):
			
			if codConj == sheetAux.range('F' + str(2+j)).value:
				print(str(j) + ' = ' + codConj + ' - ' + sheetAux.range('F' + str(2+j)).value)
				codMot = str(sheetAux.range('C' + str(2+j)).value)
				codCav = str(sheetAux.range('D' + str(2+j)).value)
				codCar = str(sheetAux.range('E' + str(2+j)).value)
		
		if isFirstLine:
			
			CheckFor('./AutoCTE/Buttons/CHECKCodSolicitante.bmp', 1000)
			WriteOn(solicitante)
		
			if seqEnd == 'X':
				
				PressKey('f3')
				ClickOn('./AutoCTE/Buttons/CONFIRMAR4.bmp')
				PressKey('enter')
			
			ClickOn('./AutoCTE/Buttons/REPETIR.bmp')
			ClickOn('./AutoCTE/Buttons/ITEM1.bmp')
			
			PressKey('right', 5)
			WriteOn(remetente)
			Sleep(0.3)
			PressKey('right', 2)
			
			if seqEnd == 'X':
				
				PressKey('f3')
				ClickOn('./AutoCTE/Buttons/CONFIRMAR4.bmp')
				PressKey('enter')
	
			else:
				
				PressKey('right')
	
			PressKey('right', 5)
			PressKey('1')
			PressKey('right', 9)
			WriteOn(destinatario)
			PressKey('right', 11)
			PressKey('2')
			WriteOn(consignatario)
			PressKey('right', 35)
		
		else:
			
			PressKey('right', 73)	
			
		WriteOn(codCav)
		PressKey('enter')
		ClickOn('./AutoCTE/Buttons/SIM.bmp', 80, 0.8)
		PressKey('right')
		WriteOn(codCar)
		PressKey('enter')
		ClickOn('./AutoCTE/Buttons/SIM.bmp', 80, 0.8)
		PressKey('right', 3)
		WriteOn(codMot)
		
		if len(codMot) != 6:
			
			PressKey('enter')
			
		ClickOn('./AutoCTE/Buttons/SIM.bmp', 80, 0.8)
		PressKey('right')
		WriteOn(cntr)
		PressKey('enter')
		PressKey('f4')
		
		looking = True
		while looking:
			
			if (
					(CheckFor('./AutoCTE/Buttons/CHECKTerceiros.bmp', 10, 0.9))      or
					(CheckFor('./AutoCTE/Buttons/CHECKTerceirosCinza.bmp', 10, 0.9))
				):
				
				looking = False
				PressKey('right', 4)
				PressKey('1')
				PressKey('enter')
				PressKey('right')
				WriteOn(peso)
				PressKey('enter')
				PressKey('right')
				WriteOn(valor)
				PressKey('right', 6)
				WriteOn(cntr)
				Sleep(0.5)
				ClickOn('./AutoCTE/Buttons/SALVAR4.bmp')
				
			elif CheckFor('./AutoCTE/Buttons/CHECKBlankProduto.bmp', 10, 0.9):
				looking = False
				WriteOn('TERCEIROS0006OF')
				PressKey('right')
				WriteOn('h40')
				PressKey('right')
				PressKey('1')
				PressKey('enter')
				PressKey('right')
				WriteOn(peso)
				PressKey('enter')
				PressKey('right')
				WriteOn(valor)
				PressKey('right', 6)
				WriteOn(cntr)
				Sleep(0.5)
				ClickOn('./AutoCTE/Buttons/SALVAR4.bmp')
								
			else:
				
				PressKey('delete')
				PressKey('down')

		print(solicitante + ' - ' + str(sheet.range('B' + str(3+i)).value))
		if solicitante == str(sheet.range('B' + str(3+i)).value):
			
			sheet.range('A' + str(2+i)).value = 'LANÇADO'
			PressKey('down')
			isFirstLine = False
			
		else:
			
			sheet.range('A' + str(2+i)).value = 'LANÇADO'
			#workbook.save("U:/AutoCTE/LISTA.xlsx")
			ClickOn('./AutoCTE/Buttons/SALVAR.bmp')
			Sleep(5.0)
			ClickOn('./AutoCTE/Buttons/INCLUIR2.bmp')
			isFirstLine = True