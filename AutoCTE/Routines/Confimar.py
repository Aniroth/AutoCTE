# -*- coding: utf-8 -*-
from AutoCTE.SubRoutines.AutoActions import ClickOn
from AutoCTE.SubRoutines.AutoActions import WriteOn
from AutoCTE.SubRoutines.AutoActions import Sleep
from AutoCTE.SubRoutines.AutoActions import CheckFor
from AutoCTE.SubRoutines.AutoActions import PressKey
import xlwings as xw
import pyautogui

def Confirmar (qtde, cod, flag):
	
	ClickOn("./AutoCTE/Buttons/CHECKAgendamento.bmp")
	PressKey("2")
	ClickOn("./AutoCTE/Buttons/CONFIRMAR5.bmp")
	
	for i in range(int(qtde)):
		
		#ClickOn("./AutoCTE/Buttons/SALVAR2.bmp")
		ClickOn("./AutoCTE/Buttons/FECHAR2.bmp")
		ClickOn("./AutoCTE/Buttons/FILTRAR.bmp")
		ClickOn()
		CheckFor("./AutoCTE/Buttons/OK3.bmp")
		PressKey("tab")
		WriteOn(str(cod))
		ClickOn("./AutoCTE/Buttons/OK3.bmp")
		ClickOn()
		Sleep()
		ClickOn("./AutoCTE/Buttons/OK4.bmp")
		ClickOn()