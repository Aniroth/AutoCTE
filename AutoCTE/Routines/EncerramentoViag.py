# -*- coding: utf-8 -*-
from AutoCTE.SubRoutines.AutoActions import ClickOn
from AutoCTE.SubRoutines.AutoActions import WriteOn
from AutoCTE.SubRoutines.AutoActions import Sleep
from AutoCTE.SubRoutines.AutoActions import CheckFor
from AutoCTE.SubRoutines.AutoActions import PressKey
import xlwings as xw
import pyautogui

def Alterar(qtde, flag):
	
	for i in range(qtde):
		
		CheckFor('./AutoCTE/Buttons/X.bmp')
		Sleep(1)
		pyautogui.keyDown ('altleft')
		WriteOn('a')
		pyautogui.keyUp ('altleft')
		ClickOn("./AutoCTE/Buttons/SIM.bmp", 80, 0.8)
		ClickOn("./AutoCTE/Buttons/OK.bmp")
		CheckFor("./AutoCTE/Buttons/CHECKLupa.bmp")
		ClickOn("./AutoCTE/Buttons/OutrasAcoes.bmp", precision = 0.98)
		ClickOn("./AutoCTE/Buttons/CompVia.bmp")
		ClickOn("./AutoCTE/Buttons/OperFrotas.bmp")
		WriteOn('01')
		ClickOn("./AutoCTE/Buttons/CodVeiculo.bmp")
		PressKey('enter')
		Sleep(1)
		PressKey('enter')
		ClickOn("./AutoCTE/Buttons/SIM.bmp", 80, 0.8)
		Sleep(1)
		while CheckFor('./AutoCTE/Buttons/CHECKValidando.bmp', 1, 0.95):
			Sleep()
		ClickOn("./AutoCTE/Buttons/SALVAR.bmp")
		CheckFor("./AutoCTE/Buttons/CHECKLupa.bmp")
		ClickOn("./AutoCTE/Buttons/SALVAR.bmp")
		Sleep(4)
		PressKey('down')
	
def Encerrar(qtde, flag):
	
	for i in range(qtde):
		ClickOn("./AutoCTE/Buttons/OK2.bmp")
		ClickOn("./AutoCTE/Buttons/SIM.bmp")
		ClickOn("./AutoCTE/Buttons/SALVAR.bmp")
		ClickOn("./AutoCTE/Buttons/OK2.bmp")