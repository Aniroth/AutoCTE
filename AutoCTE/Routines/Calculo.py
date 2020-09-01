# -*- coding: utf-8 -*-
from AutoCTE.SubRoutines.AutoActions import ClickOn
from AutoCTE.SubRoutines.AutoActions import Sleep
from AutoCTE.SubRoutines.AutoActions import CheckFor
from AutoCTE.SubRoutines.AutoActions import PressKey

def Calcular(viag, flag):
	
	for i in range(viag):

		ClickOn("./AutoCTE/Buttons/CALCULAR.bmp")
		ClickOn("./AutoCTE/Buttons/CALCULAR.bmp")
		ClickOn("./AutoCTE/Buttons/OK2.bmp")
		if CheckFor("./AutoCTE/Buttons/CONFIRMAR3.bmp"):
			Sleep(2)
		ClickOn("./AutoCTE/Buttons/CONFIRMAR3.bmp")
		if CheckFor("./AutoCTE/Buttons/CHECKManifesto.bmp"):
			Sleep(1)
			ClickOn('./AutoCTE/Buttons/CHECKDigitado.bmp')
		PressKey('down')