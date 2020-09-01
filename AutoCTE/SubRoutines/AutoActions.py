# -*- coding: utf-8 -*-
from AutoCTE.SubRoutines.ImageSearch import imagesearch
from AutoCTE.SubRoutines.ImageSearch import click_image
from AutoCTE.SubRoutines.ImageSearch import imagesearcharea
import xlwings as xw
import pyautogui
import time
import os
import re

delay = 0.5
pyautogui.PAUSE = 0.001

def ClickOn(path = None, repeat = 999999, precision = 0.95, x1 = 0, y1 = 0, x2 = 0, y2 = 0):
	
	if (path == None):
		Sleep()
		pyautogui.click()
		print('Single Click')
		return
	
	print('Looking for: ' + path + '...')
	
	if ((x1 == 0) and
		(y1 == 0) and
		(x2 == 0) and
		(y2 == 0)):
			
		for i in range(repeat):
	
			pos = imagesearch(path, precision)
			if pos[0] != -1:
				Sleep()
				print('Clicking in: ' + path)
				click_image(path, pos, "left", 0, offset=0)
				return True
	else:
		
		for i in range(repeat):
			
			print('at x1 = ' + str(x1) + ' , y1 = ' + str(y1) + ' || x2 = ' + str(x2) + ' , y2 = ' + str(y2))
			pos = imagesearcharea(path, x1, y1, x2, y2, precision)
			if pos[0] != -1:
				Sleep()
				print('Clicking in: ' + path)
				click_image(path, pos, "left", 0, offset=0)
				return True
	
	print ('Not found: ' + path)
	return False


def WriteOn(msg):

	Sleep()
	print('Writing: ' + msg)
	pyautogui.typewrite(msg)


def PressKey(key, repeat = 1):
	
	print('Pressing: ' + key + ' x' + str(repeat))
	for i in range(repeat):

		if (
				(key == 'up') or
			    (key == 'down') or
			    (key == 'right') or
			    (key == 'left') or
			    (key == 'enter') or
			    (key == 'tab')
		    ):
			Sleep(0.1)
			
		else:		
		
			Sleep()
		
		pyautogui.press(key)


def CheckFor(path, repeat = 999999, precision = 0.8):
	
	print('Looking for: ' + path + '...')
	for i in range(repeat):
		
		pos = imagesearch(path, precision)
		if pos[0] != -1:
			pyautogui.moveTo(pos[0], pos[1])
			print ('Found: ' + path)
			return True
	
	print ('Not found: ' + path)
	return False


def WarningMSG(msg):

	import ctypes
	return ctypes.windll.user32.MessageBoxW(0, msg, 'ERRO NO PROCESSO', 0)


def FindDest(reference, ext = True):
	
	path = 'path'
	
	if (os.path.isdir('U:/AutoCTE/')):
		path = 'U:/AutoCTE/'
	else:
		path = 'S:/AutoCTE/'
	
	workbook = xw.Book(path + 'Banco_Dados.xlsx')
	sheet = workbook.sheets['CódigoCliente']
	total = sheet.range('B1').current_region.last_cell.row

	for i in range(total):
		
		if (ext):
			temp = sheet.range('B' + str(i+1)).value
		else:		
			temp = sheet.range('C' + str(i+1)).value
			
		if str(temp) == str(reference):
			return sheet.range('E' + str(i+1)).value
	
	
	print('Not found dest: ' + reference)
	WarningMSG('Destinatário não cadastrado:\n' + reference)


def AutoOBS(CTE):

	obs = 'NF: ' + CTE.notas[0].numeroNFE

	if len(CTE.notas) > 1:

		for i in range(len(CTE.notas)):
			if i == 0:
				continue
			obs = obs + ' - ' + CTE.notas[i].numeroNFE

	obs = obs + ' / BOOKING: ' + CTE.booking
	obs = obs + ' / MOT: ' + CTE.motorista
	obs = obs + ' / ' + CTE.placas

	return obs


def LoopSeqEnd(CTE):

	while(True):
		
		if CTE.seqEnd == 'ITAPOA':
	
			if CheckFor('./AutoCTE/Buttons/SEQ_END_ITAPOA.bmp', 50, 0.99):
				Sleep(0.5)
				return 0
			else:
				PressKey('down')
	
		if CTE.seqEnd == 'ITAJAI':
	
			if CheckFor('./AutoCTE/Buttons/SEQ_END_ITJ.bmp', 50, 0.99):
				Sleep(0.5)
				return 0
			else:
				PressKey('down')
	
		if CTE.seqEnd == 'NAVEGANTES':
	
			if CheckFor('./AutoCTE/Buttons/SEQ_END_NVT.bmp', 50, 0.99):
				Sleep(0.5)
				return 0
			else:
				PressKey('down')


def Sleep(sleepFactor = delay):
	
	if sleepFactor > delay:
		print ('Sleeping (' + str(sleepFactor) + ')')
	
	time.sleep(sleepFactor)

def FormatTXT():
		
	_path = 'path'
	
	if (os.path.isdir('U:/AutoCTE/')):
		_path = 'U:/AutoCTE/'
	else:
		_path = 'S:/AutoCTE/'
	
	patternName = 'Número Conhecimento: \d{10}'
	patternNotas = '\d{4}\/\d{2}[0-3]\d\/[01]\d\/[12][09]\d{11}'
	conhecimentos = os.listdir(_path + 'Conhecimentos')
	
	for oldName in conhecimentos:
		
		if len(oldName) == 11:
			continue
		
		path = _path + 'Conhecimentos/' + oldName
		print (path)
		newName  = "XX"
		
		with open(path, "r+") as f:
			oldFile = f.read()
			text = oldFile.replace("\n", "")
			notas = re.findall(patternNotas, text)
			result = re.findall(patternName, text)
			
			text = ""
			
			for nota in notas:
				text = text + nota[-9:].lstrip('0')  + '\n'
			print(result)
			newName = result[0][-7:]
			print(str(newName) + ' - ' + str(len(notas)))
			f.seek(0)
			f.truncate(0)
			f.write(text)
		
		os.rename(path,_path + 'Conhecimentos/' + newName + '.txt')