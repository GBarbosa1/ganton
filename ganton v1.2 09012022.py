#ctypes é a biblioteca para compatibilidade que possibilita a utilização de funções DLL dentro do windows
#openpuxl nos possibilita a integrar uma planílha de excell (nosso banco de dados) para obter os ativos
#pyautogui e keyboard nos dão acesso e controle ao teclado

import ctypes
from openpyxl import Workbook, load_workbook
import pyautogui as pg
import keyboard
import time
wb = load_workbook(filename='Ativos.xlsx') #carrega a planilha com ativos monitorados
sh = wb.active
controle = True #variaveis que utilizei para controle de fluxo
contador = 1

def Mbox(text, title, style): #função para o pop-up de sistema que indica se ganton está desligado ou ligado
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


while True: #loop principal do programa
    
    if keyboard.is_pressed('end'): #desliga e liga ganton
        
            if controle == True:
                time.sleep(0.2)
                controle = False
                Mbox('Desligado', 'Aviso de sistema', 0)
                
            elif controle == False:
                time.sleep(0.2)
                controle = True
                Mbox('Ligado', 'Aviso de sistema',0)

    if keyboard.is_pressed('page up') and controle == True: #cicla ativos para cima (para baixo na tabela)
            contador +=1
            pg.typewrite(sh['A'+str(contador)].value,0.005)
            pg.typewrite('\n')

        
        
    if keyboard.is_pressed('page down') and controle == True: #cicla ativos para baixo (para cima na tabela)
        contador -=1
        pg.typewrite(sh['A'+str(contador)].value,0.005)
        pg.typewrite('\n')
        
        
#Gabriel Barbosa
#ver1.1
#09/01/2021

