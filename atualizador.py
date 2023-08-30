#Interface
import PySimpleGUI as sg
import psutil
import requests
import time
import wget
import os

tela = None

def baixarArquivo():
    if os.path.isfile('C:/Temp/atualizacao.zip'):
        os.remove('C:/Temp/atualizacao.zip')
        
    wget.download('http://portugues.wiki.br/corretor/corretor.zip', 'C:/Temp/atualizacao.zip', bar=bar_custom)
    print('Download concluído')
    
def progress(valor):
    global tela
    progress_bar = tela['progressbar']    
    progress_bar.UpdateBar(valor)

def encerraCorretor():
    PROCNAME = "corretor.exe"

    for proc in psutil.process_iter():
        # check whether the process name matches
        if proc.name() == PROCNAME:
            proc.kill()

def bar_custom(atual, total, largura=80): 
    global tela
    valor = int(atual/total*100)
    tela['titulo'].Update('Atualizando... Aguarde...')
    tela['status'].Update('Realizando download do arquivo... ' + str(valor) + '% Concluído')
    tela['btn_atualizar'].Update(visible=False)
    tela['btn_cancelar'].Update(visible=False)
    progress(valor) 

def telaAtualizador():       
    global tela
    sg.theme('Reddit')
    
    alcadas = [    
                [sg.Text(text='Nova Versão Disponível', text_color="Black", font=("Arial",12, "bold"), expand_x=True, justification='center', key='titulo')],
                [sg.Text(text='Deseja atualizar agora?', text_color="Black", font=("Arial",9, "bold"), expand_x=True, justification='center', key='status')],
                [sg.ProgressBar(100, orientation='h', size=(30, 10), key='progressbar', visible=True, expand_x=True)],                  
                [sg.Button('Atualizar Agora', key='btn_atualizar'), sg.Button('Depois', key='btn_cancelar')]      
              ]
    
    tela = sg.Window('Atualização Disponível', alcadas)                

    while True:                
        eventos, valores = tela.read(timeout=0.1)

        if eventos == 'btn_atualizar':            
            encerraCorretor()
            baixarArquivo()
            return None        

        if eventos is None or eventos == "Cancelar":            
            tela.close() 
        
        
            
telaAtualizador()