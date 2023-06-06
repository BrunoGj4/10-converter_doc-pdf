from os import path, listdir
from docx2pdf import convert
from tkinter import filedialog
import PySimpleGUI as sg
import aspose.words as aw


layout = [
   [sg.Text('Converta word p/ PDF')],
   [sg.Button('Converter')]
]

janela = sg.Window('Converter', layout)

while True:
   # ext = ['doc', 'docx']
   
   event, values = janela.read()
   if event == sg.WIN_CLOSED:
      break
   
   try:
      pasta = filedialog.askdirectory()
      caminho_pasta = path.join(pasta)
      # caminho_arquivos = listdir()
      lista_arquivos = listdir(pasta)

      if lista_arquivos and caminho_pasta:
         for item in range(len(lista_arquivos)):
            arquivo = lista_arquivos[item]
            convert(caminho_pasta + '/' + arquivo)
            # aw.Document(caminho_pasta + '/' + arquivo).save(caminho_pasta + '/' + arquivo[:-4] + '.pdf')
         sg.popup('Conversão Concluída.')
         break
      else:
         sg.popup('Não existem arquivos na pasta indicada')
   except FileNotFoundError:
      pass
