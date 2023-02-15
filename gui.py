from PySimpleGUI import PySimpleGUI as sg 
from converter import convert_excel_to_pptx

layout = [
    [sg.Text('Dossier Excel'), sg.Input(key='excel_dir'), sg.FolderBrowse('parcourir')],
    [sg.Text('Dossier Powerpoint'), sg.Input(key='pptx_dir'), sg.FolderBrowse('parcourir')],
    [sg.Button('Convertir les fichiers'), sg.Button('Annuler')]
]

window = sg.Window('convertisseur de déroulés', layout)


while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Annuler'):
        break
    if event == 'Convertir les fichiers':
        excel_dir = values['excel_dir']
        pptx_dir = values['pptx_dir']
        # Call the function that converts Excel to PowerPoint files
        convert_excel_to_pptx(excel_dir, pptx_dir)
        sg.popup('Conversion terminée!')
