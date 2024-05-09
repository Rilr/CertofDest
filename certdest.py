import pandas as pd
import PySimpleGUI as sg
import docx
import os
import sys

layout = [[sg.Text('HDD Data:')],
          [sg.InputText(key = 'inpath'), 
           sg.FileBrowse(file_types =[( 'XLSX Files','*.xlsx')])],
          
          [sg.Text('Create Template?'),
           sg.Radio('Yes', 'template', key='yes'),
           sg.Radio('No', 'template', key='no', default=True)],
          
          [sg.Text('Where to Save:')],
          [sg.InputText(key = 'outpath'),
           sg.FolderBrowse()]
          
          [sg.Button('Submit'), sg.Button('Cancel')]]

window = sg.Window('CoD Data Import', layout)
event, values = window.read()

outpathFolder = values['outpath']
outpathFile = outpathFolder + '/certificate-of-destruction.docx'
templateFile = outpathFolder + '/certificate-of-destruction_template.xlsx'

hddData = {
    'file_path': values['inpath'],
    'brand': 'Brand',
    'model': 'Model',
    'size': 'Capacity',
    'serial': 'Serial',
    'mdate': 'Manufacture Date',
    'ddate': 'Destruction Date',
    'dmethod': 'Destruction Method',
}

#pd.read_excel(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col']], engine='openpyxl')

def createTemplate():
    df = pd.DataFrame(columns=[value for key, value in hddData.items() if key != 'file_path'])
    df.to_excel(outpathFile, index=False)

# def gatherData():
#     pd.read_excel(hddData['file_path'], header=0, usecols=[hddData['brand'], hddData['model'], hddData['size'], hddData['serial'], hddData['mdate'], hddData['ddate'], hddData['dmethod']], engine='openpyxl')
    
# def dataInsertion():
#     pass
    
#TODO: Add error exception for if the file is open and cannot be written to
# 
def main():
    while True:
        if event in (sg.WIN_CLOSED, 'Cancel'):
            break
        if event == "Submit":
            if values['yes']:
                createTemplate()
                break
            continue
        continue
    window.close()
    
    
if __name__ == '__main__':
    main()