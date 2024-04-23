import pandas as pd
import PySimpleGUI as sg
import docx
import os
import sys

layout = [[sg.Text('HDD Data:')],
          [sg.InputText(key = 'inpath'), 
           sg.FileBrowse(file_types =[( 'XLSX Files','*.xlsx')])],
          
          [sg.Text('Where to Save:')],
          [sg.InputText(key = 'outpath'),
           sg.FolderBrowse()]
          
          [sg.Button('Ok'), sg.Button('Cancel')]]

window = sg.Window('CoD Data Import', layout)
event, values = window.read()

outpathFolder = values['outpath']
outpathFile = outpathFolder + '/cert-of-destruction.docx'

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

def gatherData():
    
def templateCreation():
    
def dataInsertion():
    
def main():
    gatherData()
    templateCreation()
    dataInsertion()

if __name__ == '__main__':
    main()