"""
given a drilling program as .xlsx:

1. create a GSL logsheet data source
2. create a DSL logsheet data source
"""

import os
import sys
import time

import pandas as pd
import numpy as np
from configparser import ConfigParser
from PyQt5 import QtCore, QtGui, QtWidgets

config = ConfigParser()
config.read('config.ini')
last_used = config.get('last_used', 'concession')

def continue_prompt(msg='Continue? [Y/N] '):
    proceed = ''
    while proceed != 'Y':
        proceed = input(msg)
        if proceed.upper() == 'N':
            return False
        elif proceed.upper() == 'Y':
            return True
        else:
            print("Epic fail! Please respond with Y or N, not " + proceed)
            time.sleep(1.2)

def selectfile():
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QFileDialog()
    name = dialog.getOpenFileName(None, 'Open Drilling Program xlsx', config.get('last_used', 'input'))[0]
    return name

def selectdir():
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QFileDialog()
    name = dialog.getExistingDirectory(None, 'Select Output Directory', config.get('last_used', 'output'))
    return name

def export(filename, data): # to csv
    data.to_csv(output_dir + filename, index=False)
    print(f'Exported as {filename}')

def append_fields():
    dp = pd.read_excel(input_dp, usecols=[
        'Seq',
        'Sample_ID',
        'Easting',
        'Northing',
        'Bathymetry',
        'SedThick',
        'Topas'
    ])
    base_name = input_dp.rsplit('/', 1)[1] # get dp name template
    base_name = base_name[:-5]    # remove extension
    dp['QR_Code'] = last_used + '_GSL_' + dp['Sample_ID']
    export(base_name + '_GSL.csv', dp)
    dp.drop(['Easting','Northing','Bathymetry','SedThick','Topas'], axis=1, inplace=True)
    dp['QR_Code'] = last_used + '_DSL_' + dp['Sample_ID']
    export(base_name + '_DSL.csv', dp)
    print('[DONE]')
    
print('DRILLING PROGRAM STAR\nDatasource creator')
if continue_prompt('\nUse concession name "' + last_used + '"? [Y/N] ') == False:
    last_used = input('Enter concession name (case sensitive): ')
print('Select Drilling Program xlsx')
input_dp = selectfile()
print('Select Output Directory')
output_dir = selectdir() + '/'
print('output dir received as:\n', output_dir)
folder = input_dp.rsplit('/', 1)[0]
config.set('last_used', 'concession', last_used)
config.set('last_used', 'input', folder)
config.set('last_used', 'output', output_dir)
with open('config.ini', 'w') as f:
    config.write(f)
append_fields()
print('Thank you for using DP Star')