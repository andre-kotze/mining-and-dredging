import os
import sys
import datetime as dt
import time
import random

import pandas as pd
import numpy as np
import xlsxwriter
from configparser import ConfigParser
from PyQt5 import QtCore, QtGui, QtWidgets

clear = lambda: os.system('cls')
GEOLOGY = '//192.168.0.5/Geology/'
outputdir = 'C:/Users/User/Desktop/DepositShit/'
key = 'Z:/Andre/KEY.txt'
traintxt = outputdir + 'Train.txt'
names = ['Beave', 'Bryornt', 'Bryon', 'Broyn', "Braaiin'", 'Bryne', 'Señor Bynor', 'Brandon', 'Brennard', 'Brorynyte', 'Byro-ryno-ryonantobony Byny-bon-rontytonbony', 'B', 'Bronynonorty Borytonbonby Rondonybobryntoborybon']
# last folder in config.ini not used

def read_sr(name, dep): 
    print('\nReading Sampling Results . . .')
    read_init = dt.datetime.now()
    #try:
    sr = pd.read_excel(name, skiprows=4, usecols=[
        'Deposit',
        'Sample N.\n&\nBarcode N.',
        'Individual  Stone Count Offshore',
        'IndividualCarats Offshore '
        ])
    sr.columns = np.arange(0,sr.shape[1])
    sr.rename({
        0:'Deposit',
        1:'Sample ID',
        2:'Stones',
        3:'Carats'}, 
        axis=1, inplace=True)
    sr.dropna(how='all', inplace=True)
    sr.reset_index(drop=True, inplace=True)
    sr.drop(0, inplace=True)  # drop blank row 6 in sheet
    sr.drop(sr[sr['Sample ID'].str.contains('Samples', na=False)].index, inplace=True)
    sr['Deposit'].fillna(method='ffill', inplace=True)
    sr = sr[sr['Deposit'].str.contains(dep, na=False)]
    sr.drop(sr[sr['Stones'] == 0].index, inplace=True)
    sr['Sample ID'].fillna(method='ffill', inplace=True)
    sr = sr.groupby(['Sample ID'], sort=False).sum()
    # 2 decimals 20210128
    sr['Carats'] = sr['Carats'].round(decimals=2)
    # now have a clean dataset per stone
    dur = dt.datetime.now() - read_init
    dur = dur.seconds + (dur.microseconds / 1000000)
    print('Sampling Results processed in', round(dur, 2), 'seconds (', len(sr), ' records)')
    return sr
    #except Exception as e:
    #    print(e, '\nError reading Sampling Results sheet')
    #    return None
with open(key, 'r') as k: 
    if k.read() != 'AKO': 
        sys.exit()
def export(output_data, depnr):
    name = 'DEP_' + depnr
    filename = outputdir + name + '.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    output_data.reset_index(inplace=True)
    output_data.to_excel(writer, index=False, sheet_name=name)
    # FORMATTING:
    wbook = writer.book
    wsheet = writer.sheets[name]
    num_format = wbook.add_format({'num_format': '0.00'})
    wsheet.set_column('A:A', 16)
    wsheet.set_column('C:C', 8, num_format)
    writer.save()
    print(f'Exported as {filename}')
    return filename

def get_dep_nr():
    depnr = input('\nEnter Deposit number: ')
    return depnr

def main():
    config = ConfigParser()
    config.read('config.ini')
    sr_filename = config.get('last_used', 'file')
    name = random.choice(names)
    print(f'Hello "{name}"')
    print('\n\nCurrent file:\t', sr_filename, '\n(enter any non-numeric value to select a different file)')
    depnr = get_dep_nr()
    while not depnr.isdigit():
        app = QtWidgets.QApplication(sys.argv)
        dialog = QtWidgets.QFileDialog()
        sr_filename = dialog.getOpenFileName(None, 'Open Sampling Results sheet', config.get('last_used', 'folder'))[0]
        config.set('last_used', 'file', sr_filename)
        with open('config.ini', 'w') as f:
            config.write(f)
        depnr = get_dep_nr()
    sr = read_sr(sr_filename, depnr)
    file = export(sr, depnr)
    os.startfile(file)
    print('[DONE]')


#clear()
main()