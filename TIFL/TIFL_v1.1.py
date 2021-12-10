'''
Typing Is For Losers version 1.0.0:

1. Read recovery data from screenlog
2. Validate
3. Output in SamplingResults format
    AS CSV
'''
import os
import sys
import time
import random
import datetime as dt

import pandas as pd
import numpy as np
from configparser import ConfigParser
from PyQt5 import QtCore, QtGui, QtWidgets

shapes = [1, '1', 2, '2', 3, '3', 4, '4', 5, '5', 6, '6', 7, '7']
colours = ['WHT', 'YLW', 'BRN', 'GRN', 'GRY', 'RED', 'BLU', 'PNK', 'PPL', 'VIL', 'ORN', 'BLK', 'PUR', 'OWH']
clarities = ['CLR', 'INC', 'BRS', 'FRA', 'CAV', 'FRO']
names = ['Beave', 'Bryornt', 'Bryon', 'Broyn', "Braaiin'", 'Señor Bynor', 'Brandon', 'Brennard', 'Byro-ryno-ryonantobony Byny-bon-rontytonbony', 'Bronynonorty Borytonbonby Rondonybobryntoborybon']

log =[]
outputdir = 'C:/Users/User/Desktop/SamplingResults/'

def check(msg="'N' to change. Any other key to continue: "): # accept seq nr implicitly
    res = input(msg)
    if len(res) == 0:
        return True
    elif res.upper() == 'N':
        return False
    try:
        seq = int(res)
        return seq
    except:
        return True
    else:
        return True

with open('Z:/Andre/KEY.txt', 'r') as k:
    if k.read() != 'AKO':
        sys.exit()

def select_sl(last):
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QFileDialog()
    filename = dialog.getOpenFileName(None, 'Open Screenlog', last)[0]
    return filename

def read_sl(filename, seq):
    print('\nReading Screenlog . . .')
    read_init = dt.datetime.now()
    try:
        sl = pd.read_excel(filename, engine='pyxlsb', usecols=[
            'Seq',
            'Sample_ID',
            'Stones',
            'Group_wt',
            'Est_Carats',
            'Est_Brkdwn',
            'Shape',
            'Colour',
            'Clarity'
            ])
        sl.dropna(how='all', subset=['Stones'], inplace=True) # drop trailing empties
        sl = sl[sl['Seq'] >= seq] # trim olds
        if sl.shape[0] == 0:
            print('[No samples]')
            sys.exit(0)
        end = sl['Seq'].tolist()[-1] # last seq
        sl.set_index('Seq', inplace=True)
        data = []
        for i, s in sl.iterrows(): # i = SeqNr
            error = []
            if s['Stones'] == 0:
                if not (s['Est_Carats'] == 0) and (len(s['Shape'] + s['Colour'] + s['Clarity']) == 0):
                    error.append('sample')
                data.append((i, s['Sample_ID'], 0, 0, '', '', '')) # 7 cols, seq for reference
            elif s['Stones'] == 1:
                if round(s['Est_Carats'], 2) != round(float(s['Est_Brkdwn']), 2):
                    error.append('carats&brkdwn')
                data.append((i, s['Sample_ID'], s['Stones'], s['Est_Carats'], s['Shape'], s['Colour'], s['Clarity']))
                if s['Shape'] not in shapes:
                    error.append('shape')
                if s['Colour'] not in colours:
                    error.append('colour')
                for sc in s['Clarity'].split('/'):
                    if sc not in clarities:
                        error.append('clarity')
            else:
                cts = s['Est_Brkdwn'].split(', ')
                total = 0
                for t in cts:
                    total += float(t)
                if round(total, 2) != round(float(s['Est_Carats']), 2):
                    error.append('Est_Carats vs Est_Brkdwn')
                shp = s['Shape'].split(', ')
                col = s['Colour'].split(', ')
                cla = s['Clarity'].split(', ')
                data.append((i, s['Sample_ID'], 1, cts[0], shp[0], col[0], cla[0])) # for stone 1
                for d in range(1,int(s['Stones'])): # for stones 2 to n
                    try:
                        data.append(('','', 1, cts[d], shp[d], col[d], cla[d]))
                    except IndexError as e:
                        print(f'Fatal Error: Check Seq {i} Stone {d} properties/format and retry')
                        sys.exit(0)
                    if shp[d] not in shapes:
                        error.append('shape')
                    if col[d] not in colours:
                        error.append('colour')
                    for sc in cla[d].split('/'):
                        if sc not in clarities:
                            error.append('clarity')
            if len(error) > 0:
                errors = ' '.join(error)
                print(f'CHECK Seq {i} {errors}')
        sbs = pd.DataFrame.from_records(data, columns=['Seq', 'Sample_ID', 'ISCO', 'Carats', 'Shape', 'Colour', 'Clarity'])
        
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('Screenlog processed in', round(dur, 2), 'seconds (', len(sbs) - 1, ' samples)')
        return sbs, end
    except Exception as e:
        print(e, '\n\n[Error reading Screenlog]\nContact your AKO')
        sys.exit(0)

def export(output_data, startSeq, endSeq):
    name = 'SR_Seq_' + str(startSeq) + '-' + str(endSeq)
    filename = outputdir + name + '.csv'
    output_data.to_csv(filename, index=False)
    print(f'Exported as {filename}')
    return filename

def main():
    print('Welcome ', end='', flush=True)
    time.sleep(0.6)
    name = random.choice(names)
    print(name)
    time.sleep(0.8)
    print('\nT.I.F.L.: Typing Is For Losers\na.k.a Typing Is Too Slow ("T.I.T.S.")\n\n')
    time.sleep(0.8)
    print('Auto Screenlog Samplings ("A.S.S.")')
    config = ConfigParser()
    config.read('config.ini')
    lastdir = config.get('last_used', 'folder')
    if lastdir == None:
        lastdir = outputdir
    filename = config.get('last_used', 'file')
    seq = config.get('last_used', 'seq')
    # confirm Screenlog in use:
    if len(filename) == 0:
        filename = select_sl(lastdir)
        config.set('last_used', 'file', filename)
        config.set('last_used', 'folder', filename.rsplit('/',1)[0])
        with open('config.ini', 'w') as f:
            config.write(f)
    print('Screenlog in use:', filename.rsplit('/',1)[1])
    if check() == False: # "'N' to select a different file. Any other key to continue: "
        filename = select_sl(lastdir)
        config.set('last_used', 'file', filename)
        config.set('last_used', 'folder', filename.rsplit('/',1)[0])
        with open('config.ini', 'w') as f:
            config.write(f)
    # input starting Seq (remembering last Seq from last run)
    print('\nStarting Sequence Nr:', seq)
    res = check("Enter to continue, or enter a different Sequence Nr to change: ")
    if res == False or type(res) == int:
        while type(res) != int:
            res = check('Enter Starting Sequence Nr: ')
        seq = res
        config.set('last_used', 'seq', str(seq))
        with open('config.ini', 'w') as f:
            config.write(f)
    # read screenlog, convert to stone-by-stone
    sbs, end = read_sl(filename, int(seq))
    file = export(sbs, seq, end)
    new_seq = str(end + 1)
    config.set('last_used', 'seq', new_seq)
    with open('config.ini', 'w') as f:
        config.write(f)
    os.startfile(file)
    print('[DONE]')

main()