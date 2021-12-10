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
import datetime as dt

import pandas as pd
import numpy as np
from configparser import ConfigParser
from PyQt5 import QtCore, QtGui, QtWidgets

lists = pd.read_csv('lists.csv')
shapes = lists['Shape'].values.tolist()
colours = lists['Colour'].values.tolist()
clarities = lists['Clarity'].values.tolist()

log =[]
outputdir = 'C:/Users/User/Desktop/'

def check(msg="'N' to cancel. Any other key to continue: "):
    if input(msg).upper() == 'N':
        return False
    else:
        return True

def prompt(msg='Continue? [Y/N] '):
    proceed = ''
    while proceed != 'Y':
        proceed = input(msg)
        if proceed.upper() == 'N':
            return False
        elif proceed.upper() == 'Y':
            return True
        else:
            print('Epic fail! Please respond with Y or N, not "' + proceed + '"')
            time.sleep(1.2)

def select_sl(last):
        app = QtWidgets.QApplication(sys.argv)
        dialog = QtWidgets.QFileDialog()
        filename = dialog.getOpenFileName(None, 'Open Screenlog', last)[0]
        return filename

def read_sl(filename, seq):
    print('\nReading Screenlog . . .')
    read_init = dt.datetime.now()
    #try:
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
    end = sl['Seq'].tolist()[-1] # last seq
    sl.set_index('Seq', inplace=True)
    #sid, isco, cts, shp, col, cla = [], [], [], [], [], []
    data = []
    for i, s in sl.iterrows(): # i = SeqNr
        if s['Stones'] == 0:
            data.append((i, s['Sample_ID'], 0, 0, '', '', '')) # 7 cols, seq for reference
        elif s['Stones'] == 1:
            if round(s['Est_Carats'], 2) != round(float(s['Est_Brkdwn']), 2): print('Check carats:', i)
            data.append((i, s['Sample_ID'], s['Stones'], s['Est_Carats'], s['Shape'], s['Colour'], s['Clarity']))
        else:
            cts = s['Est_Brkdwn'].split(', ')
            shp = s['Shape'].split(', ')
            col = s['Colour'].split(', ')
            cla = s['Clarity'].split(', ')
            data.append((i, s['Sample_ID'], 1, cts[0], shp[0], col[0], cla[0])) # for stone 1
            for d in range(1,int(s['Stones'])): # for stones 2 to n
                data.append(('','', 1, cts[d], shp[d], col[d], cla[d]))
    sbs = pd.DataFrame.from_records(data, columns=['Seq', 'Sample_ID', 'ISCO', 'Carats', 'Shape', 'Colour', 'Clarity'])
    # VALIDATION:
    for seq, row in sbs.iterrows():
        if (
            row['Shape'] not in shapes) or (
            row['Colour'] not in colours) or (
            row['Clarity'] not in clarities
            ):
            #print(f'CHECK {seq} stone properties')

    dur = dt.datetime.now() - read_init
    dur = dur.seconds + (dur.microseconds / 1000000)
    print('Screenlog processed in', round(dur, 2), 'seconds (', len(sbs), ' records)')
    return sbs, end
    #except Exception as e:
    #    print(e, '\nError reading Screenlog')
    #    return None

def export(output_data, startSeq, endSeq):
    name = 'SR_Seq_' + str(startSeq) + '-' + str(endSeq)
    filename = outputdir + name + '.csv'
    output_data.to_csv(filename, index=False)
    #writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    #output_data.reset_index(inplace=True)
    #output_data.to_excel(writer, index=False, sheet_name=name)
    # FORMATTING:
    #wbook = writer.book
    #wsheet = writer.sheets[name]
    #num_format = wbook.add_format({'num_format': '0.00'})
    #wsheet.set_column('A:A', 16)
    #wsheet.set_column('C:C', 8, num_format)
    #writer.save()
    print(f'Exported as {filename}')
    return filename

def main():
    print('\nTIFL Auto Screenlog Samplings ("ASS")')
    config = ConfigParser()
    config.read('config.ini')
    lastdir = config.get('last_used', 'folder')
    filename = config.get('last_used', 'file')
    seq = config.get('last_used', 'seq')
    # confirm Screenlog in use:
    print('Good Day BRYORNT\n')
    print('Screenlog:', filename.rsplit('/',1)[1])
    if check("'N' to select a different file. Any other key to continue: ") == False:
        filename = select_sl(lastdir)
        config.set('last_used', 'file', filename)
        config.set('last_used', 'folder', filename.rsplit('/',1)[1])
        with open('config.ini', 'w') as f:
            config.write(f)

    # input starting Seq (remembering last Seq from last run)
    print('Starting Sequence Nr:', seq)
    if check("'N' to enter a different Seq Nr. Any other key to continue: ") == False:
        seq = input('Sequence Nr to start at: ')
        config.set('last_used', 'seq', seq)
        with open('config.ini', 'w') as f:
            config.write(f)
    # read screenlog, convert to stone-by-stone
    sbs, end = read_sl(filename, int(seq))
    file = export(sbs, seq, end)
    os.startfile(file)
    print('[DONE]')



'''
    viewlog = input("\nPress Enter to exit, or L to view the log\n")
    if viewlog.upper() == "L":
        print('')
        for i in log:
            print('\t', i)
        print('')
        exit = input("\nPress Enter to exit")
    else:
        sys.exit(0)
'''

main()