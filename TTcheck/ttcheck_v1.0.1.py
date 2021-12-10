import os
import sys
import time
from datetime import datetime, timedelta

import pandas as pd
import numpy
from configparser import ConfigParser
from PyQt5 import QtCore, QtGui, QtWidgets

# set constants
#EXPLORER_DATA = '//192.168.0.5/Explorer-DATA/Projects/2020/'
YDAY = (datetime.now().date() - timedelta(days = 1)).strftime("%Y-%m-%d")
#proj = '//192.168.0.5/Explorer-DATA/Projects/2020/04 - Project - BPT 3C - (2020.10.26 -/CR Logs/'
conc_paths = {
    '2C': 'Z:/2. BPT/2C/Screenlog_BPT127_2C.XLSB', 
    '3C': 'Z:/2. BPT/3C/Screenlog_BPT127_3C.XLSB'
    }

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

def get_conc():
    print('\nEnter concession manually')
    c = input('Concession name:').upper()
    return c

def read_cr_log(file): 
    checkpoint = datetime.now()
    cr = pd.read_excel(file, sheet_name='2 - SAMPLING Log', skiprows = 7, nrows = 3)
    cr = cr[[col for col in cr.columns if cr.loc[:, col].notna().all()]].transpose()
    cr = cr.iloc[1:]
    samplelist = cr.iloc[:, 1].values.tolist()
    cr_mapcamp = cr.iloc[:, 2].values.tolist()
    if not samplelist == cr_mapcamp:
        print("CR Log Map_ID & Campaign_ID don't match")
    clist = cr.iloc[:, 0].values.tolist()
    conc = set(clist)
    duration = datetime.now() - checkpoint
    duration = duration.seconds + (duration.microseconds / 1000000)
    print('(Read CR Log took', round(duration, 2), 'seconds)\n')
    try:
        [c] = conc
    except:
        if len(conc) == 0:
            print('Concession name could not be determined automatically (list empty)')
            c = get_conc()
        elif len(conc) > 1:
            print('Multiple concession names found in CR Log:\n', conc)
            c = get_conc()
    return c, samplelist

def read_sl(conc, filter_date):
    sl = conc_paths.get(conc)
    print('Reading screenlog...')
    screenlog = pd.read_excel(sl, engine = 'pyxlsb', usecols = [
        'Sample_ID', 
        'Start_Date',
    ])
    screenlog['Start_Date'] = pd.TimedeltaIndex(screenlog['Start_Date'], unit = 'd') + datetime(1899, 12, 30)
    screenlog = screenlog[screenlog['Start_Date'] == filter_date]
    #del screenlog['Start_Date']
    samplelist = screenlog['Sample_ID'].values.tolist()
    return samplelist 

def compare_CR_and_Slog(samplelist, samplelist2):
    if samplelist == samplelist2:
        print('Samples in CR log match those in Screenlog.')
        return True
    else:
        print('\nMismatch between CR log and Screenlog')
        print(len(samplelist), 'samples in CR log,', len(samplelist2), 'samples in Screenlog\n\nCHECK:')
        for s in samplelist:
            if s not in samplelist2:
                print(s, 'occurs in CR log, but not in Screenlog')
        for t in samplelist2:
            if t not in samplelist:
                print(t, 'occurs in Screenlog, but not in CR log')
        return False

def select_file(date):
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QFileDialog()
    cr_logfile = dialog.getOpenFileName(None, 'Open CR Log for ' + date, config.get('last_used', 'folder'))[0]
    config.set('last_used', 'file', cr_logfile)
    config.set('last_used', 'folder', cr_logfile.rsplit('/',1)[0])
    with open('config.ini', 'w') as f:
        config.write(f)
    return cr_logfile

def find_crlog(path, date):
    matches = []
    for file in os.listdir(path):
        if date in file:
            matches.append(file)
    if len(matches) == 0:
        print('No matching CR Log for', date, 'found in', path)
        if continue_prompt('Select CR Log manually? [Y/N] ') == True:
            return select_file(date)
        else:
            return None
    elif len(matches) == 1:
        return matches[0]
    else:
        print('Multiple files for', date, 'found in', path)
        if continue_prompt('Select CR Log manually? [Y/N] ') == True:
            return select_file(date)
        else:
            return None


def main():
    print('Run this script after 00:00 and before creating hole files from raw tooltrend data')
    print('It will compare the Sample IDs in the screenlog and the CR logs')
    if continue_prompt("use yesterday's date " + YDAY + " ? [Y/N] ") == True:
        filterdate = YDAY
    else:
        filterdate = input('\nEnter filter date YYYY-MM-DD:\n')
    config = ConfigParser()
    config.read('config.ini')
    cr_logfile = config.get('last_used', 'file')
    cr_logdir = config.get('last_used', 'folder')
    # READ CRLOG
    conc, cr_list = read_cr_log(cr_logfile)
    # READ SCREENLOG

    checkpoint = datetime.now()
    scrnlog_list = read_sl(conc, filterdate)
    duration = datetime.now() - checkpoint
    duration = duration.seconds + (duration.microseconds / 1000000)
    print('(Read Screenlog took', round(duration, 2), 'seconds)\n')
    # COMPARE SAMPLES
    if compare_CR_and_Slog(cr_list, scrnlog_list) == True:
        print('Done!')
    else:
        print("\n You're f*cked!, check CR and Screenlog sample IDs")

main()

# CR log 1st sample in cell D10