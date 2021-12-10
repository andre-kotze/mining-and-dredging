'''
PyQt5 GUI with buttons and options:
-show 'active' SL, option to change
-show 'active' CRL, option to change

-button to check:
    load logs
    compare SIDs, times
    alert for 00:00 samples
-log report in scrolledtext box

'''
import os
import sys
import datetime as dt
import time

import pandas as pd
import numpy as np
import openpyxl
from configparser import ConfigParser
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (
    QApplication, 
    QMainWindow, 
    QWidget, 
    QLabel, 
    QPushButton,
    QMessageBox,
    QFileDialog,
    QDesktopWidget,
    QVBoxLayout,
    QHBoxLayout,
    QDateEdit,
    QGridLayout,
    QToolTip,
    QTabWidget,
    QProgressDialog
)
from PyQt5.QtGui import QIcon, QPalette, QColor

__version__ = '1.0.0'
__author__ = 'AKO_Geo'

TODAY = dt.datetime.now().date().strftime('%Y%m%d')
YESTERDAY = (dt.datetime.now().date() - dt.timedelta(days = 1))#.strftime("%Y-%m-%d")
outputdir = os.getcwd()
config = ConfigParser()
config.read('config.ini')
savedest = config.get('last_used', 'savedest')
cpaths = 'concessions.csv'


# Main widget/UI and tab holder (and settings and about tab)
class GoingDeepUI(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.lddsl = config.get('last_used', 'screenlog')
        self.lddcrl = config.get('last_used', 'crlog')
        self.checkdir = config.get('last_used', 'checkdir')
        self.jf1, self.jf2 = '[none selected]', '[none selected]'
        self.sldf, self.crldf = None, None

        # Init tab screen
        self.tabs = QTabWidget() # self arg non-critical
        self.recontab = QWidget(self)
        self.checktab = QWidget(self)
        self.jointab = QWidget(self)
        #self.settab = QWidget() # self arg non-critical
        self.abouttab = QWidget()

        # Add tabs
        self.tabs.addTab(self.recontab, 'Recon')
        self.tabs.addTab(self.checktab, 'Check')
        self.tabs.addTab(self.jointab, 'Join')
        #self.tabs.addTab(self.settab, 'Settings')
        self.tabs.addTab(self.abouttab, 'About')

        self.abouttab.layout = QVBoxLayout(self)
        abouttext = QLabel(
            '''IMDH CR Log Recon\n\nVersion 1.0.0\n
            Written 2021 by Andre "Fucking Legend" "Drikus" Kotze\n
            email: andre@nammail.net'''
        )
        self.abouttab.layout.addWidget(abouttext)
        self.abouttab.layout.addStretch(1)
        self.abouttab.setLayout(self.abouttab.layout)

        self.recontab.layout = self.create_recon()
        self.recontab.setLayout(self.recontab.layout)

        self.checktab.layout = self.create_check()
        self.checktab.setLayout(self.checktab.layout)

        self.jointab.layout = self.create_join()
        self.jointab.setLayout(self.jointab.layout)

        '''
        self.settab.layout = self.create_savesettings()
        self.settab.setLayout(self.settab.layout)
        '''
        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def create_recon(self):
        # Loaded SL label
        self.sl_file = QLabel(os.path.split(self.lddsl)[1])
        self.sl_file.setToolTip(self.lddsl)
        self.slbbtn = QPushButton('...', self)
        self.slbbtn.clicked.connect(self.change_sl)
        self.slbbtn.setFixedSize(20, 20)
        self.readsl_btn = QPushButton('Import', self)
        self.readsl_btn.clicked.connect(lambda: read_sl(self.lddsl, self.datepick.date().toString("yyyy-MM-dd")))

        # Loaded CRL label
        self.crl_file = QLabel(os.path.split(self.lddcrl)[1])
        self.crl_file.setToolTip(self.lddcrl)
        self.crlbbtn = QPushButton('...', self)
        self.crlbbtn.clicked.connect(self.change_crl)
        self.crlbbtn.setFixedSize(20, 20)
        self.readcrl_btn = QPushButton('Import', self)
        self.readcrl_btn.clicked.connect(lambda: read_crl(self.lddcrl, self.datepick.date().toString("yyyy-MM-dd")))

        # box containing labels
        reconlayout = QGridLayout()
        reconlayout.setColumnStretch(1,1)

        reconlayout.addWidget(QLabel('Screenlog:'),0,0)
        reconlayout.addWidget(self.sl_file,0,1)
        reconlayout.addWidget(self.slbbtn,0,2)
        reconlayout.addWidget(self.readsl_btn,0,3)

        reconlayout.addWidget(QLabel('CR log:'),1,0)
        reconlayout.addWidget(self.crl_file,1,1)
        reconlayout.addWidget(self.crlbbtn,1,2)
        reconlayout.addWidget(self.readcrl_btn,1,3)

        self.datepick = QDateEdit(calendarPopup=True)
        self.datepick.setDate(YESTERDAY)
        self.reconbtn = QPushButton('RECONCILE', self)
        self.reconbtn.clicked.connect(self.reconcile)
        reconlayout.addWidget(QLabel('Filter date:'),2,0)
        reconlayout.addWidget(self.datepick,2,1)
        reconlayout.addWidget(self.reconbtn,2,2,1,2)

        return reconlayout

    def create_check(self):
        self.checkdirlbl = QLabel(self.checkdir)
        self.truncate_label(self.checkdir, self.checkdirlbl)
        self.checkdirbtn = QPushButton('...')
        self.checkdirbtn.clicked.connect(self.select_dir)
        self.checkdirbtn.setFixedSize(20, 20)
        self.checkfiles = QPushButton('CHECK')
        self.checkfiles.clicked.connect(self.check_files)
        
        checklayout = QGridLayout()
        checklayout.setColumnStretch(1,1)
        checklayout.addWidget(QLabel('Folder:'),0,0)
        checklayout.addWidget(self.checkdirlbl,0,1)
        checklayout.addWidget(self.checkdirbtn,0,2)
        checklayout.addWidget(self.checkfiles,1,1)
        
        return checklayout
        
    def create_join(self):
        # File 1 label
        self.file1lbl = QLabel(self.jf1, self)
        self.file1btn = QPushButton('...', self)
        self.file1btn.clicked.connect(self.select_file1)
        self.file1btn.setFixedSize(20, 20)

        # File 2 label
        self.file2lbl = QLabel(self.jf2, self)
        self.file2btn = QPushButton('...', self)
        self.file2btn.clicked.connect(self.select_file2)
        self.file2btn.setFixedSize(20, 20)

        self.joinbtn = QPushButton('JOIN')
        self.joinbtn.clicked.connect(self.join_files)

        joinlayout = QGridLayout()
        joinlayout.setColumnStretch(1,1)

        joinlayout.addWidget(QLabel('File 1:'),0,0)
        joinlayout.addWidget(self.file1lbl,0,1)
        joinlayout.addWidget(self.file1btn,0,2)
        joinlayout.addWidget(QLabel('File 2:'),1,0)
        joinlayout.addWidget(self.file2lbl,1,1)
        joinlayout.addWidget(self.file2btn,1,2)
        #joinlayout.addWidget(self.savedestlbl,2,1)
        #joinlayout.addWidget(self.savedestbtn,2,2)
        joinlayout.addWidget(self.joinbtn,2,1)

        return joinlayout

    # function to truncate long paths and add tooltip (DIRS ONLY)
    def truncate_label(self, path, labelobj): # path = filepath string, labelobj = QLabel
        if len(path) > 50:
            trunclbl = '...' + path[-50:]
            labelobj.setText(trunclbl)
            labelobj.setToolTip(path)
        else:
            labelobj.setText(path)
            labelobj.setToolTip(None)

    def change_sl(self):
        file = QFileDialog().getOpenFileName(None, 'Select Geo Screenlog xlsb', os.path.split(self.lddsl)[0])[0]
        if len(file) != 0:
            self.lddsl = file
            self.sl_file.setText(os.path.split(self.lddsl)[1])
            self.sl_file.setToolTip(self.lddsl)
            config.set('last_used', 'screenlog', self.lddsl)
            with open('config.ini', 'w') as f:
                config.write(f)

    def change_crl(self):
        file = QFileDialog().getOpenFileName(None, 'Select CR log xlsm', os.path.split(self.lddcrl)[0])[0]
        if len(file) != 0:
            self.lddcrl = file
            self.crl_file.setText(os.path.split(self.lddcrl)[1])
            self.crl_file.setToolTip(self.lddcrl)
            config.set('last_used', 'crlog', self.lddcrl)
            with open('config.ini', 'w') as f:
                config.write(f)


    # for choosing HDdir (check):
    def select_dir(self): 
        dialog = QtWidgets.QFileDialog()
        dirname = dialog.getExistingDirectory(None, 'Select folder', config.get('last_used', 'checkdir'))
        if len(dirname) != 0:
            dirname += '/'
            self.checkdir = dirname
            config.set('last_used', 'checkdir', dirname)
            with open('config.ini', 'w') as f:
                config.write(f)
            print('checkdir changed to', self.checkdir)
            self.truncate_label(dirname, self.checkdirlbl)

    # for choosing holedata files
    def select_file1(self):
        filename = QFileDialog().getOpenFileName(None, 'Select Hole Data xlsx', None, 'Excel workbook (*.xlsx)')[0]
        if len(filename) != 0:
            self.jf1 = filename
            self.file1lbl.setText(os.path.basename(self.jf1))
            self.file1lbl.setToolTip(self.jf1)
        else:
            print('select file cancelled')

    def select_file2(self):
        filename = QFileDialog().getOpenFileName(None, 'Select Hole Data xlsx', None, 'Excel workbook (*.xlsx)')[0]
        if len(filename) != 0:
            self.jf2 = filename
            self.file2lbl.setText(os.path.basename(self.jf2))
            self.file2lbl.setToolTip(self.jf2)
        else:
            print('select file cancelled')
    '''
    def import_sl(self):
        busy = QMessageBox(QMessageBox.NoIcon, 'Loading', 'Reading Screenlog')
        busy.show()
        filterdate = self.datepick.date().toString("yyyy-MM-dd")
        self.sldf = read_sl(self.lddsl, filterdate)
        print(f'import complete, received {len(self.sldf)} records')
        if self.sldf is not None:
            self.readsl_btn.setText('Imported')
        busy.hide()
            
    def import_crl(self):
        busy = QMessageBox(QMessageBox.NoIcon, 'Loading', 'Reading CR log')
        busy.show()
        self.crldf = read_crl(self.lddcrl)
        print(f'import complete, received {len(self.crldf)} records')
        if self.crldf is not None:
            self.readcrl_btn.setText('Imported')
        busy.hide()
    '''
    def reconcile(self):
        print(f'\nreconciling: {os.path.basename(self.lddsl)} and {os.path.basename(self.lddcrl)}')
        filterdate = self.datepick.date().toString("yyyy-MM-dd")
        print(f'using date: {filterdate}')
        self.sldf = read_sl(self.lddsl, filterdate)
        self.crldf = read_crl(self.lddcrl, filterdate)
        
        compare_tables(self.sldf, self.crldf)

    def check_files(self):
        print(f'checking files in {self.checkdir}')
        if not os.path.exists(self.checkdir):
            print('Unreachable directory:\n' + self.checkdir)
        else:
            hdcheck(self.checkdir)
    
    def join_files(self):
        print(f'joining {os.path.basename(self.jf1)} and {os.path.basename(self.jf2)}')

def join(f1, f2):
    return None
    '''
    msg = QMessageBox()
    msg.setText(report)
    msg.setWindowTitle(msgtitle)
    msg.setWindowIcon(QIcon('icon.ico'))
    if folder != None:
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Open)
        msg.button(QMessageBox.Open).setText('Open Folder')
        ret = msg.exec()
        if ret == QMessageBox.Open:
            os.startfile(folder)
    else:
        msg.exec()
    '''

def read_crl(filename, filterdate):
    print('reading cr_log')
    try:
        read_init = dt.datetime.now()
        #y, m, d = filterdate.split('-')
        #y, m, d = int(y), int(m), int(d)
        crl = pd.read_excel(filename, sheet_name='Time Sheet', usecols=['HOLE', 'SAMPLE ID', 'START', 'STOP'], index_col='HOLE')
        #crl.Start_Datetime = pd.TimedeltaIndex(crl.START, unit = 'd')
        #crl.STOP = pd.to_datetime(crl.STOP, unit = 'd')
        crl.dropna(subset=['SAMPLE ID'], inplace=True)
        '''
        starts, stops = [], []
        for s in crl.itertuples(index=False):
            h, m, s = str(s.START).split(':')
            starts.append(pd.Timestamp(y,m,d,int(h),int(m),int(s)))
            h, m, s = str(s.STOP).split(':')
            stops.append(pd.Timestamp(y,m,d,int(h),int(m),int(s)))
        crl.Start_Datetime = starts
        crl.End_Datetime = stops
        '''
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('CR log processed in', round(dur, 2), 'seconds (', len(crl), ' records)')
    except Exception as e:
        print(e, '\nError reading CR log')
        crl = None
    return crl

def read_sl(filename, filterdate): 
    size = os.stat(filename).st_size / 2**20
    dur = round(size*(0.1*size + 1.48),1) # seconds
    if dur > 4:
        slreadmsg = QMessageBox(QMessageBox.NoIcon, 'Reading Screenlog', f'Estimated time {dur} seconds')
        slreadmsg.show()
    print('\nReading Screenlog . . .')
    try:
        read_init = dt.datetime.now()
        sl = pd.read_excel(filename, engine='pyxlsb', usecols=['Sample_ID', 'Start_Date', 'End_Date', 'Start_Time', 'End_Time'])
        #parse dates, times
        epoch = dt.datetime(1899, 12, 30)
        sl['Start_Date'] = pd.TimedeltaIndex(sl['Start_Date'], unit = 'd') + epoch
        sl['End_Date'] = pd.TimedeltaIndex(sl['End_Date'], unit = 'd') + epoch
        sl['Start_Date'] = sl['Start_Date'].dt.strftime('%Y-%m-%d')
        sl['End_Date'] = sl['End_Date'].dt.strftime('%Y-%m-%d')
        sl['Start_Time'] = pd.to_datetime(sl['Start_Time'], unit = 'd')
        sl['End_Time'] = pd.to_datetime(sl['End_Time'], unit = 'd')
        sl = sl[sl.Start_Date == filterdate]
        #sl.Start_Datetime = pd.TimedeltaIndex(sl.Start_Date + sl.Start_Time, unit = 'd')
        #sl.End_Datetime = pd.TimedeltaIndex(sl.End_Date + sl.End_Time, unit = 'd')
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('Screenlog processed in', round(dur, 2), 'seconds (', len(sl), ' records)')
    except Exception as e:
        print(e, '\nError reading Screenlog')
        sl = None
    try:
        slreadmsg.hide()
    except NameError:
        pass
    return sl

def compare_tables(sldf, crldf):
    # compare SIDs and start/end times. check for 00:00 holes
    if type(sldf) == None or type(crldf) == None:
        print('empty df(s) received')
        return
    print(f'Comparing ({len(sldf)} records from SL, {len(crldf)} records from CR log')
    slids = sldf['Sample_ID'].tolist()
    crlids = crldf['SAMPLE ID'].tolist()
    if slids != crlids:
        elist = ['SAMPLE_ID ERRORS']
        elist.append('Only found in Screenlog:')
        for slid in slids:
            if slid not in crlids:
                elist.append(slid)
        elist.append('\nOnly found in CR log:')
        for crlid in crlids:
            if crlid not in slids:
                elist.append(crlid)
        report = QMessageBox(QMessageBox.NoIcon, 'Check completed', f'{len(elist)-3} mismatches')
        report.setDetailedText('\n'.join(elist))
        report.setWindowIcon(QIcon('icon.ico'))
    else:
        report = QMessageBox(QMessageBox.NoIcon, 'Check completed', 'Sample IDs match')
        report.setWindowIcon(QIcon('icon.ico'))
    report.exec()
    '''
    # now times
    starts = crldf.START.tolist()
    stops = crldf.STOP.tolist()
    times, elist = [], []
    if len(starts) != len(stops):
        print('Times count mismatch!')
    for n in range(len(starts)):
        times.append(starts[n])
        times.append(stops[n])
    for n in range(len(times)):
        try:
            if not times[n+1] > times[n]:
                elist.append(f'Time error: {times[n+1]} vs {times[n]}')
        except IndexError:
            pass
    if len(elist) != 0:
        elist.insert(0,'CR LOG TIME ERRORS')
        QMessageBox(QMessageBox.NoIcon, 'Check completed', f'{len(elist)-1} time errors')
        report.setDetailedText('\n'.join(elist))
        report.setWindowIcon(QIcon('icon.ico'))
    report.exec()

    # SL times vs CRL times:
    elist = []
    geostarts = sldf.Start_Time.tolist()
    geoends = sldf.End_Time.tolist()
    lee = dt.timedelta(minutes=5)
    for j in zip(geostarts, starts, geoends, stops, slids):
        if abs(j[0] - j[1]) > lee or abs(j[2] - j[3]) > lee:
            elist.append(j[4])
    if len(elist) != 0:
        elist.insert(0,'CR vs SL TIME MISMATCHES')
        QMessageBox(QMessageBox.NoIcon, 'Check completed', f'{len(elist)-1} time mismatches')
        report.setInformativeText('Times in CR log and screenlog differ by more than 5 minutes')
        report.setDetailedText('\n'.join(elist))
        report.setWindowIcon(QIcon('icon.ico'))
    report.exec()
    '''



def check_trend(dir, f):
    try:
        ttdata = pd.read_excel(dir + f, skiprows=1, usecols=['Drill Depth\nmm'])
        depthslist = ttdata['Drill Depth\nmm'].values.tolist()
        sid = f[:-5]
        if depthslist[0] > 0 and depthslist[-1] > 0:
            return f'{sid}: SOH and EOH data missing'
        elif depthslist[-1] > 0:
            return f'{sid}: EOH data missing'
        elif depthslist[0] > 0:
            return f'{sid}: SOH data missing'
        else:
            return None
    except Exception as e:
        print(e)
        return f'{sid}: Unreadable file'

def hdcheck(dir):
    tlist, elist = [], []
    for file in os.listdir(dir):
        if '~$' in file:
            pass
        elif '.xlsx' in file:
            tlist.append(file)
    progress = QProgressDialog('Checking files...', 'Cancel', 0, len(tlist))
    progress.setWindowTitle('Hole Data Integrity Check')
    progress.setWindowIcon(QIcon('icon.ico'))
    progress.setWindowModality(QtCore.Qt.WindowModal)
    for i, file in enumerate(tlist):
        status = check_trend(dir, file)
        if status != None:
            elist.append(status)
        progress.setValue(i+1)
        if progress.wasCanceled():
            break
    progress.setValue(len(tlist))
    if len(elist) == 0:
        report = QMessageBox(QMessageBox.NoIcon, 'Check completed', 'No incomplete files found')
        report.setWindowIcon(QIcon('icon.ico'))
    else:
        details = ''
        for e in elist:
            details += e + '\n'
        report = QMessageBox(QMessageBox.NoIcon, 'Check completed', f'{len(elist)} incomplete files found')
        report.setDetailedText(details)
        report.setWindowIcon(QIcon('icon.ico'))
    report.exec()


# MAIN FUNCTION
def main():
    # Create an instance of QApplication
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.WindowText, QColor(255, 255, 255))
    palette.setColor(QPalette.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ToolTipBase, QColor(0, 0, 0))
    palette.setColor(QPalette.ToolTipText, QColor(255, 255, 255))
    palette.setColor(QPalette.Text, QColor(255, 255, 255))
    palette.setColor(QPalette.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ButtonText, QColor(255, 255, 255))
    palette.setColor(QPalette.BrightText, QColor(0, 128, 255))
    palette.setColor(QPalette.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))
    app.setPalette(palette)
    # Show the GUI
    godeep = QMainWindow()
    godeep.setWindowTitle('IMDH CR Log check v1')
    godeep.setWindowIcon(QIcon('icon.ico'))
    godeep.resize(360, 220)
    godeep.tabwidget = GoingDeepUI(godeep)
    godeep.setCentralWidget(godeep.tabwidget)
    godeep.show()
    # main loop
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()