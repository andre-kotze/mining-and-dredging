'''
version DOS con GUI
read diamond data from 3 sources and reconcile:
1. Diamond Record Sheet             ONE_LINE_PER_SAMPLE
    sheet_name = 3C_DiamondSortingLogs
    nr, size, colour, clarity
    2 separate column headings :o
    date_format = 2020-11-24

2. Screenlog                        ONE_LINE_PER_SAMPLE
    first sheet
    nr, size, colour, clarity
    date_format = 2020-11-24

USE SCREENLOG AS BASE DATASET (complete list of all sampling)

3. SamplingResults                  ONE_LINE_PER_STONE
    first sheet
    nr, size, colour, clarity
    date_format = 2020/11/24 as string

Filter/trim by DATE (DRILL START)

compare lists:
    per sample? (sl, drs)
    per stone?  (sr)

    compare DRS with SL? Then SL with SR

    for example Nov 24-25:
        SL:     50
        DRS:    50, 2 ADT
        SR:     50, 2 ADT
'''
import os
import sys
import time
import datetime as dt

import pandas as pd
import numpy as np

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
    QRadioButton,
    QButtonGroup,
    QToolTip,
    QTabWidget,
    QGroupBox,
    QSpacerItem,
    QGridLayout,
    QDateEdit,
    QTextBrowser,
    QDialog
)
from PyQt5.QtGui import QIcon

__version__ = '2.0.0'
__author__ = 'AKO_Geo'

config = ConfigParser()
config.read('config.ini')
EPOCH = dt.datetime(1899, 12, 30)
lists = pd.read_csv('lists.csv')
shapes = lists['Shape'].values.tolist()
colours = lists['Colour'].values.tolist()
clarities = lists['Clarity'].values.tolist()

props = {0:'Carat', 1:'Shape', 2:'Colour', 3:'Clarity'}

class Las_Preguntas(QMessageBox):
    def __init__(self, title='Continue', msg='¿Sí o no?'):
        super(QMessageBox, self).__init__()
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon('icon.ico'))
        self.setText(msg)
        self.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        self.button(QMessageBox.Yes).setText('Sí')
        self.button(QMessageBox.No).setText('No')

class Los_Errors(QMessageBox):
    def __init__(self, title='ERROR', error='Error occurred', reportlist=None):
        super(QMessageBox, self).__init__()
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon('icon.ico'))
        self.setText(error)
        if reportlist != None:
            report = '\n'.join(reportlist)
            self.setDetailedText(report)
            #self.setStyleSheet("QTextEdit{min-width: 480px; min-height: 666px}")
        self.exec()
    

    
class El_Informe(QDialog):
    def __init__(self, title='REPORT', text='Recon completed', reportlist=None):
        super(QWidget, self).__init__()
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon('icon.ico'))
        self.resize(800, 600)
        
        self.report = QTextBrowser()
        #self.report.setHtml(True)
        self.report.setOpenExternalLinks(True)
        self.report.anchorClicked.connect(self.open_log)
        self.ok = QPushButton('OK')
        self.ok.clicked.connect(self.close)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        
        layout = QVBoxLayout()
        layout.addWidget(self.report)
        footer = QHBoxLayout()
        footer.addStretch(1)
        footer.addWidget(self.ok)
        layout.addLayout(footer)
        #self.ok.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.setLayout(layout)

        for ln in reportlist:
            self.report.append(ln)

    def open_log(self, log):
        address = str(log.toString())
        self.report.setSource(QtCore.QUrl())
        if os.path.exists(address):
            os.startfile(address)
        else:
            Los_Errors('404', 'File not found')


class El_Recon_Dos(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__()
        self.act_sl = config.get('last_used', 'sl')
        self.act_drs = config.get('last_used', 'drs')
        self.act_sr = config.get('last_used', 'sr')
        self.drsdf, self.sldf, self.srdf = None, None, None
        self.last_start = dt.datetime.strptime(config.get('last_used', 'start'),'%Y-%m-%d')
        self.last_end = dt.datetime.strptime(config.get('last_used', 'end'), '%Y-%m-%d')

        mainlayout = QVBoxLayout()
        self.fileslayout = self.files()
        self.dateslayout = self.dates()
        self.buttonslayout = self.buttons()
        mainlayout.addLayout(self.fileslayout)
        mainlayout.addLayout(self.dateslayout)
        mainlayout.addLayout(self.buttonslayout)
        self.setLayout(mainlayout)

    def files(self):
        # Loaded SL label
        self.sl_file = QLabel(os.path.basename(self.act_sl))
        self.sl_file.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.sl_file.setToolTip(self.act_sl)
        self.slbbtn = QPushButton('...', self)
        self.slbbtn.clicked.connect(self.change_sl)
        self.slbbtn.setFixedSize(20, 20)
        self.readsl_btn = QPushButton('Load', self)
        self.readsl_btn.clicked.connect(self.import_sl)

        # Loaded DRS label
        self.drs_file = QLabel(os.path.basename(self.act_drs))
        self.drs_file.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.drs_file.setToolTip(self.act_drs)
        self.drsbbtn = QPushButton('...', self)
        self.drsbbtn.clicked.connect(self.change_drs)
        self.drsbbtn.setFixedSize(20, 20)
        self.readdrs_btn = QPushButton('Load', self)
        self.readdrs_btn.clicked.connect(self.import_drs)

        # Loaded SR label
        self.sr_file = QLabel(os.path.basename(self.act_sr))
        self.sr_file.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.sr_file.setToolTip(self.act_sr)
        self.srbbtn = QPushButton('...', self)
        self.srbbtn.clicked.connect(self.change_sr)
        self.srbbtn.setFixedSize(20, 20)
        self.readsr_btn = QPushButton('Load', self)
        self.readsr_btn.clicked.connect(self.import_sr)

        flay = QGridLayout()
        flay.setColumnStretch(1,1)

        flay.addWidget(QLabel('Screenlog:'),0,0)
        flay.addWidget(self.sl_file,0,1)
        flay.addWidget(self.slbbtn,0,2)
        flay.addWidget(self.readsl_btn,0,3)

        flay.addWidget(QLabel('Diamond Record Sheet:'),2,0)
        flay.addWidget(self.drs_file,2,1)
        flay.addWidget(self.drsbbtn,2,2)
        flay.addWidget(self.readdrs_btn,2,3)

        flay.addWidget(QLabel('Sampling Results:'),1,0)
        flay.addWidget(self.sr_file,1,1)
        flay.addWidget(self.srbbtn,1,2)
        flay.addWidget(self.readsr_btn,1,3)

        return flay

    def dates(self):
        self.start_date = QDateEdit(calendarPopup=True)
        self.start_date.setStyleSheet("QDateEdit{min-width: 90px}")
        self.start_date.setDate(self.last_start)
        self.end_date = QDateEdit(calendarPopup=True)
        self.end_date.setStyleSheet("QDateEdit{min-width: 90px}")
        self.end_date.setDate(self.last_end)
        dlay = QHBoxLayout()
        dlay.addWidget(QLabel('Date range:'))
        dlay.addStretch(1)
        dlay.addWidget(self.start_date)
        dlay.addWidget(QLabel(' to '))
        dlay.addWidget(self.end_date)

        return dlay

    def buttons(self):
        blay = QHBoxLayout()
        self.drs_vs_sl_btn = QPushButton('DRS vs SL', self)
        self.sl_vs_sr_btn = QPushButton('SL vs SR', self)
        self.drs_vs_sl_btn.setMaximumWidth(150)
        self.sl_vs_sr_btn.setMaximumWidth(150)
        self.drs_vs_sl_btn.clicked.connect(self.recon_drs_sl)
        blay.addWidget(self.drs_vs_sl_btn)
        blay.addWidget(self.sl_vs_sr_btn)
        return blay

    def change_sl(self):
        file = QFileDialog().getOpenFileName(None, 'Select Screenlog xlsb', os.path.dirname(self.act_sl))[0]
        if len(file) != 0 and file != self.act_sl:
            self.act_sl = file
            self.sl_file.setText(os.path.basename(self.act_sl))
            self.sl_file.setToolTip(self.act_sl)
            self.readsl_btn.setText('Load')
            self.readsl_btn.setEnabled(True)
            config.set('last_used', 'sl', self.act_sl)
            with open('config.ini', 'w') as f:
                config.write(f)

    def change_drs(self):
        file = QFileDialog().getOpenFileName(None, 'Select Diamond Record Sheet xlsx', os.path.dirname(self.act_drs))[0]
        if len(file) != 0 and file != self.act_drs:
            self.act_drs = file
            self.drs_file.setText(os.path.basename(self.act_drs))
            self.drs_file.setToolTip(self.act_drs)
            self.readsl_btn.setText('Load')
            self.readsl_btn.setEnabled(True)
            config.set('last_used', 'drs', self.act_drs)
            with open('config.ini', 'w') as f:
                config.write(f)
    
    def change_sr(self):
        file = QFileDialog().getOpenFileName(None, 'Select Sampling Results xlsx', os.path.dirname(self.act_sr))[0]
        if len(file) != 0 and file != self.act_sr:
            self.act_sr = file
            self.sr_file.setText(os.path.basename(self.act_sr))
            self.sr_file.setToolTip(self.act_sr)
            self.readsr_btn.setText('Load')
            self.readsr_btn.setEnabled(True)
            config.set('last_used', 'sr', self.act_sr)
            with open('config.ini', 'w') as f:
                config.write(f)

    def import_sl(self):
        self.readsl_btn.setEnabled(False)
        self.readsl_btn.setText('Loading')
        QApplication.processEvents()
        self.sldf = read_sl(self.act_sl)
        if self.sldf is not None:
            print(f'import complete, received {len(self.sldf)} records')
            self.readsl_btn.setText('Reload')
            #
        else:
            print('import failed')
            self.readsl_btn.setText('Load')
        self.readsl_btn.setEnabled(True)
        return
    
    def import_drs(self):
        self.readdrs_btn.setEnabled(False)
        self.readdrs_btn.setText('Loading')
        QApplication.processEvents()
        self.drsdf = read_drs(self.act_drs)
        if self.drsdf is not None:
            print(f'import complete, received {len(self.drsdf)} records')
            self.readdrs_btn.setText('Reload')
            #
        else:
            print('import failed')
            self.readdrs_btn.setText('Load')
        self.readdrs_btn.setEnabled(True)
        return
    
    def import_sr(self):
        self.readsr_btn.setEnabled(False)
        self.readsr_btn.setText('Loading')
        QApplication.processEvents()
        self.srdf = read_sr(self.act_sr)
        if self.srdf is not None:
            print(f'import complete, received {len(self.srdf)} records')
            self.readsr_btn.setText('Reload')
            #
        else:
            print('import failed')
            self.readsr_btn.setText('Load')
        self.readsr_btn.setEnabled(True)
        return

    def recon_drs_sl(self):
        heading, report = ['EL RECON: DRS vs SCREENLOG'], [] 
        dates = [self.start_date.date().toString("yyyy-MM-dd"), self.end_date.date().toString("yyyy-MM-dd")]
        heading.append(f'{dates[0]} to {dates[1]}\n\n') 
        # update last date in config
        config.set('last_used', 'start', dates[0])
        config.set('last_used', 'end', dates[1])
        with open('config.ini', 'w') as f:
            config.write(f)
        if self.drsdf is None:
            edf = Las_Preguntas('Import DRS', 'No DRS loaded, import now?')
            if edf.exec() == QMessageBox.Yes:
                self.import_drs()
                self.recon_drs_sl()
            else:
                return
        if self.sldf is None:
            edf = Las_Preguntas('Import Screenlog', 'No Screenlog loaded, import now?')
            if edf.exec() == QMessageBox.Yes:
                self.import_sl()
                self.recon_drs_sl()
            else:
                return
        #filter drsdf:
        el_drsdf = self.drsdf[self.drsdf['Date_Drilled'].between(dates[0],dates[1])]
        audits = []
        mmcount = 0
        # reformat drs:
        carats, shapes, colours, clarities = [], [], [], []
        for i, row in el_drsdf.iterrows():
            row.dropna(inplace=True)
            if len(row) > 3 and (len(row)-5) % 4 != 0:
                report.append(f'Missing property: Sample {i} row length: {len(row)}')
            carats.append(row.iloc[5::4].values.tolist())
            shapes.append(row.iloc[6::4].values.tolist())
            colours.append(row.iloc[7::4].values.tolist())
            clarities.append(row.iloc[8::4].values.tolist())
            if 'ADT' in i:
                audits.append(i)
        el_drsdf = el_drsdf.iloc[:, :5]
        el_drsdf['Carats'] = carats
        el_drsdf['Shapes'] = shapes
        el_drsdf['Colours'] = colours
        el_drsdf['Clarities'] = clarities
        # now have a drs dataset per sample
        # filter sl:
        el_sldf = self.sldf[self.sldf['Start_Date'].between(dates[0], dates[1])]
        el_sldf.Est_Brkdwn = el_sldf.Est_Brkdwn.apply(listify_cts)
        el_sldf.Shape = el_sldf.Shape.apply(listify_shp)
        el_sldf.Colour = el_sldf.Colour.apply(listify_colcla)
        el_sldf.Clarity = el_sldf.Clarity.apply(listify_colcla)

        # begin comparison:
        heading.append(f'Sample count:\tSL: {len(el_sldf)}\tDRS: {len(el_drsdf)} (+{len(audits)} audits)')
        if len(el_drsdf) - len(el_sldf) != len(audits):
            report.append('Sample count mismatch')
        else:
            report.append('Sample counts match')
        for sid, sample in el_sldf.iterrows():
            ind_err = []
            try:
                if sample.Est_Brkdwn != el_drsdf.loc[sid, 'Carats']:
                    drsval = el_drsdf.loc[sid, 'Carats']
                    ind_err.append(f'{sample.Est_Brkdwn}\n{drsval}')
                if sample.Shape != el_drsdf.loc[sid, 'Shapes']:
                    drsval = el_drsdf.loc[sid, 'Shapes']
                    ind_err.append(f'{sample.Shape}\n{drsval}')
                if sample.Colour != el_drsdf.loc[sid, 'Colours']:
                    drsval = el_drsdf.loc[sid, 'Colours']
                    ind_err.append(f'{sample.Colour}\n{drsval}')
                if sample.Clarity != el_drsdf.loc[sid, 'Clarities']:
                    drsval = el_drsdf.loc[sid, 'Clarities']
                    ind_err.append(f'{sample.Clarity}\n{drsval}')
            except KeyError as k:
                ind_err.append(f'Sample not found: {str(k)}')
            except Exception as e:
                ind_err.append(f'Err: {type(e)}: {str(e)}')
            if len(ind_err) > 0:
                report.append(f'\n{sid}')
                report.extend(ind_err)
                link = f'open_log({os.path.dirname(self.act_sl)}/Scanned Logs/Diamond Sorting Logs/ML103A_DSL_{sid}.jpg)' 
                report.append(f'<a href="{link}">View Log</a>') # VIEW LOG link here..
                mmcount += 1
        if mmcount == 0:
            report.append('Stone weights/properties: 100% match')
        else:
            heading.insert(-1, '***Screenlog value shown first')
            report.insert(0, f'{mmcount} total mismatches')
            report.append('\n\n[end of report]')
        print(report)
        self.final = El_Informe('REPORT','Recon finished', heading + report)
        self.final.show()
        

    def recon_sl_sr(self):
        Los_Errors('Exciting', 'this section not work\nunder construction')
        return
        heading, report = ['EL RECON: SCREENLOG vs SAMPLING RESULTS'], []
        dates = [self.start_date.date().toString("yyyy-MM-dd"), self.end_date.date().toString("yyyy-MM-dd")]
        heading.append(f'{dates[0]} to {dates[1]}\n\n') # report 1
        # update last date in config
        config.set('last_used', 'start', dates[0])
        config.set('last_used', 'end', dates[1])
        with open('config.ini', 'w') as f:
            config.write(f)
        if self.sldf is None:
            edf = Las_Preguntas('Import Screenlog', 'No Screenlog loaded, import now?')
            if edf.exec() == QMessageBox.Yes:
                self.import_sl()
                self.recon_sl_sr()
            else:
                return
        if self.srdf is None:
            edf = Las_Preguntas('Import SR', 'No SR loaded, import now?')
            if edf.exec() == QMessageBox.Yes:
                self.import_sr()
                self.recon_sl_sr()
            else:
                return
        # filter sl:
        el_sldf = self.sldf[self.sldf['Start_Date'].between(dates[0], dates[1])]
        el_sldf.Est_Brkdwn = el_sldf.Est_Brkdwn.apply(listify_cts)
        el_sldf.Shape = el_sldf.Shape.apply(listify_shp)
        el_sldf.Colour = el_sldf.Colour.apply(listify_colcla)
        el_sldf.Clarity = el_sldf.Clarity.apply(listify_colcla)
        # filter sr:
        el_srdf = self.srdf[self.srdf['Date'].between(dates[0], dates[1])]

        # begin comparison:
        final = Los_Errors('REPORT', 'Recon finished', report)
        return

def read_drs(name): # date filter later, upon recon
    print('\nReading Diamond Record Sheet . . .')
    read_init = dt.datetime.now()
    try:
        sheetname = None
        sheets = pd.ExcelFile(name).sheet_names
        for sh in sheets:
            if 'DiamondSortingLog' in sh:
                sheetname = sh
                break
        if sheetname == None:
            report = sheets.insert(0, 'Sheets found in DRS:')
            errout = Los_Errors('DRS Sheet Name', 'Could not find a sheet name containing "DiamondSortingLog"', report)
            return None
        drs = pd.read_excel(name, sheet_name=sheetname, skiprows=1, usecols='A,C,D,H,AM:AO,AV:AOM')
        drs.dropna(how='all', inplace=True)
        drs.dropna(axis=1, how='all', inplace=True)
        drs.columns = np.arange(0,drs.shape[1])
        drs.rename({
            0:'Seq',
            1:'Sample_ID',
            2:'Date_Drilled',
            3:'Date_Sorted_(start)', # for Audits, use date sorted as date
            4:'Total_Stones',
            5:'Total_Carats',
            6:'Group_Weight'}, 
            axis=1, inplace=True)
        drs['Date_Drilled'].fillna(drs['Date_Sorted_(start)'], inplace=True)
        drs.drop(['Date_Sorted_(start)'], axis=1, inplace=True)
        drs['Group_Weight'].fillna(drs['Total_Carats'], inplace=True)
        drs['Date_Drilled'] = pd.to_datetime(drs['Date_Drilled'], errors='coerce')
        ''' FILTER LATER:'''
        #drs = drs[drs['Date Drilled'].between(dates[0], dates[1])] # ev dropped here
        drs.dropna(axis=1, how='all', inplace=True) # drop empty cols
        col_dict = {}
        props = {1:'Carat', 2:'Shape', 3:'Colour', 4:'Clarity'}
        max_stones = (drs.shape[1] - 6) / 4 # inferred, check that == int
        if max_stones != int(max_stones):
            errout = Los_Errors('Fatal Error', f'Error in Diamond Record Sheet! \nMax number of stones evaluates to {max_stones}')
            return None
        # stones start at column 7
        for d in range(1, int(max_stones) + 1): # 1 to 19
            for p in range(1,5):
                col_dict[d*4 + p + 2] = 'D' + str(d) + props[p]
        drs.rename(col_dict, axis=1, inplace=True)
        drs.set_index('Sample_ID', inplace = True, drop=True)
        print(drs.shape)
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('Diamond Record Sheet processed in', round(dur, 2), 'seconds (', len(drs), ' records)')
        return drs#, max_stones
    except Exception as e:
        print(e, '\nError reading Diamond Record Sheet')
        return None

# func for listifying sl cells
def listify_cts(entry):
    if pd.notna(entry):
        try:
            l = entry.split(',')
            return [float(s.strip()) for s in l]
        except AttributeError:
            return [float(entry)]
    else:
        return []

def listify_shp(entry):
    if pd.notna(entry):
        try:
            l = entry.split(',')
            return [int(s.strip()) for s in l]
        except AttributeError:
            return [int(entry)]
    else:
        return []

def listify_colcla(entry):
    if pd.notna(entry):
        try:
            l = entry.split(',')
            return [s.strip() for s in l]
        except AttributeError:
            return [str(entry).strip()]
    else:
        return []

def read_sl(name): # date filter later, upon recon
    print('\nReading Screenlog . . .')
    read_init = dt.datetime.now()
    try:
        sl = pd.read_excel(name, engine='pyxlsb', usecols=[
            'Seq',
            'Sample_ID',
            'Start_Date',
            'Stones',
            'Group_wt',
            'Est_Carats',
            'Est_Brkdwn',
            'Shape',
            'Colour',
            'Clarity'
            ], index_col='Sample_ID')
        #sl.dropna(how='all', subset=['Stones'], inplace=True)
        sl.Stones.fillna(0, inplace=True)
        # experiment: change "sl['Stones'].fillna" to "sl.Stones.fillna"
        sl['Start_Date'] = pd.TimedeltaIndex(sl['Start_Date'], unit = 'd') + EPOCH
        ''' FILTER LATER:'''
        #sl = sl[sl['Start_Date'].between(dates[0], dates[1])]
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('Screenlog processed in', round(dur, 2), 'seconds (', len(sl), ' records)')
        return sl
    except Exception as e:
        print(e, '\nError reading Screenlog')
        return None

def read_sr(name):  # date filter later, upon recon
    print('\nReading Sampling Results . . .')
    read_init = dt.datetime.now()
    try:
        sr = pd.read_excel(name, skiprows=4, usecols=[
            'Date',
            'Sample N.\n&\nBarcode N.',
            'Individual  Stone Count Offshore',
            'IndividualCarats Offshore ',
            'Individual Stone Shape',
            'Individual Stone Colour',
            'Individual Stone Clarity'
            ])
        sr.rename(columns={'Sample N.\n&\nBarcode N.':'Sample_ID'}, inplace=True)
        sr.dropna(how='all', inplace=True)
        sr.drop([1, 2], inplace=True)
        sr['Date'] = pd.to_datetime(sr['Date'], errors='coerce')
        sr['Date'].fillna(method='ffill', inplace=True)
        ''' FILTER LATER:'''
        #sr = sr[sr['Date'].between(dates[0], dates[1])]
        sr['Sample_ID'].fillna(method='ffill', inplace=True)
        sr.drop(sr[sr['Sample_ID'].str.contains('Samples')].index, inplace=True)
        # make a dataset per sample:
        sr = sr.groupby('Sample_ID', sort=False).agg(
            count=('Individual  Stone Count Offshore','sum'),
            carats=('IndividualCarats Offshore ','sum'),
            brkdwn=('IndividualCarats Offshore ',lambda f: f.tolist()),
            shapes=('Individual Stone Shape', lambda f: f.tolist()),
            colour=('Individual Stone Colour', lambda f: f.tolist()),
            clarity=('Individual Stone Clarity', lambda f: f.tolist())
        )
        #sr.set_index(['Sample_ID'], inplace = True, drop=True)
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('Sampling Results processed in', round(dur, 2), 'seconds (', len(sr), ' records)')
        return sr
    except Exception as e:
        print(e, '\nError reading Sampling Results sheet')
        return None




'''
def compare_lists(list1, list2, label): # use df1.eq(df2)
    mismatches = []
    for i, (d, s) in enumerate(zip(list1, list2)):
        if str(d) != str(s):
            try:
                if round(float(d), 2) != round(float(s), 2):
                    issue = ('\t' + label + ':\t"' + str(d) + '"\t\t\t"' + str(format(s, ':4.2f')) + '"')
                    mismatches.append([i, issue])
            except Exception as e:
                log.append(e)
                issue = ('\t' + label + ':\t"' + str(d) + '"\t\t\t"' + str(s) + '"')
                mismatches.append([i, issue])
    return mismatches # a list of lists

def compare_stones(list1, list2, sid): # use df1.eq(df2)
    mismatches = []
    for i, (sl, sr) in enumerate(zip(list1, list2)):
        if str(sl) != str(sr):
            try:
                if round(float(sl), 2) != round(float(sr), 2):
                    issue = ('\t' + props.get(i) + ':\t"' + str(sl) + '"\t\t\t"' + str(format(sr, ':4.2f')) + '"')
                    mismatches.append([sid, issue])
            except Exception as e:
                log.append(e)
                issue = ('\t' + props.get(i) + ':\t"' + str(sl) + '"\t\t\t"' + str(sr) + '"')
                mismatches.append([sid, issue])
    return mismatches # a list of lists

def compare_series(series1, series2, label, ids): # ids is a series of SampleIDs
    if series1.eq(series2).any() == False:
        mm = compare_lists(series1.values.tolist(), series2.values.tolist(), label)
        if len(mm) == 0:
            print(label, 'match')
        else:
            for m in mm:
                print('Check', label, 'SampleID:\t', ids.loc[m[0]])
    else:
        print('No mismatches in', label)

def sl_vs_drs(sl, drs, max_stones):
    print('Comparing Screenlog and Diamond Record Sheet data (Audits not included) . . .')
    drs.dropna(how='all', subset=['Seq No.'], inplace=True)
    drs.reset_index(drop=True, inplace=True) # no audits in screenlog
    # lines to attempt to read in DRS:
    if len(drs) != len(sl):
        print('Number of samples in Screenlog (', len(sl), ') and DRS (', len(drs), ') do not match')
    else:
        print(len(sl), 'samples in Screenlog match', len(drs), 'samples in DRS')

    collec_cts, collec_shps, collec_cols, collec_cla = [], [], [], []
    for i, r in drs.iterrows():
        #prt('\nrow: ' + str(i) + '\t')     iterator tracker
        drs_carats, drs_count, drs_shapes, drs_colours, drs_clarities = [], [], [], [], []
        for n in range(7,(max_stones*4)+6,4): # (max_stones*4)+6   +6??? +4??? +7/8??
            #print(n, end='', flush=True)
            if pd.notnull(drs.iloc[i, n]): # insert validation here
                drs_carats.append(format(drs.iloc[i, n], '.2f'))
                drs_shapes.append(format(drs.iloc[i, n+1], '.0f'))
                drs_colours.append(str(drs.iloc[i, n+2]))
                drs_clarities.append(str(drs.iloc[i, n+3]))
            #prt('.')
        collec_cts.append(', '.join(drs_carats))
        collec_shps.append(', '.join(drs_shapes))
        collec_cols.append(', '.join(drs_colours))
        collec_cla.append(', '.join(drs_clarities))

    sl.fillna('', inplace=True)
    drs['Group Weight'].fillna('0.0', inplace=True)
    drs['Total Carats'].fillna('0.0', inplace=True)
    mm_cts = compare_lists(collec_cts, sl['Est_Brkdwn'].values.tolist(),'Carat')
    mm_shps = compare_lists(collec_shps, sl['Shape'].values.tolist(),'Shape')
    mm_cols =  compare_lists(collec_cols, sl['Colour'].values.tolist(),'Colour')
    mm_cla = compare_lists(collec_cla, sl['Clarity'].values.tolist(),'Clarity')

    
    mms = (mm_cts + mm_shps + mm_cols + mm_cla)
    if len(mms) > 0:
        print('Seq:\tSampleID:\tIssue\t\tDRS_val:\t\t\t\tSL_val:')
    else:
        print('No mismatches in Carats, Shapes, Colours, Clarities')
    for mm in mms:
        print(drs.loc[mm[0], 'Seq No.'], '\t', drs.loc[mm[0], 'Sample ID'], mm[1])

    compare_series(drs['Total Stones'], sl['Stones'],'Total Stones',drs['Sample ID'])
    compare_series(drs['Group Weight'], sl['Group_wt'],'Group Weight',drs['Sample ID'])
    compare_series(drs['Total Carats'], sl['Est_Carats'],'Total Carats',drs['Sample ID'])

    
    print('\n\n\n[ DRS vs SL complete ]')

def sl_vs_sr(sl, sr):
    print('Comparing Screenlog and SamplingResults data (Audits not included) . . .')
    # convert sl data to per stone i.e. for each sample create 1 line, and 1 additional line/stone > 1
    # create a table similar to the SR sheet
    #sl_cts, sl_shps, sl_cols, sl_cla = [], [], [], []
    sl_per_stone = []
    sl_count = 0
    samples = sl['Sample_ID'].values.tolist()
    for i, r in sl.iterrows():
        stonecount = int(r['Stones'])
        sl_count += stonecount
        if stonecount == 0:
            sl_per_stone.append([format(0, '.2f'), '', '', ''])
        elif stonecount == 1:
            sl_per_stone.append([r['Est_Brkdwn'], r['Shape'], r['Colour'], r['Clarity']])
        else:
            for s in range(stonecount):
                sl_per_stone.append([
                    (r['Est_Brkdwn'].split(', '))[s],  # ct
                    (r['Shape'].split(', '))[s],       # shape
                    (r['Colour'].split(', '))[s],      # colour
                    (r['Clarity'].split(', '))[s]      # clarity
                ])


    #drop audits, sterilizations from sr
    sr.drop(sr[
        (sr['Sample N.\n&\nBarcode N.'].str.contains('ADT')) | (sr['Sample N.\n&\nBarcode N.'].str.contains('Steril'))
        ].index, inplace=True)
    # reset index sr
    sr.reset_index(drop=True, inplace=True)
    sr_count = sr['Individual  Stone Count Offshore'].sum()
    sr.fillna('', inplace=True)
    print(f'Screenlog shows {sl_count} total stones, SR shows {sr_count}')
    print('SampleID\tProperty\tScreenlog value\t\tSR value')
    for i, r in sr.iterrows():
        sr_stone = [r['IndividualCarats Offshore '], r['Individual Stone Shape'], r['Individual Stone Colour'], r['Individual Stone Clarity']]
        mms = compare_stones(sl_per_stone[i], sr_stone, r['Sample N.\n&\nBarcode N.'])
        for mm in mms:
            print(mm[0], mm[1])

    print('\n\n\n[ SL vs SR complete ]')
    '''







# MAIN FUNCTION
def main():
    # Create an instance of QApplication
    dos = QApplication(sys.argv)
    # Show the GUI
    recon = QMainWindow()
    recon.setWindowTitle('IMDH Recon II')
    recon.setWindowIcon(QIcon('icon.ico'))
    recon.resize(360, 220)
    recon.ui = El_Recon_Dos(recon)
    recon.setCentralWidget(recon.ui)
    recon.show()
    # main loop
    sys.exit(dos.exec_())


if __name__ == '__main__':
    main()