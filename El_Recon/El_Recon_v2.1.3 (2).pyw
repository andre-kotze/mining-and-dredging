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

__version__ = '2.1.3'
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
    def __init__(self, title=' ', msg='No'):
        super(QMessageBox, self).__init__()
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon('icon.ico'))
        self.setText(msg)
        self.setStandardButtons(QMessageBox.Ok | QMessageBox.No)
        self.button(QMessageBox.Ok).setText('No')
        self.exec()

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
        if address[:4] == 'file':
            address = address[8:]
            print(f'address: {address}')
            self.report.setSource(QtCore.QUrl())
            if os.path.exists(address):
                print('calling..')
                #os.startfile(address)
            else:
                Los_Errors('404', f'File not found\n{address}')
        else:
            Los_Errors('WTF',f'Unrecognised link:\n{address}')

class El_Recon_Dos(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__()
        self.act_sl = config.get('last_used', 'sl')
        self.act_drs = config.get('last_used', 'drs')
        self.act_sr = config.get('last_used', 'sr')
        self.report_nr = int(config.get('stats', 'reports'))
        print(f'report nr: {self.report_nr + 1}')
        self.gt_mistakes = int(config.get('stats', 'mistakes'))
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
        self.slbbtn.setFixedSize(22, 22)
        self.readsl_btn = QPushButton('Load', self)
        self.readsl_btn.clicked.connect(self.import_sl)
        self.open_sl_btn = QPushButton('Open', self)
        self.open_sl_btn.setFlat(True)
        self.open_sl_btn.setToolTip('Open in Excel')
        self.open_sl_btn.setStyleSheet('QPushButton {color:blue; max-width:40px; text-decoration:underline}')
        self.open_sl_btn.clicked.connect(lambda: os.startfile(self.act_sl))

        # Loaded DRS label
        self.drs_file = QLabel(os.path.basename(self.act_drs))
        self.drs_file.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.drs_file.setToolTip(self.act_drs)
        self.drsbbtn = QPushButton('...', self)
        self.drsbbtn.clicked.connect(self.change_drs)
        self.drsbbtn.setFixedSize(22, 22)
        self.readdrs_btn = QPushButton('Load', self)
        self.readdrs_btn.clicked.connect(self.import_drs)
        self.open_drs_btn = QPushButton('Open', self)
        self.open_drs_btn.setFlat(True)
        self.open_drs_btn.setToolTip('Open in Excel')
        self.open_drs_btn.setStyleSheet('QPushButton {color:blue; max-width:40px; text-decoration:underline}')
        self.open_drs_btn.clicked.connect(lambda: os.startfile(self.act_drs))

        # Loaded SR label
        self.sr_file = QLabel(os.path.basename(self.act_sr))
        self.sr_file.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.sr_file.setToolTip(self.act_sr)
        self.srbbtn = QPushButton('...', self)
        self.srbbtn.clicked.connect(self.change_sr)
        self.srbbtn.setFixedSize(22, 22)
        self.readsr_btn = QPushButton('Load', self)
        self.readsr_btn.clicked.connect(self.import_sr)
        self.open_sr_btn = QPushButton('Open', self)
        self.open_sr_btn.setFlat(True)
        self.open_sr_btn.setToolTip('Open in Excel')
        self.open_sr_btn.setStyleSheet('QPushButton {color:blue; max-width:40px; text-decoration:underline}')
        self.open_sr_btn.clicked.connect(lambda: os.startfile(self.act_sr))

        flay = QGridLayout()
        flay.setColumnStretch(1,1)

        flay.addWidget(QLabel('Screenlog:'),0,0)
        flay.addWidget(self.sl_file,0,1)
        flay.addWidget(self.open_sl_btn,0,2)
        flay.addWidget(self.slbbtn,0,3)
        flay.addWidget(self.readsl_btn,0,4)

        flay.addWidget(QLabel('Diamond Record Sheet:'),2,0)
        flay.addWidget(self.drs_file,2,1)
        flay.addWidget(self.open_drs_btn,2,2)
        flay.addWidget(self.drsbbtn,2,3)
        flay.addWidget(self.readdrs_btn,2,4)

        flay.addWidget(QLabel('Sampling Results:'),1,0)
        flay.addWidget(self.sr_file,1,1)
        flay.addWidget(self.open_sr_btn,1,2)
        flay.addWidget(self.srbbtn,1,3)
        flay.addWidget(self.readsr_btn,1,4)

        return flay

    def dates(self):
        self.start_date = QDateEdit(calendarPopup=True)
        self.start_date.setStyleSheet("QDateEdit{min-width: 90px}")
        self.start_date.setDate(self.last_start)
        self.end_date = QDateEdit(calendarPopup=True)
        self.end_date.setStyleSheet("QDateEdit{min-width: 90px}")
        self.end_date.setDate(self.last_end)
        dlay = QHBoxLayout()
        dlay.addStretch(1)
        dlay.addWidget(QLabel('Date range:'))
        dlay.addWidget(self.start_date)
        dlay.addWidget(QLabel(' to '))
        dlay.addWidget(self.end_date)

        return dlay

    def buttons(self):
        blay = QHBoxLayout()
        self.drs_vs_sl_btn = QPushButton('DRS vs SL', self)
        self.sl_vs_sr_btn = QPushButton('SL vs SR', self)
        self.drs_vs_sr_btn = QPushButton('DRS vs SR', self)
        self.threesome_btn = QPushButton('3-WAY RECON', self)
        #self.threesome_btn.setDisabled(True)
        #self.drs_vs_sr_btn.setDisabled(True)
        self.drs_vs_sl_btn.setMaximumWidth(120)
        self.sl_vs_sr_btn.setMaximumWidth(120)
        self.drs_vs_sr_btn.setMaximumWidth(120)
        self.threesome_btn.setMaximumWidth(120)
        self.drs_vs_sl_btn.clicked.connect(self.recon_drs_sl)
        self.sl_vs_sr_btn.clicked.connect(self.recon_sl_sr)
        self.drs_vs_sr_btn.clicked.connect(lambda: Las_Preguntas('NO'))
        self.threesome_btn.clicked.connect(lambda: Las_Preguntas('NO'))
        blay.addWidget(self.drs_vs_sl_btn)
        blay.addWidget(self.sl_vs_sr_btn)
        blay.addWidget(self.drs_vs_sr_btn)
        blay.addWidget(self.threesome_btn)
        return blay

    def change_sl(self):
        file = QFileDialog().getOpenFileName(None, 'Select Screenlog xlsb', os.path.dirname(self.act_sl))[0]
        if len(file) != 0 and file != self.act_sl:
            self.act_sl = file
            self.sldf = None
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
            self.drsdf = None
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
            self.srdf = None
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
        inferred_conc = os.path.splitext(self.act_sl)[0].rsplit('_', 1)[1]
        report = [] 
        audits = []
        mmcount = 0
        dates = [self.start_date.date().toString("yyyy-MM-dd"), self.end_date.date().toString("yyyy-MM-dd")]
        # update last date in config
        config.set('last_used', 'start', dates[0])
        config.set('last_used', 'end', dates[1])
        with open('config.ini', 'w') as f:
            config.write(f)
        if self.drsdf is None:
            self.import_drs()
            if self.drsdf is None:
                Los_Errors('ERROR', 'Failed to import DRS')
                return

        if self.sldf is None:
            self.import_sl()
            if self.sldf is None:
                Los_Errors('ERROR', 'Failed to import SL')
                return
        #filter drsdf:
        el_drsdf = self.drsdf[self.drsdf['Date_Drilled'].between(dates[0],dates[1])]
        if len(el_drsdf) == 0:
            Los_Errors('No data', 'Empty DRS for selected dates')
            return
        # reformat drs:
        carats, shapes, colours, clarities = [], [], [], []
        for i, row in el_drsdf.iterrows():
            row.dropna(inplace=True)
            if len(row) > 3 and (len(row)-5) % 4 != 0:
                report.append(f'Missing property: Sample {i} row length: {len(row)}')
                mmcount += 1
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
        if len(el_sldf) == 0:
            Los_Errors('No data', 'Empty SL for selected dates')
            return
        el_sldf.Stones.fillna(0, inplace=True)
        el_sldf.Est_Brkdwn = el_sldf.Est_Brkdwn.apply(listify_cts)
        el_sldf.Shape = el_sldf.Shape.apply(listify_shp)
        el_sldf.Colour = el_sldf.Colour.apply(listify_colcla)
        el_sldf.Clarity = el_sldf.Clarity.apply(listify_colcla)

        # begin comparison:
        if len(el_drsdf) - len(el_sldf) != len(audits):
            refs = [f'Sample count mismatch:\tSL: {len(el_sldf)}\tDRS: {len(el_drsdf) - len(audits)} ({len(el_drsdf)} - {len(audits)} audits)']
        else:
            refs = [f'Sample counts match:\tSL: {len(el_sldf)}\tDRS: {len(el_drsdf) - len(audits)} ({len(el_drsdf)} - {len(audits)} audits)']
        for sid, sample in el_sldf.iterrows():
            ind_err = []
            try:
                if sample.Est_Brkdwn != el_drsdf.loc[sid, 'Carats']:
                    ind_err.append(f'{sample.Est_Brkdwn}\n{el_drsdf.loc[sid, "Carats"]}')
                if sample.Shape != el_drsdf.loc[sid, 'Shapes']:
                    ind_err.append(f'{sample.Shape}\n{el_drsdf.loc[sid, "Shapes"]}')
                if sample.Colour != el_drsdf.loc[sid, 'Colours']:
                    ind_err.append(f'{sample.Colour}\n{el_drsdf.loc[sid, "Colours"]}')
                if sample.Clarity != el_drsdf.loc[sid, 'Clarities']:
                    ind_err.append(f'{sample.Clarity}\n{el_drsdf.loc[sid, "Clarities"]}')
            except KeyError as k:
                ind_err.append(f'Sample not found in DRS: {str(k)}')
            except Exception as e:
                ind_err.append(f'Err: {type(e)}: {str(e)}')
            if len(ind_err) > 0:
                report.append(f'\n{sid}')
                report.extend(ind_err)
                link = f'file:///{os.path.dirname(self.act_sl)}/Scanned Logs/Diamond Sorting Logs/{inferred_conc}_DSL_{sid}.jpg'
                anchor = f'<a href="{link}" target="_blank">View Log</a>'
                if os.path.exists(link.split('///')[1]):
                    report.append(anchor) # VIEW LOG link here..
                mmcount += 1
        if mmcount == 0:
            refs.append('Stone weights/properties: 100% match')
        else:
            refs.append('* Screenlog value shown first')
            if mmcount == 1:
                refs.append(f'{mmcount} discrepancy found:')
            else:
                refs.append(f'{mmcount} discrepancies found:')
            report.append('\n\n[end of report]')
        
        title = 'Geological Screen Log vs Diamond Record Sheet'
        content = {
            'title' : title, #str
            'dates' : f'{inferred_conc}: {dates[0]} to {dates[1]}', #str
            'refs' : refs, #list
            'report' : report, #list
            'num' : (self.report_nr + 1)
        }
        if external_report(content):
            self.report_nr += 1
            self.gt_mistakes += mmcount
            config.set('stats', 'reports', str(self.report_nr))
            config.set('stats', 'mistakes', str(self.gt_mistakes))
            with open('config.ini', 'w') as f:
                config.write(f)
        else:
            print('report creation failed')
        

    def recon_sl_sr(self):
        inferred_conc = os.path.splitext(self.act_sl)[0].rsplit('_', 1)[1]
        report = []
        mmcount = 0
        dates = [self.start_date.date().toString("yyyy-MM-dd"), self.end_date.date().toString("yyyy-MM-dd")]
        # update last date in config
        config.set('last_used', 'start', dates[0])
        config.set('last_used', 'end', dates[1])
        with open('config.ini', 'w') as f:
            config.write(f)
        if self.sldf is None:
            self.import_sl()
            if self.sldf is None:
                Los_Errors('ERROR', 'Failed to import SL')
                return
        if self.srdf is None:
            self.import_sr()
            Los_Errors('ERROR', 'Failed to import SR')
            if self.srdf is None:
                return
        # filter sl:
        el_sldf = self.sldf[self.sldf['Start_Date'].between(dates[0], dates[1])]
        if len(el_sldf) == 0:
            Los_Errors('No data', 'Empty SL for selected dates')
            return
        el_sldf.Est_Brkdwn = el_sldf.Est_Brkdwn.apply(listify_cts)
        el_sldf.Shape = el_sldf.Shape.apply(listify_shp)
        el_sldf.Colour = el_sldf.Colour.apply(listify_colcla)
        el_sldf.Clarity = el_sldf.Clarity.apply(listify_colcla)
        unsorted = len(el_sldf[pd.isna(el_sldf.Stones)]) # count unsorteds
        el_sldf.dropna(how='all', subset=['Stones'], inplace=True) # drop unsorteds

        # filter sr:
        el_srdf = self.srdf[self.srdf['Date'].between(dates[0], dates[1])]
        if len(el_srdf) == 0:
            Los_Errors('No data', 'Empty SR for selected dates')
            return
        el_srdf['Sample_ID'].fillna(method='ffill', inplace=True)
        el_srdf.drop(el_srdf[el_srdf['Sample_ID'].str.contains('Samples')].index, inplace=True)
        
        
        el_srdf.loc[el_srdf['IndividualCarats Offshore '] == 0, 'IndividualCarats Offshore '] = np.nan
        # make a dataset per sample:
        el_srdf = el_srdf.groupby('Sample_ID', sort=False).agg(
            count=('Individual  Stone Count Offshore','sum'),
            carats=('IndividualCarats Offshore ','sum'),
            brkdwn=('IndividualCarats Offshore ', lambda f: f.dropna().tolist()),
            shapes=('Individual Stone Shape', lambda f: f.dropna().tolist()),
            colours=('Individual Stone Colour', lambda f: f.dropna().tolist()),
            clarities=('Individual Stone Clarity', lambda f: f.dropna().tolist())
        )
        
        # begin comparison:
        audits = el_srdf[(el_srdf.index.str.contains('ADT')) | (el_srdf.index.str.contains('Steril')) | (el_srdf.index.str.contains('PURGE'))].index
        if len(el_srdf) - len(el_sldf) != len(audits):
            refs = [f'Sample count mismatch:\tSL: {len(el_sldf)}\tSR: {len(el_srdf) - len(audits)} ({len(el_srdf)} - {len(audits)} audits)']
        else:
            refs = [f'Sample counts match:\tSL: {len(el_sldf)}\tSR: {len(el_srdf) - len(audits)} ({len(el_srdf)} - {len(audits)} audits)']
        el_srdf.drop(audits, inplace=True)
        refs.append(f'{unsorted} unsorted samples dropped (ABD/ABT) from SL')
        for sid, sample in el_sldf.iterrows():
            ind_err = []
            try:
                if sample.Est_Brkdwn != el_srdf.loc[sid, 'brkdwn']:
                    srval = el_srdf.loc[sid, 'brkdwn']
                    ind_err.append(f'{sample.Est_Brkdwn}\n{srval}')
                if sample.Shape != el_srdf.loc[sid, 'shapes']:
                    srval = el_srdf.loc[sid, 'shapes']
                    ind_err.append(f'{sample.Shape}\n{srval}')
                if sample.Colour != el_srdf.loc[sid, 'colours']:
                    srval = el_srdf.loc[sid, 'colours']
                    ind_err.append(f'{sample.Colour}\n{srval}')
                if sample.Clarity != el_srdf.loc[sid, 'clarities']:
                    srval = el_srdf.loc[sid, 'clarities']
                    ind_err.append(f'{sample.Clarity}\n{srval}')
            except KeyError as k:
                ind_err.append(f'Sample not found: {str(k)}')
            except Exception as e:
                ind_err.append(f'Err: {type(e)}: {str(e)}')
            if len(ind_err) > 0:
                report.append(f'\n{sid}')
                report.extend(ind_err)
                link = f'file:///{os.path.dirname(self.act_sl)}/Scanned Logs/Diamond Sorting Logs/{inferred_conc}_DSL_{sid}.jpg'
                anchor = f'<a href="{link}" target="_blank">View Log</a>'
                if os.path.exists(link.split('///')[1]):
                    report.append(anchor) # VIEW LOG link here..
                mmcount += 1
        if mmcount == 0:
            refs.append('Stone weights/properties: 100% match')
        else:
            refs.append('* Screenlog value shown first')
            if mmcount == 1:
                refs.append(f'{mmcount} discrepancy found')
            else:
                refs.append(f'{mmcount} discrepancies found')
        report.append('\n\n[end of report]')
        
        title = 'Geological Screen Log vs Sampling Results Sheet'
        content = {
            'title' : title, #str
            'dates' : f'{inferred_conc}: {dates[0]} to {dates[1]}', #str
            'refs' : refs, #list
            'report' : report, #list
            'num' : (self.report_nr + 1)
        }

        if external_report(content):
            self.report_nr += 1
            self.gt_mistakes += mmcount
            config.set('stats', 'reports', str(self.report_nr))
            config.set('stats', 'mistakes', str(self.gt_mistakes))
            with open('config.ini', 'w') as f:
                config.write(f)
        else:
            print('report creation failed')

def external_report(content): #content as a dict of strings of html
    filename = f'reports/report_{content["num"]}.html'
    report = open(filename, 'w')
    with open('report.html', 'r') as template:
        for l in template.readlines():
            report.write(l)
        report.write(f'<h2>REPORT: {content["title"]}</h2>')
        report.write(f'<h3>{content["dates"]}</h3>')
        for r in content['refs']:
            report.write(f'<strong>{r}</strong><br>')
        for e in content['report']:
            report.write(f'<p>{e}</p>')
        report.write(f'<p style="color: red">report nr {content["num"]}</p>')
        report.write('</body></html>')
        #report finalised
    report.close()
    os.startfile(os.path.abspath(filename))
    return True

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
            Los_Errors('DRS Sheet Name', 'Could not find a sheet name containing "DiamondSortingLog"', report)
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
            Los_Errors('Fatal Error', f'Error in Diamond Record Sheet! \nMax number of stones evaluates to {max_stones}')
            return None
        # stones start at column 7
        for d in range(1, int(max_stones) + 1): # 1 to 19
            for p in range(1,5):
                col_dict[d*4 + p + 2] = 'D' + str(d) + props[p]
        drs.rename(col_dict, axis=1, inplace=True)
        drs.set_index('Sample_ID', inplace = True, drop=True)
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
            'EOH_Code',
            'Stones',
            'Group_wt',
            'Est_Carats',
            'Est_Brkdwn',
            'Shape',
            'Colour',
            'Clarity'
            ], index_col='Sample_ID')
        sl['Start_Date'] = pd.TimedeltaIndex(sl['Start_Date'], unit = 'd') + EPOCH
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
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('Sampling Results processed in', round(dur, 2), 'seconds (', len(sr), ' records)')
        return sr
    except Exception as e:
        print(e, '\nError reading Sampling Results sheet')
        return None

# MAIN FUNCTION
def main():
    # Create an instance of QApplication
    dos = QApplication(sys.argv)
    # Show the GUI
    recon = QMainWindow()
    recon.setWindowTitle('IMDH Recon II')
    recon.setWindowIcon(QIcon('icon.ico'))
    recon.resize(360, 200)
    recon.ui = El_Recon_Dos(recon)
    recon.setCentralWidget(recon.ui)
    recon.show()
    # main loop
    sys.exit(dos.exec_())


if __name__ == '__main__':
    main()