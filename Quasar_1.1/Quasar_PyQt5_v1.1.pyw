'''
PyQt5 GUI with buttons and options:
-show 'loaded' SL, option to change

-option for last sorted/logged sample, or all
-option for full/minimal fields

-button to export csv
-button to export shp

-remember last csv and shp export dest
'''
import os
import sys
import datetime as dt
import time

import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
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
)
from PyQt5.QtGui import QIcon, QPalette, QColor

__version__ = '1.0.10'
__author__ = 'AKO_Geo'

TODAY = dt.datetime.now().date().strftime('%Y%m%d')
config = ConfigParser()
config.read('config.ini')
#increment = config.get('last_used', 'increment') 
    # True: auto-increment filenames (overwrite = False)

# records -2:All
# records -3:Last Sorted
# records -2:Last Logged
# columns -2:Full
# columns -3:Minimal

class QuasarFeedback(QMessageBox):
    def __init__(self, title, msg, reportlist=None):
        super(QMessageBox, self).__init__()
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon('icon.ico'))
        self.setText(msg)
        if reportlist is not None:
            report = '\n'.join(reportlist)
            self.setDetailedText(report)
            self.setStyleSheet("QTextEdit{min-width: 480px; min-height: 666px}")

# Main widget and tab holder (and settings tab)
class QuasarUI(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.initQW()

    def initQW(self):
        # init vars
        QuasarExport.csvdest = config.get('last_used', 'csvdest')
        QuasarExport.shpdest = config.get('last_used', 'shpdest')
        QuasarExport.setdirs = config.get('last_used', 'setdirs') # True: save folders are set (choose = False)
        QuasarExport.lddsl = config.get('last_used', 'file')

        #QuasarExport.increment = increment
        self.csvdir = QLabel(QuasarExport.csvdest)
        self.shpdir = QLabel(QuasarExport.shpdest)
        self.truncate_label(QuasarExport.csvdest, self.csvdir)
        self.truncate_label(QuasarExport.shpdest, self.shpdir)

        # Init tab screen
        self.tabs = QTabWidget(self) # self arg non-critical
        self.exptab = QuasarExport(self)
        self.settab = QWidget(self) # self arg non-critical
        self.abouttab = QWidget()

        # Add tabs
        self.tabs.addTab(self.exptab, 'Export')
        self.tabs.addTab(self.settab, 'Settings')
        self.tabs.addTab(self.abouttab, 'About')

        self.abouttab.layout = QVBoxLayout(self)
        self.abouttext = QLabel(
            '''QUASAR Quick Sampling Results\n\nVersion 1.0.10\n
            Written 2021 by Andre "Fucking Legend" "Drikus" Kotze\n
            email: andre@nammail.net'''
        )
        self.abouttab.layout.addWidget(self.abouttext)
        self.abouttab.layout.addStretch(1)
        self.abouttab.setLayout(self.abouttab.layout)
        
        self.settab.layout = self.create_savesettings()
        self.settab.setLayout(self.settab.layout)
        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def create_savesettings(self):
        mainlayout = QVBoxLayout()
        title = QLabel('Export filename and folder options:')
        title.setMaximumHeight(24)

        destdircontainer = QHBoxLayout()
        self.destdirgroup = self.createDestGroup()
        self.destdirgroup.setStyleSheet('QLabel:disabled, QPushButton:disabled, QGroupBox:title:disabled {color:grey}')
        destdircontainer.addWidget(self.destdirgroup)
        destdirspacer = QSpacerItem(24, 40)
        destdircontainer.addItem(destdirspacer)

        self.btnchoose = QRadioButton('Choose on export') # -2
        self.btnchoose.setToolTip('Select folder and filename manually')
        self.btnchoose.toggled.connect(lambda: self.toggledirs(self.btnchoose))

        self.btnauto = QRadioButton('Set save directories:') # -3
        self.btnauto.setToolTip('Define fixed paths for output files')
        self.btnauto.toggled.connect(lambda: self.toggledirs(self.btnauto))
        if QuasarExport.setdirs == 'True':
            self.btnauto.setChecked(True)
        else:
            self.btnchoose.setChecked(True)
        
        self.saveoptgroup = QButtonGroup()
        self.saveoptgroup.addButton(self.btnchoose)  
        self.saveoptgroup.addButton(self.btnauto)

        mainlayout.addWidget(title)
        mainlayout.addWidget(self.btnchoose)
        mainlayout.addWidget(self.btnauto)
        mainlayout.addLayout(destdircontainer)
        mainlayout.addStretch(0)
        return mainlayout
        
    def createDestGroup(self):
        groupBox = QGroupBox('Destination directories')

        self.csvbtn = QPushButton('...', self)
        self.csvbtn.clicked.connect(lambda: self.select_dir('Select csv output directory', 'csvdest'))
        self.csvbtn.setFixedSize(20, 20)
        self.shpbtn = QPushButton('...', self)
        self.shpbtn.clicked.connect(lambda: self.select_dir('Select shp output directory', 'shpdest'))
        self.shpbtn.setFixedSize(20, 20)

        csvln = QHBoxLayout()
        csvln.addWidget(QLabel('CSV:'))
        csvln.addStretch(1)
        csvln.addWidget(self.csvdir)
        csvln.addWidget(self.csvbtn)

        shpln = QHBoxLayout()
        shpln.addWidget(QLabel('SHP:'))
        shpln.addStretch(1)
        shpln.addWidget(self.shpdir)
        shpln.addWidget(self.shpbtn)
        
        dirslayout = QVBoxLayout()
        dirslayout.addLayout(csvln)
        dirslayout.addLayout(shpln)
        groupBox.setLayout(dirslayout)
        return groupBox

    def toggledirs(self, btn):
        if btn.text() == 'Choose on export': # Choose on export
            if btn.isChecked():
                self.destdirgroup.setDisabled(True)
                #self.destdirgroup.setStyleSheet('QGroupBox:title {color: grey}; QLabel {color: grey;}; QPushButton {color: grey;}')
                QuasarExport.setdirs = 'False'
            else:
                self.destdirgroup.setDisabled(False)
                #self.destdirgroup.setStyleSheet('QGroupBox:title {color: white}; QLabel {color: white;}; QPushButton {color: white;}')
                QuasarExport.setdirs = 'True'
        elif btn.text() == 'Set save directories:': # Set dirs
            if btn.isChecked():
                self.destdirgroup.setDisabled(False)
                #self.destdirgroup.setStyleSheet('QGroupBox:title {color: white}; QLabel {color: white;}; QPushButton {color: white;}')
                QuasarExport.setdirs = 'True'
            else:
                self.destdirgroup.setDisabled(True)
                #self.destdirgroup.setStyleSheet('QGroupBox:title {color: grey}; QLabel {color: grey;}; QPushButton {color: grey;}')
                QuasarExport.setdirs = 'False'
        else:
            print(f'No button with id {btn.id()}')
            # Set save directories
        save_cfg('last_used', 'setdirs', QuasarExport.setdirs)

    # function to truncate long paths and add tooltip
    def truncate_label(self, destobj, dirlabel): # destobj = QuasarExport.xxxdest (to be measured), dirlabel = self.xxxdir
        if len(destobj) > 64:
            trunclbl = '...' + destobj[-64:]
            dirlabel.setText(trunclbl)
            dirlabel.setToolTip(destobj)
        else:
            dirlabel.setText(destobj)
            dirlabel.setToolTip(None)
        
    # for choosing destdirs:
    def select_dir(self, desc, ftypedest): # ftd = string (csvdest OR shpdest)
        dirname = QFileDialog().getExistingDirectory(None, desc, config.get('last_used', ftypedest)) + '/'
        if len(dirname) > 1:
            save_cfg('last_used', ftypedest, dirname)
            if ftypedest == 'csvdest':
                QuasarExport.csvdest = dirname
                self.truncate_label(QuasarExport.csvdest, self.csvdir)
            elif ftypedest == 'shpdest':
                QuasarExport.shpdest = dirname
                self.truncate_label(QuasarExport.shpdest, self.shpdir)
            else:
                QuasarFeedback('ERROR', f'Unknown ftypedest: {ftypedest}')

key = 'C:/Users/User/Desktop/SHARE/Andre/KEY.txt'
# Main function: Export tab
class QuasarExport(QWidget): #TableWidget
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.initQE()

    def initQE(self):
        # Loaded SL label
        self.lbl = QLabel('Loaded:')
        self.csl = QLabel(self.lddsl.rsplit('/', 1)[1], self)
        self.csl.setToolTip(self.lddsl)
        self.bbtn = QPushButton('...', self)
        self.bbtn.clicked.connect(self.change_lddsl)
        self.bbtn.setFixedSize(20, 20)

        # export csv button
        self.csvbtn = QPushButton('Export csv', self)
        self.csvbtn.clicked.connect(lambda: self.lookbusy('csv'))

        # export shape button
        self.shpbtn = QPushButton('Export shp', self)
        self.shpbtn.clicked.connect(lambda: self.lookbusy('shp'))

        # box containing label
        r1 = QHBoxLayout()
        r1.addWidget(self.lbl)
        r1.addWidget(self.csl)
        r1.addStretch(1)
        r1.addWidget(self.bbtn)

        # box containing records/fields boxes
        r2 = QHBoxLayout()
        self.recgroup = QButtonGroup()
        self.colgroup = QButtonGroup()
        r2.addWidget(self.createRecordGroup())
        r2.addWidget(self.createFieldGroup())

        # box containing export buttons
        r3 = QHBoxLayout()
        r3.addWidget(self.csvbtn)
        r3.addWidget(self.shpbtn)

        # box containing boxes
        vbox = QVBoxLayout()
        vbox.addLayout(r1)
        vbox.addLayout(r2)
        vbox.addStretch(1)
        vbox.addLayout(r3)

        self.setLayout(vbox)

    def createRecordGroup(self):
        groupBox = QGroupBox('Records to include')
        btnall = QRadioButton('All records')
        btnall.setToolTip('Export all rows')
        btnsorted = QRadioButton('Last sorted')
        btnsorted.setToolTip('All rows up to last non-empty "Status"')
        btnlogged = QRadioButton('Last logged')
        btnlogged.setToolTip('All rows up to last non-empty "Logger"')
        btnabd = QRadioButton('ABD/ABT holes')
        btnabd.setToolTip('Abandoned and aborted holes only')
        btnsorted.setChecked(True)
        
        self.recgroup.addButton(btnall)      # -2
        self.recgroup.addButton(btnsorted)   # -3
        self.recgroup.addButton(btnlogged)   # -4
        self.recgroup.addButton(btnabd)      # -5
        c1 = QVBoxLayout()
        c1.addWidget(btnall)
        c1.addWidget(btnsorted)
        c1.addWidget(btnlogged)
        c1.addWidget(btnabd)
        groupBox.setLayout(c1)
        return groupBox

    def createFieldGroup(self):
        groupBox = QGroupBox('Columns to include')
        btnfull = QRadioButton('All')      #-2
        btnfull.setToolTip('Parse all columns')
        btnmin = QRadioButton('Basic')     #-3
        btnmin.setToolTip('Minimal columns: Seq, ID, Coordinates, Stones')
        btnfull.setChecked(True)
        
        self.colgroup.addButton(btnfull)
        self.colgroup.addButton(btnmin)
        c2 = QVBoxLayout()
        c2.addWidget(btnfull)
        c2.addWidget(btnmin)
        groupBox.setLayout(c2)
        return groupBox

    def change_lddsl(self):
        file = QFileDialog().getOpenFileName(None, 'Select Screenlog xlsb', os.path.dirname(self.lddsl), 'Excel Binary Workbook (*.xlsb)')[0]
        if len(file) != 0 and file != self.lddsl:
            self.lddsl = file
            self.csl.setText(self.lddsl.rsplit('/', 1)[1])
            self.csl.setToolTip(self.lddsl)
            save_cfg('last_used', 'file', file)
            save_cfg('last_used', 'folder', os.path.dirname(file) + '/')

    def lookbusy(self, ftype):
        self.csvbtn.setDisabled(True)
        self.shpbtn.setDisabled(True)
        QApplication.processEvents()
        export(self.lddsl, self.recgroup.checkedId(), self.colgroup.checkedId(), ftype)
        self.csvbtn.setEnabled(True)
        self.shpbtn.setEnabled(True)

# for manual save
def save_file(desc, defname, tfilter): 
    file = QFileDialog().getSaveFileName(None, desc, config.get('last_used', 'manualdest') + defname, tfilter)[0]
    if file == '':
        return None, None
    filepath, filename = file.rsplit('/', 1)
    if filename[-4:] == '.csv' or filename[-4:] == '.shp':
        filename = filename[:-4]
    save_cfg('last_used', 'manualdest', filepath + '/')
    return filepath, filename # return filename sans extension

def exportcsv(df, outdest, filename):
    if os.path.exists(f'{outdest}/{filename}.csv') and QuasarExport.setdirs == 'True':
        n = 1
        newname = f'{filename}_{n}'
        while os.path.exists(f'{outdest}/{filename}_{n}.csv'):
            n += 1
            newname = f'{filename}_{n}'
        filename = newname
    df.to_csv(f'{outdest}/{filename}.csv')
    return f'Saved as {filename}.csv', outdest

def fieldlength_from_settings(row):
    ret = f'{row[1]}:{int(row[2])}'
    if pd.notna(row[3]):
        ret += f'.{int(row[3])}'
    return ret

def get_schema(dfcols):
    propschema = {}
    setts = pd.read_csv('fieldlengths.csv')
    setts['T_L_P'] = setts.agg(lambda r: fieldlength_from_settings(r), axis='columns')
    setts.drop(columns=['TYPE', 'LENGTH', 'PRECISION'], inplace=True)
    for f in setts.itertuples():
        propschema[f.FIELD] = f.T_L_P
    if len(dfcols) != len(propschema):
        print('schema columns mismatch')
    full_schema = {'geometry': 'Point', 'properties': propschema}
    return full_schema

def exportshp(df, outdest, filename):
    shp_schema = get_schema(df.columns)
    df['geometry'] = df.apply(lambda x: Point((x.E_planned, x.N_planned)), axis=1) # need to convert to float?
    geodf = gpd.GeoDataFrame(df, crs='EPSG:32733', geometry='geometry')
    if os.path.exists(f'{outdest}/{filename}.shp') and QuasarExport.setdirs == 'True':
        n = 1
        newname = f'{filename}_{n}'
        while os.path.exists(f'{outdest}/{filename}_{n}.shp'):
            n += 1
            newname = f'{filename}_{n}'
        filename = newname
    try:
        geodf.to_file(f'{outdest}/{filename}.shp', driver='ESRI Shapefile', schema=shp_schema) # catch value and key error
    except Exception as e1:
        print(e1)
        try:
            geodf.to_file(f'{outdest}/{filename}.shp', driver='ESRI Shapefile')
            return f'Saved as {filename}.shp\nSCHEMA NOT USED: Please contact AKO', outdest
        except Exception as e2:
            print(e2)
            return f'Export failed. Error:\n{e1}\n{e2}', None

    return f'Saved as {filename}.shp', outdest

def export(file, last, cols, ftype): 
    #file=Screenlog #last=recordsToInclude #cols=columnsToParse #ftype=exportFiletype
    size = os.stat(file).st_size
    mb = int(size / 2**20)
    reading = QuasarFeedback(f'Export {ftype}', f'Exporting to {ftype}')
    reading.setStandardButtons(QMessageBox.NoButton)
    QApplication.processEvents()
    reading.show()
    # determine fields to parse
    if cols == -2: # all fields
        columns = None
        datetimes = True
    elif cols == -3: # quick/minimal
        columns = [
            'Seq',
            'Sample_ID',
            'E_planned',
            'N_planned',
            'Logger',
            'EOH_Code',
            'Status',
            'Stones'
        ]
        datetimes = False
    else:
        print(f'Unknown "export_type": got {cols}')

    desc = 'SamplingResults'
    # determine record to end at
    if last == -2:
        endat = 'Sample_ID'
    elif last == -3:
        endat = 'Status'
    elif last == -4:
        endat = 'Logger'
    elif last == -5:
        endat = 'EOH_Code'
        desc = 'ABD'
    else:
        print(f'Unknown "last_record": got {last}')

    # derive concession name:
    concesh = file.rsplit('_', 1)[1][:-5]
    filename = f'{concesh}_{desc}_{TODAY}'

    if QuasarExport.setdirs != 'True':
        if ftype == 'csv':
            outdest, filename = save_file('Save csv as...', filename, 'Comma-Separated Values (*.csv)')
        elif ftype == 'shp':
            outdest, filename = save_file('Save shp as...', filename, 'ESRI Shapefile (*.shp)')
    else:
        if ftype == 'csv':
            outdest = QuasarExport.csvdest
        elif ftype == 'shp':
            outdest = QuasarExport.shpdest
    if outdest is not None:
        reading.setText(f'Reading Screenlog ({mb}MB)...')
        QApplication.processEvents()
        outdata, warnings = read_sl(file, columns, datetimes, endat) # datetimes = Bool, to parse dates/times or not
        if outdata is None:
            report, folder = 'SL Read Error\nContact your AKO\n', None
        else:
            # EXPORT FILE:
            if ftype == 'csv':
                report, folder = exportcsv(outdata, outdest, filename)
            elif ftype == 'shp':
                reading.setText('Saving as Shapefile...')
                QApplication.processEvents()
                report, folder = exportshp(outdata, outdest, filename)
            else:
                report, folder = f'No filetype "{ftype}"', None
        if warnings is not None:
            report += f'Warning: {warnings}'
    else:
        report, folder = 'Cancelled', None
    reading.hide()
    # messagebox: Exported successfully
    success = QuasarFeedback(f'Export {ftype}', report)
    if folder is not None:
        success.setStandardButtons(QMessageBox.Ok | QMessageBox.Open)
        success.button(QMessageBox.Open).setText('Open Folder')
        ret = success.exec()
        if ret == QMessageBox.Open:
            os.startfile(folder)
    else:
        success.exec()

def read_sl(filename, columns, dtbool, endat): 
    print('\nReading Screenlog . . .')
    try:
        read_init = dt.datetime.now()
        sl = pd.read_excel(filename, engine='pyxlsb', usecols=columns)#, index_col=0)
        if dtbool == True:
            try:
                epoch = dt.datetime(1899, 12, 30)
                sl['Start_Date'] = pd.TimedeltaIndex(sl['Start_Date'], unit = 'd') + epoch
                sl['End_Date'] = pd.TimedeltaIndex(sl['End_Date'], unit = 'd') + epoch
                sl['Start_Date'] = sl['Start_Date'].dt.strftime('%Y-%m-%d')
                sl['End_Date'] = sl['End_Date'].dt.strftime('%Y-%m-%d')
                sl['Start_Time'] = pd.to_datetime(sl['Start_Time'], unit = 'd')
                sl['Start_Time'] = sl['Start_Time'].dt.strftime('%H:%M')
                sl['End_Time'] = pd.to_datetime(sl['End_Time'], unit = 'd')
                sl['End_Time'] = sl['End_Time'].dt.strftime('%H:%M')
            except Exception as d:
                print(d, '\nError parsing dates in Screenlog')
                return None, d
        #sl.set_index('Seq', drop=False, inplace=True)
        endindex = sl[endat].last_valid_index()
        sl = sl[sl.index <= endindex]
        if endat == 'EOH_Code': # filter for ABD/ABT
            sl = sl[sl['EOH_Code'].isin(['ABD', 'ABT'])]
        #dropcrap:
        for field in sl.columns:
            if 'Unnamed' in field:
                sl.drop(field, axis=1, inplace=True)
        #sl.dropna(how='all', axis=1, inplace=True)

        sl = sl.round(2)
        dur = dt.datetime.now() - read_init
        dur = dur.seconds + (dur.microseconds / 1000000)
        print('Screenlog processed in', round(dur, 2), 'seconds (', len(sl), ' records)')
        return sl, None
    except Exception as e:
        print(e, '\nError reading Screenlog')
        return None, e

with open(key, 'r') as k:
    if k.read() != 'AKO':
        sys.exit()

def save_cfg(section, key, value):
    config.set(section, key, value)
    with open('config.ini', 'w') as cfg:
        config.write(cfg)

# MAIN FUNCTION
def main():
    # Create an instance of QApplication
    qsr = QApplication(sys.argv)
    qsr.setStyle('Fusion')
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
    qsr.setPalette(palette)
    # Show the GUI
    quasar = QMainWindow()
    quasar.setWindowTitle('IMDH Quick Sampling Results v1')
    quasar.setWindowIcon(QIcon('icon.ico'))
    quasar.resize(400, 240)
    quasar.tabwidget = QuasarUI(quasar)
    quasar.setCentralWidget(quasar.tabwidget)
    quasar.show() # here or elsewhere

    # main loop
    sys.exit(qsr.exec_())

if __name__ == '__main__':
    main()