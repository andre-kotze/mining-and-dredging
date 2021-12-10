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

import fiona
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
    QProgressDialog,
    QCheckBox,
    QGridLayout,
    QFrame
)
from PyQt5.QtGui import QIcon, QPalette, QColor, QPainter, QPixmap, QTransform
from PyQt5.QtCore import QThread, pyqtSignal, QPropertyAnimation, QObject, pyqtProperty, Qt

__version__ = '1.1.1'
__author__ = 'AKO_Geo'

TODAY = dt.datetime.now().date().strftime('%Y%m%d')
config = ConfigParser()
config.read('config.ini')
#increment = config.get('last_used', 'increment') 
    # True: auto-increment filenames (overwrite = False)

# look into appending onto existing shps

read_stat = ''
# records -2:All
# records -3:Last Sorted
# records -2:Last Logged
# columns -2:Full
# columns -3:Minimal

class Loader(QProgressDialog):
    def __init__(self, desc, limit):
        super().__init__()
        self.setWindowIcon(QIcon('icon.ico'))
        self.setWindowModality(Qt.WindowModal)
        self.setWindowTitle('Quasar Export')
        self.setMinimum(0)
        self.setCancelButtonText('Cancel')
        self.setMinimumDuration(100)

        self.setMaximum(limit)
        self.setLabelText(desc)

class Worker(QObject):
    done = pyqtSignal()

    def __init__(self, arglist, parent=None):
        super().__init__(parent)
        # readargs = [filename, columns, dtbool, endat ]
        self.arglist = arglist

    def readfile(self):
        print('Worker start reading')
        Quasar.sldf = read_sl(self.arglist)
        self.done.emit()
        print(f'Quasar.sldf:\ntype: {type(Quasar.sldf)}\nlen: {len(Quasar.sldf)}')
        print("done")

    def exportshpfile(self):
        print('Worker start exporting')
        exportshp(self.arglist)
        self.done.emit()
        print('done')


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
class Quasar(QWidget):
    csvdest = config.get('last_used', 'csvdest')
    shpdest = config.get('last_used', 'shpdest')
    setdirs = config.get('last_used', 'setdirs') # True: save folders are set (choose = False)
    lddsl = config.get('last_used', 'file')
    use_schema = config.get('last_used', 'schema')
    last_results = config.get('last_used', 'results')
    sldf = None
        #Quasar.append_set = config.get('last_used', 'append')
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        
        #Quasar.increment = increment
        self.csvdir = QLabel(Quasar.csvdest)
        self.shpdir = QLabel(Quasar.shpdest)
        self.truncate_label(Quasar.csvdest, self.csvdir)
        self.truncate_label(Quasar.shpdest, self.shpdir)

        # Init tab screen
        self.tabs = QTabWidget(self)
        self.apptab = QuasarAppend(self)
        self.exptab = QuasarExport(self)
        self.settab = QWidget(self)
        self.abouttab = QWidget(self)

        # Add tabs
        self.tabs.addTab(self.apptab, 'Append')
        self.tabs.addTab(self.exptab, 'Export')
        self.tabs.addTab(self.settab, 'Settings')
        self.tabs.addTab(self.abouttab, 'About')

        self.abouttab.layout = QVBoxLayout(self)
        self.abouttext = QLabel(
            '''QUASAR Quick Sampling Results\n\nVersion 1.1.4\n
            Written 2021 by Andrew "Drikus" Kotez\n
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
        if Quasar.setdirs == 'True':
            self.btnauto.setChecked(True)
        else:
            self.btnchoose.setChecked(True)
        
        self.saveoptgroup = QButtonGroup()
        self.saveoptgroup.addButton(self.btnchoose)
        self.saveoptgroup.addButton(self.btnauto)

        shp_opts = QHBoxLayout()
        self.shp_schema = QCheckBox('Use schema')
        self.shp_schema.setToolTip('Use shapefile schema for optimal field lengths (smaller filesize)')
        if Quasar.use_schema == 'True':
            self.shp_schema.setChecked(True)
        self.shp_schema.toggled.connect(self.toggle_schema)

        shp_opts.addWidget(QLabel('SHP export options:'))
        shp_opts.addWidget(self.shp_schema)

        mainlayout.addWidget(title)
        mainlayout.addWidget(self.btnchoose)
        mainlayout.addWidget(self.btnauto)
        mainlayout.addLayout(destdircontainer)
        mainlayout.addLayout(shp_opts)
        mainlayout.addStretch(0)
        return mainlayout
        
    def createDestGroup(self):
        groupBox = QGroupBox('Destination directories')

        self.csvbtn = QPushButton('...', self)
        self.csvbtn.clicked.connect(lambda: self.select_dir('Select csv output directory', 'csvdest'))
        self.csvbtn.setFixedWidth(22)
        self.shpbtn = QPushButton('...', self)
        self.shpbtn.clicked.connect(lambda: self.select_dir('Select shp output directory', 'shpdest'))
        self.shpbtn.setFixedWidth(22)

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
                Quasar.setdirs = 'False'
            else:
                self.destdirgroup.setDisabled(False)
                #self.destdirgroup.setStyleSheet('QGroupBox:title {color: white}; QLabel {color: white;}; QPushButton {color: white;}')
                Quasar.setdirs = 'True'
        elif btn.text() == 'Set save directories:': # Set dirs
            if btn.isChecked():
                self.destdirgroup.setDisabled(False)
                #self.destdirgroup.setStyleSheet('QGroupBox:title {color: white}; QLabel {color: white;}; QPushButton {color: white;}')
                Quasar.setdirs = 'True'
            else:
                self.destdirgroup.setDisabled(True)
                #self.destdirgroup.setStyleSheet('QGroupBox:title {color: grey}; QLabel {color: grey;}; QPushButton {color: grey;}')
                Quasar.setdirs = 'False'
        else:
            print(f'No button with id {btn.id()}')
            # Set save directories
        save_cfg('last_used', 'setdirs', Quasar.setdirs)

    def toggle_schema(self):
        if self.shp_schema.isChecked():
            Quasar.use_schema = 'True'
        else:
            Quasar.use_schema = 'False'
        save_cfg('last_used', 'schema', Quasar.use_schema)

    # function to truncate long paths and add tooltip
    def truncate_label(self, destobj, dirlabel): # destobj = Quasar.xxxdest (to be measured), dirlabel = self.xxxdir
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
                Quasar.csvdest = dirname
                self.truncate_label(Quasar.csvdest, self.csvdir)
            elif ftypedest == 'shpdest':
                Quasar.shpdest = dirname
                self.truncate_label(Quasar.shpdest, self.shpdir)
            else:
                QuasarFeedback('ERROR', f'Unknown ftypedest: {ftypedest}')

    def change_lddsl(self):
        file = QFileDialog.getOpenFileName(None, 'Select Screenlog xlsb', os.path.dirname(Quasar.lddsl), 'Excel Binary Workbook (*.xlsb)')[0]
        if len(file) != 0 and file != Quasar.lddsl:
            Quasar.lddsl = file
            self.exptab.csl.setText(os.path.basename(file))
            self.apptab.csl.setText(os.path.basename(file))
            self.exptab.csl.setToolTip(file)
            self.apptab.csl.setToolTip(file)
            save_cfg('last_used', 'file', file)
            save_cfg('last_used', 'folder', os.path.dirname(file) + '/')

key = 'C:/Users/User/Desktop/SHARE/Andre/KEY.txt'
# Main function: Export tab
class QuasarExport(QWidget): #TableWidget
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)

        # Loaded SL label
        self.csl = QLabel(os.path.basename(Quasar.lddsl), self)
        self.csl.setToolTip(Quasar.lddsl)
        self.bbtn = QPushButton('...', self)
        self.bbtn.clicked.connect(parent.change_lddsl)
        self.bbtn.setFixedWidth(22)
        self.savedfolder = None

        # export csv button
        self.csvbtn = QPushButton('Export csv', self)
        self.csvbtn.clicked.connect(lambda: self.lookbusy('csv'))

        # export shape button
        self.shpbtn = QPushButton('Export shp', self)
        self.shpbtn.clicked.connect(lambda: self.lookbusy('shp'))

        # box containing label
        r1 = QHBoxLayout()
        r1.addWidget(QLabel('Loaded:'))
        r1.addWidget(self.csl)
        r1.addStretch(1)
        r1.addWidget(self.bbtn)
        #r1.addWidget(self.readbtn) # TEMP

        # box containing records/fields boxes
        r2 = QHBoxLayout()
        self.recgroup = QButtonGroup()
        self.colgroup = QButtonGroup()
        r2.addWidget(self.createRecordGroup())
        r2.addWidget(self.createFieldGroup())
        '''
        r3 = QHBoxLayout()
        self.shp_append = QCheckBox('Append:')
        if Quasar.append_set == 'True':
            self.shp_append.setChecked(True)
        self.shp_append.toggled.connect(self.toggle_append)
        self.ldd_results = QLabel(os.path.basename(self.last_results))
        self.results_bbtn = QPushButton('...', self)
        self.results_btn.clicked.connect(self.change_results)
        self.results_btn.setFixedSize(20, 20)
        '''
        # box containing export buttons
        r4 = QHBoxLayout()
        r4.addWidget(self.csvbtn)
        r4.addWidget(self.shpbtn)

        # box containing boxes
        vbox = QVBoxLayout()
        vbox.addLayout(r1)
        vbox.addLayout(r2)
        #vbox.addLayout(r3)
        vbox.addStretch(1)
        vbox.addLayout(r4)

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

    def lookbusy(self, ftype):
        self.csvbtn.setDisabled(True)
        self.shpbtn.setDisabled(True)
        QApplication.processEvents()
        self.export(Quasar.lddsl, self.recgroup.checkedId(), self.colgroup.checkedId(), ftype)
        self.csvbtn.setEnabled(True)
        self.shpbtn.setEnabled(True)

    def export(self, file, last, cols, ftype):
        global read_stat
        #file=Screenlog #last=recordsToInclude #cols=columnsToParse #ftype=exportFiletype
        size = os.stat(file).st_size
        mb = round((size / 2**20), 2)
        dur = round(mb*(0.1*mb + 1.48),1)
        read_stat = f'\n{os.path.basename(file)}, {mb}, '
        est_read = int(dur * 2)
        print(f'est_read = {dur} * 2 = {est_read} iterations')
        est_toshp = int(dur * 8)
        print(f'est_toshp = {dur} * 8 = {est_toshp} iterations')
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
            print(f'Unknown "export_type" (╯°□°)╯͡ ┻━┻ got: {cols}')

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

        if Quasar.setdirs != 'True':
            if ftype == 'csv':
                outdest, filename = save_file('Save csv as...', filename, 'Comma-Separated Values (*.csv)')
            elif ftype == 'shp':
                outdest, filename = save_file('Save shp as...', filename, 'ESRI Shapefile (*.shp)')
        else:
            if ftype == 'csv':
                outdest = Quasar.csvdest
            elif ftype == 'shp':
                outdest = Quasar.shpdest
        if outdest is None:
            report, folder = 'Cancelled', None
        else:
            self.worker = Worker([file, columns, datetimes, endat]) # no parent, so it can be moved. arglist = [filename, columns, dtbool, endat]
            self.external = QThread(self)
            self.worker.moveToThread(self.external)
            self.read_prog = Loader('Reading Screenlog...', est_read)
            self.worker.done.connect(self.read_prog.cancel)

            # OLD LINE: START WORKER
            #self.worker.readfile([file, columns, datetimes, endat])
            # NEW LINE: TRIGGER READ WHEN THREAD STARTED
            self.external.started.connect(self.worker.readfile)
            self.external.start()

            count = 0
            while not self.read_prog.wasCanceled():
                count += 1
                time.sleep(0.5)
                #rprog = (est_read * count) / (count + (est_read/2))
                self.read_prog.setValue(count)#int(rprog))
            #self.read_prog.cancel()
            print(' thread exit')
            self.external.exit()
            print(type(Quasar.sldf))
            if Quasar.sldf is None:
                print('empty df received')
                report, folder = 'SL Read Error\n(╯°□°)╯͡ ┻━┻\n', None
            else:
                # EXPORT FILE:
                if ftype == 'csv':
                    report, folder = exportcsv(Quasar.sldf, outdest, filename)
                elif ftype == 'shp':
                    self.worker = Worker([Quasar.sldf, outdest, filename]) # no parent, so it can be moved
                    self.external = QThread(self)
                    self.worker.moveToThread(self.external)
                    self.write_prog = Loader('Saving Shapefile...', est_toshp)
                    self.worker.done.connect(self.write_prog.cancel)
                    self.external.started.connect(self.worker.exportshpfile)
                    self.external.start()
                    count = 0
                    while not self.write_prog.wasCanceled():
                        count += 1
                        #prog = (est_toshp * count) / (count + (est_toshp/2))
                        time.sleep(0.5)
                        self.write_prog.setValue(count)#int(prog))
                    #self.write_prog.cancel()
                    print(' thread exit')
                    self.external.exit()
                    if outdest is None:
                        print('No feedback received')
                        report, folder = 'SHP Export Error\n(╯°□°)╯͡ ┻━┻\n', None
                    else:
                        report, folder = 'Success!', outdest
                else:
                    report, folder = f'No filetype "{ftype}"', None
        with open('read_time_stats.csv', 'a') as stats:
            stats.write(read_stat)
        success = QuasarFeedback(f'Export {ftype}', report)
        if folder is not None:
            success.setStandardButtons(QMessageBox.Ok | QMessageBox.Open)
            success.button(QMessageBox.Open).setText('Open Folder')
            ret = success.exec()
            if ret == QMessageBox.Open:
                os.startfile(folder)
        else:
            success.exec()

class QuasarAppend(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)

        sl_head = QLabel('Screenlog:')
        sl_head.setMaximumHeight(24)
        self.csl = QLabel(os.path.basename(Quasar.lddsl), self)
        self.csl.setToolTip(Quasar.lddsl)
        self.csl.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        self.sl_bbtn = QPushButton('...', self)
        self.sl_bbtn.clicked.connect(parent.change_lddsl)
        self.sl_bbtn.setFixedWidth(22)

        sr_head = QLabel('Sampling Results:')
        sr_head.setMaximumHeight(24)
        self.results = QLabel(os.path.basename(Quasar.last_results))
        self.results.setToolTip(Quasar.last_results)
        self.results.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        self.results_bbtn = QPushButton('...', self)
        self.results_bbtn.clicked.connect(self.change_results)
        self.results_bbtn.setFixedWidth(22)

        self.append_btn = QPushButton('Append', self)
        self.append_btn.clicked.connect(self.append_to_shp)
        self.append_btn.setMaximumWidth(120)

        lout = QGridLayout()
        lout.addWidget(sl_head,0,0,1,1)
        lout.addWidget(self.csl,1,0,1,1)
        lout.addWidget(self.sl_bbtn,1,1,1,1)
        lout.addWidget(sr_head,2,0,1,1)
        lout.addWidget(self.results,3,0,1,1)
        lout.addWidget(self.results_bbtn,3,1,1,1)
        lout.addWidget(self.append_btn,4,0,1,2)
        lout.setRowStretch(4,1)

        self.setLayout(lout)

    def change_results(self):
        file = QFileDialog.getOpenFileName(None, 'Select Sampling Results shp', os.path.dirname(Quasar.last_results), 'ESRI Shapefile (*.shp)')[0]
        if len(file) != 0 and file != Quasar.last_results:
            Quasar.last_results = file
            self.results.setText(os.path.basename(file))
            self.results.setToolTip(file)
            save_cfg('last_used', 'results', file)

    def append_to_shp(self, slfile, srfile):
        size = os.stat(slfile).st_size
        mb = round((size / 2**20), 2)
        dur = round(mb*(0.1*mb + 1.48),1)

        est_read = int(dur * 2)
        print(f'est_read = {dur} * 2 = {est_read} iterations')
        #est_toshp = int(dur * 8)
        #print(f'est_toshp = {dur} * 8 = {est_toshp} iterations')
        outdest, filename = save_file('Save shp as...', filename, 'ESRI Shapefile (*.shp)')

        self.worker = Worker([slfile, None, True, 'Status']) # no parent, so it can be moved. arglist = [filename, columns, dtbool, endat]
        self.external = QThread(self)
        self.worker.moveToThread(self.external)
        self.read_prog = Loader('Reading Screenlog...', est_read)
        self.worker.done.connect(self.read_prog.cancel)
        self.external.started.connect(self.worker.readfile)
        self.external.start()
        count = 0
        while not self.read_prog.wasCanceled():
            count += 1
            time.sleep(0.5)
            self.read_prog.setValue(count)
        print(' thread exit')
        self.external.exit()
        print(type(Quasar.sldf))
        if Quasar.sldf is None:
            print('empty df received')
            report, folder = 'SL Read Error\n(╯°□°)╯͡ ┻━┻\n', None
        else:
            # EXPORT FILE:
            with fiona.open(srfile) as sr_shp:
            return

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
    if os.path.exists(f'{outdest}/{filename}.csv') and Quasar.setdirs == 'True':
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

def exportshp(saveargs):
    global read_stat
    df, outdest, filename = saveargs
    shp_init =  dt.datetime.now()
    shp_schema = get_schema(df.columns)
    df['geometry'] = df.apply(lambda x: Point((x.E_planned, x.N_planned)), axis=1)
    geodf = gpd.GeoDataFrame(df, crs='EPSG:32733', geometry='geometry')
    if os.path.exists(f'{outdest}/{filename}.shp') and Quasar.setdirs == 'True':
        n = 1
        newname = f'{filename}_{n}'
        while os.path.exists(f'{outdest}/{filename}_{n}.shp'):
            n += 1
            newname = f'{filename}_{n}'
        filename = newname
    try:
        geodf.to_file(f'{outdest}/{filename}.shp', driver='ESRI Shapefile', schema=shp_schema) # catch value and key error
        msg = f'Saved as {filename}.shp'
    except Exception as e1:
        print(e1)
        try:
            geodf.to_file(f'{outdest}/{filename}.shp', driver='ESRI Shapefile')
            msg = f'Saved as {filename}.shp\nSCHEMA NOT USED: Please contact AKO'
        except Exception as e2:
            print(e2)
            msg = f'Export failed. Error:\n{e1}\n{e2}'
            return None
    dur = dt.datetime.now() - shp_init
    dur = dur.seconds + (dur.microseconds / 1000000)
    print(f'Shapefile created in {round(dur, 2)} seconds')
    read_stat += f'{round(dur, 2)}'

def read_sl(readargs):
    filename, columns, dtbool, endat = readargs
    global read_stat
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
                QuasarFeedback('Epic Fail', 'Error parsing dates in Screenlog', [d])
                return None
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
        read_stat += f'{round(dur, 2)}, '
        return sl
    except Exception as e:
        print(e, '\nError reading Screenlog')
        QuasarFeedback('Epic Fail', 'Error reading Screenlog', [e])
        return None

with open(key, 'r') as k:
    if k.read() != 'AKO':
        sys.exit()

def save_cfg(section, key, value):
    config.set(section, key, value)
    with open('config.ini', 'w') as cfg:
        config.write(cfg)

# MAIN FUNCTION
#def main():
    

if __name__ == '__main__':
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
    mainwidget = Quasar(quasar)
    quasar.setCentralWidget(mainwidget)
    quasar.show() # here or elsewhere

    # main loop
    sys.exit(qsr.exec_())