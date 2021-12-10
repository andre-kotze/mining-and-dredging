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
from PyQt5.QtWidgets import (
    QFileDialog,
    QApplication, 
    QMainWindow, 
    QWidget, 
    QLabel, 
    QPushButton,
    QMessageBox,
    QVBoxLayout,
    QHBoxLayout,
    QCheckBox
)
from PyQt5.QtGui import QIcon, QPalette, QColor

config = ConfigParser()
config.read('config.ini')
last_used = config.get('last_used', 'concession')
COLS=[
        'Seq',
        'Sample_ID',
        'Easting',
        'Northing',
        'Bathymetry',
        'SedThick',
        'Topas',
        'Feature'
    ]

class DPStarUI(QWidget):
    def __init__(self, parent):
        super(QWidget, self).__init__(parent)
        self.last_dp = config.get('last_used', 'dp')
        self.last_outdir = config.get('last_used', 'output')
        #derive concession from DP filename:
        self.conc = os.path.basename(self.last_dp).split('_', 1)[0]

        mlay = QVBoxLayout()
        dpline = QHBoxLayout()
        expline = QHBoxLayout()


        #show active files/dirs
        self.dp_lbl = QLabel(os.path.basename(self.last_dp))
        self.dp_lbl.setToolTip(self.last_dp)
        self.dp_bbtn = QPushButton('...', self)
        self.dp_bbtn.clicked.connect(self.selectfile)
        self.dp_bbtn.setFixedSize(20, 20)
        dpline.addWidget(QLabel('DP: '))
        dpline.addWidget(self.dp_lbl)
        dpline.addWidget(self.dp_bbtn)
        #derived conc name 
        self.conc_info = QLabel(f'Concession: {self.conc}')
        self.conc_info.setStyleSheet("QLabel{color: grey}")
        self.conc_info.setToolTip('Name derived from DP filename')
        #tickboxes DSL, GSL
        self.gsl_opt = QCheckBox('GSL')
        self.gsl_opt.setChecked(True)
        self.dsl_opt = QCheckBox('DSL')
        self.dsl_opt.setChecked(True)
        self.exp_btn = QPushButton('Export')
        self.exp_btn.clicked.connect(self.export)
        expline.addWidget(self.gsl_opt)
        expline.addWidget(self.dsl_opt)
        expline.addWidget(self.exp_btn)

        mlay.addLayout(dpline)
        mlay.addWidget(self.conc_info)
        mlay.addLayout(expline)
        self.setLayout(mlay)

    def selectfile(self):
        file = QFileDialog().getOpenFileName(None, 'Open Drilling Program xlsx', os.path.dirname(self.last_dp))[0]
        if len(file) != 0 and file != self.last_dp:
            self.last_dp = file
            self.conc = os.path.basename(self.last_dp).split('_', 1)[0]
            self.conc_info.setText(f'Concession: {self.conc}')
            self.dp_lbl.setText(os.path.basename(self.last_dp))
            self.dp_lbl.setToolTip(self.last_dp)
            config.set('last_used', 'dp', self.last_dp)
            with open('config.ini', 'w') as f:
                config.write(f)

    def selectdir(self):
        dirname = QFileDialog().getExistingDirectory(None, 'Select Output Directory', self.last_outdir)
        if len(dirname) != 0 and dirname != self.last_outdir:
            self.last_outdir = dirname
            #self.outdir_lbl.setText(os.path.basename(self.last_dp)) # last branch in path
            #self.outdir_lbl.setToolTip(self.last_outdir)
            config.set('last_used', 'output', self.last_outdir)
            with open('config.ini', 'w') as f:
                config.write(f)
            return dirname
        else:
            return None
            

    def export(self):
        report = []
        folder = None
        if self.gsl_opt.isChecked() or self.dsl_opt.isChecked():
            folder = self.selectdir()
            dpdf = pd.read_excel(self.last_dp)#, usecols=COLS)
            for col in COLS:
                if col not in list(dpdf.columns.values):
                    report.append(f'No column "{col}"')
                    pass # handle here
            base_name = os.path.basename(self.last_dp)[:-5] # get dp name template
            if self.gsl_opt.isChecked():
                dpdf['QR_Code'] = self.conc + '_GSL_' + dpdf['Sample_ID']
                gsl_fname = self.last_outdir + '/' + base_name + '_GSL.csv'
                dpdf.to_csv(gsl_fname, index=False)
                report.append(f'Saved as {os.path.basename(gsl_fname)}')
            if self.dsl_opt.isChecked():
                dpdf['QR_Code'] = self.conc + '_DSL_' + dpdf['Sample_ID']
                dsl_fname = self.last_outdir + '/' + base_name + '_DSL.csv'
                dpdf.to_csv(dsl_fname, columns=['Seq', 'Sample_ID', 'QR_Code'], index=False)
                report.append(f'Saved as {os.path.basename(dsl_fname)}')
        else:
            report.append('No formats checked, no output created')
        # message report:
        msg = QMessageBox(QMessageBox.NoIcon, 'Export Data Source', '\n'.join(report))
        msg.setWindowIcon(QIcon('icon.ico'))
        msg.setStandardButtons(QMessageBox.Ok)
        if folder != None:
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Open)
            msg.button(QMessageBox.Open).setText('Open Folder')
        ret = msg.exec()
        if ret == QMessageBox.Open:
            os.startfile(folder)


# MAIN FUNCTION
def main():
    # Create an instance of QApplication
    dpstar = QApplication(sys.argv)
    dpstar.setStyle('Fusion')
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
    dpstar.setPalette(palette)
    # Show the GUI
    dpui = QMainWindow()
    dpui.setWindowTitle('DP⛧ DrillingProgram Star v2')
    dpui.setWindowIcon(QIcon('icon.ico'))
    dpui.ui = DPStarUI(dpui)
    dpui.setCentralWidget(dpui.ui)
    dpui.show()
    # main loop
    sys.exit(dpstar.exec_())

if __name__ == '__main__':
    main()