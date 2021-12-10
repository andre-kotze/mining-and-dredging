import sys
from PyQt5 import QtCore, QtGui, QtWidgets

app = QtWidgets.QApplication(sys.argv)
dialog = QtWidgets.QFileDialog()


print('choose dir . . .')

name = dialog.getExistingDirectory(None, 'Select Output Directory')
print('dir received as:\n ', name)


print('length:', len(name))
