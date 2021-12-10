# --- import the required modules
import os
import shutil
import sys
#import socket

import eel
import cv2
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog

print('RAZZ BLAMMY MITAZZY LOADING . . .')
eel.init("web")

logdirs = {'GSL':'Geological Screen Logs/', 'DSL':'Diamond Sorting Logs/', 'GeoT':'Geotech Logs/'}
detector = cv2.QRCodeDetector()

path = '//192.168.0.5/Geology/scans/' # path to scans
ospath = '//192.168.0.190/SHARE/' # path to SHARE

#host = socket.gethostname()
#if host.upper() == 'GEOTECH-PC':
#    ospath = 'Z:/'

csv = 'ScanLogPaths.csv'

# executed when setting screen loaded
@eel.expose
def gettable():
    cpaths = pd.read_csv(csv)
    conclist = '<table class="table" border=1>'
    for c in cpaths.itertuples():  # need to open filedialog to select screenlog
        conclist += f'<tr><td style="width: 40%">{c.Concession}</td><td style="width: 60%">{c.Path}</td></tr>'
    conclist += '</table>'
    return conclist

# executed on homescreen load ... what?
@eel.expose
def refresh_concs():
    global cpaths
    cpaths = pd.read_csv(csv)

# --- count the number of .jpg files in the scans folder
@eel.expose
def detect():
    global jpglist
    scans = os.listdir(path)
    jpglist = []
    jpgnames = '<p>'
    for f in scans:
        if '.jpg' in f:
            jpglist.append(f)
            jpgnames += f + '<br>'
    jpgnames += '</p>'
    if len(jpglist) == 0:
        return {'finds':'<p>No Files Found</p>', 'count':0}
    else:
        return {'finds':jpgnames, 'count':len(jpglist)}

@eel.expose
def conc_options(): # for delete conc
    cpaths = pd.read_csv(csv)
    conclist = ''
    for c in cpaths.itertuples():
        conclist += f'<option name={c.Concession}>{c.Concession}</option>'
    return conclist

@eel.expose
def killprogram():
    sys.exit(0)

# --- decode and rename function, called by "preprocess"
def scanrename(array, file):
    try:
        data = detector.detectAndDecode(array)[0]
        if data != '':
            ScanName = data + '.jpg'
            os.rename(path + file, path + ScanName)
            if ScanName not in successlist:
                successlist.append(ScanName)
            return True
        else:
            return False
    except:
        return False
    
# --- image preprocessing function, called by "rename"
def preprocess(img, file):
    crop = img[100:480,40:600]
    if scanrename(crop, file) == False: # qr crop attempt
        scales = [0.5, 0.9, 0.3, 0.6, 0.4, 0.8, 0.7, 1]
        for scale in scales: # qr cropped resized
            crpimg = cv2.resize(crop, (int(380 * scale), int(560 * scale)))
            if scanrename(crpimg, file):
                return True
        for scale in scales: # full img resize
            array = cv2.resize(img, (int(2480 * scale), int(3507 * scale)))
            if scanrename(array, file):
                return True
        return False
    else:
        return True # true if initial attempt != False

# 1.2.1 fullimgresize will be skipped, only called from rotate
def fullimgresize(img, file, scales):
    for scale in scales:
        fi_resize = cv2.resize(img, (int(2480 * scale), int(3507 * scale)))
        if scanrename(fi_resize, file):
            return True
    return False

def imggauss(img, file, kernels):
    for kernel in kernels:
        gaussimg = cv2.GaussianBlur(img,(kernel,kernel),0)
        if scanrename(gaussimg, kernel):
            return True
    return False

def imgmedian(img, file, kernels):
    for kernel in kernels:
        medianimg = cv2.medianBlur(img,kernel)
        if scanrename(medianimg, kernel):
            return True
    return False

def imgbilateral(img, file, kernels):
    for kernel in kernels:
        bilateral = cv2.bilateralFilter(img,kernel,75,75)
        if scanrename(bilateral, kernel):
            return True
    return False

def imgrotate(img, file, scales, angles, kernels):
    pivot = tuple(np.array(img.shape[1::-1]) / 2)
    for angle in angles:
        rot_mat = cv2.getRotationMatrix2D(pivot, angle, 1.0)
        rotated = cv2.warpAffine(img, rot_mat, img.shape[1::-1], flags=cv2.INTER_LINEAR)
        if fullimgresize(rotated, file, scales):
            return True
        elif imgmedian(rotated, file, kernels):
            return True
        elif imggauss(rotated, file, kernels):
            return True
        elif imgbilateral(rotated, file, kernels):
            return True
        else:
            pass
    return False

# --- image preprocessing function, called by "bruteforce"
def try_harder(img, file):
    scales = [1, 0.5, 0.9, 0.3, 0.6, 0.4, 0.8, 0.7]#, 0.45, 0.75, 0.666] 
    kernels = [7, 5, 9, 3]#, 11] # added 11
    angles = [1, -1, 2, -2, 3, -3]#, 4, -4] # added 4
    #attempt = [
    #    "Gaussian blur", 
    #    "median blur", 
    #    "bilateral filter", 
    #    "rotated image"
    #    ]
    crop = img[100:480,40:600]
    # whole 9 yards for cropped codes
    for array in [crop, img]:
        for scale in scales:
            pimg = cv2.resize(array, (int(380 * scale), int(560 * scale)))
            if imggauss(pimg, file, kernels):
                return True
            elif imgmedian(pimg, file, kernels):
                return True
            elif imgbilateral(pimg, file, kernels):
                return True
            elif imgrotate(pimg, file, scales, angles, kernels):
                return True
            else:
                pass
    return False

@eel.expose
def rename():
    global successlist
    successlist = []
    renamedfiles = '<p>'
    for p, file in enumerate(jpglist):
        if 'scan' in file:
            img = cv2.imread(path + file)
            if not preprocess(img, file):
                print('Could not read', file)
        else:
            successlist.append(file)
        completion = str(int((p+1) / len(jpglist) * 100)) + '%'
        print(f'Rename progress: {completion}')
        eel.jupdateprog(completion)
    for s in successlist:
        renamedfiles += s + '<br>'
    renamedfiles += '</p>'
    return {'renames':renamedfiles, 'count':len(successlist)}

@eel.expose
def bruteforce():
    global successlist
    successlist = []
    renamedfiles = '<p>'
    for p, file in enumerate(jpglist):
        if 'scan' in file:
            img = cv2.imread(path + file)
            if not try_harder(img, file):
                print('Could not read', file)
        else:
            successlist.append(file)
        completion = str(int((p+1) / len(jpglist) * 100)) + '%'
        print(f'Rename progress: {completion}')
        eel.jupdateprog(completion)
    for s in successlist:
        renamedfiles += s + '<br>'
    renamedfiles += '</p>'
    return {'renames':renamedfiles, 'count':len(successlist)}

def move_log(scan):
    #cpaths = pd.read_csv(csv)
    try:
        target, logtype = scan.split('_')[0:2] #read concession from filename
    except ValueError as v:
        print(v)
        result = f'Could not move {scan} (could not unpack filename)'
        return result
    try:
        movepath = ospath + cpaths.loc[cpaths['Concession'] == target, 'Path'].iloc[0] + 'Scanned Logs/'
    except IndexError as e:
        print(e)
        result = f'Could not move {scan} ("{target}" is not a known concession)'
        return result
    except KeyError as e:
        print(e)
        result = f'Could not move {scan} (valid concession name not found in scan name)'
        return result
    try:
        destpath = movepath + logdirs.get(logtype)
        if not os.path.exists(destpath + scan):
            shutil.move(path + scan, destpath + scan)
            result = f'Moved to {target} {logtype}'
            return result
        else:
            result = f'File already exists in {target} {logtype}'
            return result
    except KeyError as e:
        print(e)
        result = f'Could not move {scan} (log type (GSL, DSL, GeoT) not found in scan name)'
        return result

@eel.expose
def movefiles():
    global movedlist
    movedlist = []
    movedfiles = '<p style = "margin : 0;">'
    for scan in successlist:
        moveresult = move_log(scan)
        movedfiles += moveresult + '<br>'
        if 'Moved to' in moveresult:
            movedlist.append(moveresult)
    movedfiles += '</p>'
    return {'moves':movedfiles, 'count':len(movedlist)}

def select_path():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(title='Select Concession Folder') + '/'
    path = path.split('SHARE/')[1]
    return path

@eel.expose
def add_conc(name):
    cpaths = pd.read_csv(csv)
    path = select_path()
    if name in cpaths['Concession'].tolist():
        message = 'Concession "' + name + '" already exists'
    elif len(name) == 0 or len(path) == 0:
        message = 'Field omitted'
    elif not os.path.exists(ospath + path):
        message = 'Invalid path'
    else:
        updated = cpaths.append({'Concession':name, 'Path':path}, ignore_index = True)
        updated.to_csv('ScanLogPaths.csv', index = False)
        message = 'Changes Saved'
    newtable = gettable()
    return {'table': newtable, 'msg': message}


@eel.expose
def remove_conc(name):
    print('removing concession', name)
    cpaths = pd.read_csv(csv)
    if name in cpaths['Concession'].tolist():
        updated = cpaths[cpaths['Concession'] != name]
        message = name + ' removed'
        updated.to_csv('ScanLogPaths.csv', index = False)
    newtable = gettable()
    print(message)
    return {'table': newtable, 'msg': message}

print('\nINITIALIZING WEB APP . . .')
eel.start("home.html")