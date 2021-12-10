import os
import sys
import time
import shutil
from datetime import (datetime, timedelta)
import tkinter as tk
from tkinter import filedialog

import pandas as pd
import numpy as np
import piexif
import cv2
import eel
from exif import (Image as exif, DATETIME_STR_FORMAT)
from PIL import Image, ImageDraw, ImageFont
from configparser import ConfigParser

print('BEEP BEEP BOOP LOADING . . .')
eel.init("web")

#settings = 'C:/Users/User/Desktop/SHARE/IMDH App Settings/PhotoRename/'
csv = 'concessions.csv'
x , y = 800, 600 
config = ConfigParser()
config.read('config.ini')
last_used = config.get('last_used', 'concession')

@eel.expose
def selectpath(): # for add/edit, reselect ppath and spath
    root = tk.Tk()
    root.withdraw()
    ppath = filedialog.askdirectory(title='Select Photos Folder') + '/'
    spath = filedialog.askopenfilename(title='Select Screenlog File')
    return ppath, spath

# executed when 'detect' is run
def change_conc(newly_selected):
    config.set('last_used', 'concession', newly_selected)
    with open('config.ini', 'w') as f:
        config.write(f)

@eel.expose
def killprogram():
    sys.exit(0)

@eel.expose
def lastpath():
    cpaths = pd.read_csv(csv)
    if last_used in cpaths['Concession'].tolist():
        path = cpaths.loc[cpaths['Concession'] == last_used, 'PathToPhotos'].iloc[0]
        return path
    else:
        return 'No Concession Selected'

# on start, lists conc options. Last used is preselected
@eel.expose
def conc_options():
    cpaths = pd.read_csv(csv)
    conclist = ''
    for c in cpaths['Concession'].tolist():
        if c == last_used:
            conclist += f'<option name={c} selected>{c}</option>'
        else:
            conclist += f'<option name={c}>{c}</option>'
    return conclist

# executed when setting screen loaded
@eel.expose
def gettable():
    cpaths = pd.read_csv(csv)
    conclist = '<table class="table" border=1>'
    for c in cpaths.itertuples():  # need to open filedialog to select screenlog
        conclist += f'<tr><td style="width: 10%">{c.Concession}</td><td style="width: 30%">{c.PathToPhotos}</td><td style="width: 60%">SHARE{c.PathToScreenlog.split("SHARE")[1]}</td></tr>'
    conclist += '</table>'
    return conclist

@eel.expose
def add_conc(name):
    cpaths = pd.read_csv(csv)
    if name in cpaths['Concession'].tolist():
        message = 'Concession "' + name + '" already exists'
    else:
        ppath, spath = selectpath()
        updated = cpaths.append({'Concession':name, 'PathToPhotos':ppath, 'PathToScreenlog':spath}, ignore_index = True)
        updated.to_csv('concessions.csv', index = False)
        message = name + ' added'
    newtable = gettable()
    print(message)
    return {'table': newtable, 'msg': message}

@eel.expose
def remove_conc(name):
    cpaths = pd.read_csv(csv)
    if name in cpaths['Concession'].tolist():
        updated = cpaths[cpaths['Concession'] != name]
        message = name + ' removed'
        updated.to_csv('concessions.csv', index = False)
    newtable = gettable()
    print(message)
    return {'table': newtable, 'msg': message}



# GENBOARD dims:
def fontsize(length):
    fs = -10 * length +220
    return fs

# FIND EARLIEST DATETIME IN PHOTOS, AND CREATE DF WITH IMAGE FILES AND DATESTAKEN
@eel.expose
def detect(selectedconcession): # check_photos
    global photodf, earliest, sl_file
    global concession, path
    concession = selectedconcession
    datelist, photoslist = [], []
    cpaths = pd.read_csv(csv)
    path = cpaths.loc[cpaths['Concession'] == selectedconcession, 'PathToPhotos'].iloc[0]
    sl_file = cpaths.loc[cpaths['Concession'] == selectedconcession, 'PathToScreenlog'].iloc[0]
    sl_label_list = sl_file.split('/')[-4:]
    sl_file_label = '/'.join(sl_label_list)

    files = os.listdir(path)
    if len(files) == 0:
        return {'num':0,'path':path, 'sl_file':sl_file_label}
    for f in files:
        if '.JPG' in f:
            photoslist.append(f)
            with open(path + f, 'rb') as img:
                image = exif(img)
                td = image.datetime
                datetaken = datetime.strptime(td,"%Y:%m:%d %H:%M:%S")
                datelist.append(datetaken)
    if len(photoslist) == 0:
        return {'num':0,'path':path, 'sl_file':sl_file_label}
    data = {'Photo':photoslist, 'Date_Taken':datelist}
    photodf = pd.DataFrame(data)
    # CONVERT DATETAKENS FROM TIMESTAMP TO PY_DATETIME FORMAT
    photodf.Date_Taken = photodf.Date_Taken.dt.to_pydatetime()
    earliest = min(datelist) # datetime of oldest photo
    earliest = datetime.date(earliest) # date of oldest photo
    print(f'Earliest date: {earliest}')
    change_conc(selectedconcession) # update last used conc in cfg
    return {'num':len(photodf),'path':path, 'sl_file':sl_file_label}

# READ ENTIRE SCREENLOG (1st step of SORT)
@eel.expose
def read_screenlog():
    global screenlog
    print('\nReading Screenlog . . .')
    try:
        screenlog = pd.read_excel(
            sl_file, 
            engine = 'pyxlsb', 
            usecols = [
                'Seq',
                'Sample_ID', 
                'Start_Date',
                'Start_Time',
                'Smpl_photos'
            ])
    except Exception as e:
        print(e, f'\n\nError reading Screenlog ({sl_file})\nExiting . . .')
        sys.exit(0)  
    # DROP ROWS WITH NULL TIME DATA
    screenlog.dropna(subset=['Start_Time'], inplace=True)
    screenlog.reset_index(inplace=True, drop=True)
    # CONVERT DATE FROM FLOAT TO TIMESTAMP FORMAT
    screenlog['Start_Date'] = pd.TimedeltaIndex(
        screenlog['Start_Date'], 
        unit = 'd') + datetime(1899, 12, 30)
    print('\t\t[OK]')
    # FIND SEQUENCE NUMBER OF FIRST SAMPLE ON EARLIEST DATE
    seq = None
    for i, c in screenlog.iterrows():
        day = datetime.date(c['Start_Date'])
        if day == earliest:
            seq = int(c['Seq']) - 1 # READ FROM ONE SAMPLE EARLIER TO ACCOMMODATE MIDNIGHT-CROSSING SAMPLES
            break
    if seq == None:
        print(f'Earliest dated photo ({earliest}) does not match any sample in Screenlog. Please check')
        sys.exit(0)
    print(f'Seq of first sample on {earliest}: {seq}')

    # USED TO BE A SEPARATE FUNC:
    print('\nFiltering . . .')
    screenlog.drop(screenlog[screenlog.Seq < seq].index, inplace = True) # filter in place
    screenlog.loc[:, 'Start_Time'] = pd.TimedeltaIndex(screenlog.loc[:, 'Start_Time'], unit = 'd') # CONVERT DATE AND TIME FROM FLOAT TO TIMESTAMP FORMAT (date AND time)
    screenlog.loc[:, 'Start_Datetime'] = (screenlog.loc[:, 'Start_Date'] + screenlog.loc[:, 'Start_Time']) # MERGE START DATE AND TIME INTO..........:  "DATETIME"
    del screenlog['Start_Date'], screenlog['Start_Time']    # REMOVE NOW-DEFUNCT COLUMNS
    # here find closest sample to first photo (minimum positive timedelta) EASIER TO FIND FIRST SAMPLE WITH PHOTOS
    screenlog.reset_index(inplace=True, drop=True)
    screenlog.Start_Datetime = screenlog.Start_Datetime.dt.to_pydatetime()    # CONVERT DATETIMES FROM TIMESTAMP TO PY_DATETIME FORMAT
    # SET END TIME AS NEXT SAMPLE'S START TIME
    for i in range(0, len(screenlog) - 1):
        screenlog.loc[i, 'End_Datetime'] = screenlog.loc[i + 1, 'Start_Datetime'] # (FINAL SAMPLE END TIME WILL BE NULL)
    screenlog.End_Datetime = screenlog.End_Datetime.dt.to_pydatetime()
    # ITERATE THROUGH SAMPLES, TO COMPILE LISTS OF JPGs PER SAMPLE, ADD LIST IN DATAFRAME
    photoslist, countlist = [], []
    all_zero = True
    fnzi = None
    for i, sample in screenlog.iterrows():
        samplephotos = []
        for ph in photodf.itertuples(index=False):
            if pd.isnull(sample['End_Datetime']) & (ph.Date_Taken > sample['Start_Datetime']):
                samplephotos.append(ph.Photo)
            elif (ph.Date_Taken > sample['Start_Datetime']) & (ph.Date_Taken < sample['End_Datetime']):
                samplephotos.append(ph.Photo)
            else:
                pass
        # find index of last leading empty:
        if len(samplephotos) == 0 and all_zero == True: # check if is another consecutive leading empty sample
            fnzi = i
        else:
            all_zero = False
        countlist.append(len(samplephotos))
        photoslist.append(samplephotos) #samplephotos is a list of JPGs for one sample
    #photoslist is the list of samplelists

    screenlog['Count'] = countlist
    screenlog['Photoslist'] = photoslist
    # drop leading empties here:
    if fnzi != None:
        screenlog = screenlog[screenlog.index > fnzi]
    preview = screenlog.to_json(orient='records', date_format='iso')#index=False)
    return preview # send json to js



# FUNCTION TO GENERATE NAMEBOARDS
def genboard(sample_id, datetaken): # sample_id = str, datetaken = pydatetime
    img = Image.new('RGB', (x, y), color = (255, 255, 255))
    size = fontsize(len(sample_id))
    fnt = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', size)
    lfnt = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', 140)
    d = ImageDraw.Draw(img)
    w = fnt.getsize(sample_id)[0]
    cw = lfnt.getsize(concession)[0]
    d.text(((x-cw)/2, (y/6.6)), concession, font = lfnt, fill = (0, 0, 0))
    d.text(((x-w)/2, (y/2)), sample_id, font = fnt, fill = (0, 0, 0))
    boardfname = (concession + '_' + sample_id + '_01.JPG')
    bdate = datetaken
    imgdate = bdate.strftime(DATETIME_STR_FORMAT)
    exif_dict = {'0th':{306:imgdate},'Exif':{36867:imgdate, 36868:imgdate}}
    exif_bytes = piexif.dump(exif_dict)
    img.save(path + boardfname,exif=exif_bytes)
    mtime = time.mktime(bdate.timetuple())
    os.utime(path + boardfname,(mtime,mtime))
    return boardfname

# --- ITERATE THROUGH LIST OF LISTS, RENAMING FILES:
@eel.expose
def rename():
    print('Renaming Photos . . .')
    log = '<p>'
    renamelist = []
    genlist = []
    # 20210202 rewrite to iterate over sl df instead of photoslist
    for item in screenlog.itertuples(index=False):
        sampleid = item.Sample_ID
        dateinfo = item.Start_Datetime
        for index, data in enumerate(item.Photoslist): # LIST OF JPGs FOR ONE SAMPLE ('item')
            if index == 0:
                nb = genboard(sampleid, dateinfo)
                genlist.append(nb)
                log += (f'<br>-Nameboard generated as {nb}<br>')#, '\t\t\twith time of', newdate)
            filename = (f'{concession}_{sampleid}_{(index + 2):02d}.JPG')
            os.rename(path + data, path + filename)
            log += (f'{data} renamed to {filename}<br>')
            renamelist.append(filename)
    log += '</p>'
    print('[FINISHED]')
    return {'log':log, 'renamed':len(renamelist), 'genned':len(genlist)}


# --- start program on home page in fullscreen mode
print('\nINITIALIZING WEB APP . . .')
eel.start("home.html",mode="firefox")#, cmdline_args=['--start-fullscreen'])

