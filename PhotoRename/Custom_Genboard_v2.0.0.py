import os
from PIL import Image, ImageDraw, ImageFont
from exif import (Image as exif, DATETIME_STR_FORMAT)
import piexif
import datetime as dt
import time

savepath = 'D:/SamplePhotoRename/'
# CARD DIMENSIONS
x = 800 #1152
y = 600 #864

def continue_prompt(msg='Continue? [Y/N] '):
    proceed = ''
    while proceed != 'Y':
        proceed = input(msg)
        if proceed.upper() == 'N':
            return False
        elif proceed.upper() == 'Y':
            return True
        else:
            print('Epic fail! Please respond with Y or N, not "' + proceed + '"')
            time.sleep(1.2)

def get_datetime():
    print('YYYY-MM-DD date for genboard')
    valid_date = False
    while valid_date == False:
        date = input('Enter date:\n ')
        try: 
            date_obj = dt.datetime.strptime(date,'%Y-%m-%d')
            valid_date = True
            break
        except ValueError:
            print('Unable to use', date, '\nPlease use the format YYYY-MM-DD')
    print('HH:MM (24hr format) time for genboard')
    valid_time = False
    while valid_time == False:
        time = input('Enter time:\n ')
        try: 
            time_obj = dt.datetime.strptime(time,'%H:%M').time()
            valid_time = True
            break
        except ValueError:
            print('Unable to use', time, '\nPlease use 24HR format HH:MM')
    timetaken = dt.datetime.combine(date_obj, time_obj)
    return timetaken

def fontsize(length):
    fs = -10 * length +220
    return fs

def genboard(concession, sample_id, bdate):
    img = Image.new('RGB', (x, y), color = (255, 255, 255))
    d = ImageDraw.Draw(img)
    size = fontsize(len(sample_id))
    fnt = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', size)
    lfnt = ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf', 140)
    w, h = fnt.getsize(sample_id)
    cw, ch = lfnt.getsize(concession)
    d.text(((x-cw)/2, (y/6.6)), concession, font = lfnt, fill = (0, 0, 0))
    d.text(((x-w)/2, (y/2)), sample_id, font = fnt, fill = (0, 0, 0))
    boardfname = concession + '_' + sample_id + '_01.jpg'
    imgdate = bdate.strftime(DATETIME_STR_FORMAT)
    exif_dict = {'0th':{306:imgdate},'Exif':{36867:imgdate, 36868:imgdate}}
    exif_bytes = piexif.dump(exif_dict)
    img.save(savepath + boardfname,exif=exif_bytes)
    mtime = time.mktime(bdate.timetuple())
    os.utime(savepath + boardfname,(mtime,mtime))
    print(f'Board saved as {boardfname} in {savepath}')

def main():
    print('Generate a nameboard for sample photos:')
    concession = input('Enter Concession name: ')
    sample_id = input('Enter Sample ID: ')
    print('Date and time info for "Date Taken", "Date Modified"')
    bdate = get_datetime()
    genboard(concession, sample_id, bdate)
    print('[FINISHED]')

main()




