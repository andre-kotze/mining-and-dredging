import os
import time
import sys

import pandas as pd
import datetime as dt

yesterday = (dt.datetime.now().date() - dt.timedelta(days = 1)).strftime("%Y-%m-%d")
outputdir = 'C:/Users/Geo/Desktop/TTanalysis/'
holedatadir = outputdir + 'HoleData/'
settings = 'Z:/IMDH App Settings/TTanalysis/'
if not os.path.exists(settings):
    settings = 'C:/Users/Geo/Documents/Python Scripts/TTanalysis/backup settings/'
lithologies = pd.read_csv(settings + 'Lithologies.csv')
disposable_screenlog = False # TO NOT DELETE ORIGINAL SCREENLOG

def continue_prompt(msg='Continue? [Y/N] ', h=False):
    proceed = ''
    while proceed != 'Y':
        proceed = input(msg)
        if proceed.upper() == 'N':
            return False
        elif proceed.upper() == 'Y':
            return True
        elif h == True and proceed.upper() == 'HELP':
            return 'HELP'
        else:
            print("Epic fail! Please respond with Y or N, not " + proceed)
            time.sleep(1.2)

def find_screenlog():
    cpaths = pd.read_csv(settings + 'concessionpaths.csv')
    valid = concessions['Concession'].values.tolist()
    print('Saved concession paths:', ', '.join(valid))
    concession = (input('Enter Concession name:\t')).upper()
    while concession not in valid:
        concession = input('Try again..\nEnter Concession name:\t')
    sl = cpaths.loc[cpaths['Concession'] == concession, 'PathToScreenlog'].iloc[0]
    if os.path.exists(sl):
        return sl
    else:
        print('File not found at', sl)
        exit = input('Press Enter to exit')
        sys.exit(0)

def read_trends():
    trendlist = []
    for file in sorted(os.listdir(holedatadir)):
        if '.xlsx' in file:
            trendlist.append(file[:-5])
        elif 'Screenlog' in file and '.XLSB' in file:
            screenlogfile = holedatadir + file
            disposable_screenlog = True
        elif '~$' in file:
            print(file, 'is currently open in another program. Please close all files in the folder and try again')
            exit = input('Press Enter to exit')
            sys.exit(0)
    if screenlogfile == None:
        print('No Screenlog file found in folder:\n', holedatadir)
        if continue_prompt('Find screenlog from concession name? [Y/N] ') == True:
            screenlogfile = find_screenlog()
        else:
            sys.exit(0)
    return trendlist, screenlogfile

def check_files(slist, tlist):
    slist.sort()
    tlist.sort()
    matches = len(tlist)
    if slist == tlist:
        print('Samples in screenlog and hole data files match.')
        return matches
    else:
        print('Mismatch in screenlog samples and tooltrend files')
        print(len(slist), 'samples in Screenlog,', len(tlist), 'sample hole data files\nCHECK:')
        for s in slist:
            if s not in tlist:
                print(s, 'occurs in screenlog, but not in hole data')
        for t in tlist:
            if t not in slist:
                print(t, 'occurs in hole data, but not in screenlog')
                matches -= 1
        return matches

def get_date():
    valid_date = False
    while valid_date == False:
        date = input('Enter filter date YYYY-MM-DD:\n')
        try: 
            dt.datetime.strptime(date,'%Y-%m-%d')
            valid_date = True
            break
        except ValueError:
            print('Unable to use', date, '\nPlease use the format YYYY-MM-DD')
    return date

def read_screenlog(screenlogfile, date): # date in YYYY-MM-DD
    read_init = dt.datetime.now()
    try:
        screenlog = pd.read_excel(screenlogfile, engine='pyxlsb', usecols=[
            'Sample_ID',
            'Start_Date',
            'Max_Pen',
            'Unit_1', 'm1',
            'Unit_2', 'm2',
            'Unit_3', 'm3',
            'Unit_4', 'm4',
            'Unit_5', 'm5',
            'Unit_6', 'm6',
            'Unit_7', 'm7',
            'Unit_8', 'm8',
            'Unit_9', 'm9',
            'Unit_10', 'm10'
            ])
    except Exception as e:
        print(e, '\nError reading screenlog')
    # convert dates in screenlog from float format to datetime format
    screenlog['Start_Date'] = pd.TimedeltaIndex(screenlog['Start_Date'], unit = 'd') + dt.datetime(1899, 12, 30)
    screenlog = screenlog[screenlog['Start_Date'] == date]
    del screenlog['Start_Date']
    dur = dt.datetime.now() - read_init
    dur = dur.seconds + (dur.microseconds / 1000000)
    print('Screenlog processed in', round(dur, 2), 'seconds')
    return screenlog

def get_units(sample): # given a screenlog entry(row), returns a dict with SampleID as key, with units, start depths and end depths
    unitslist = []
    s_depth = 0
    for u in range(1,11):  # construct a list: [unit,startdepth,enddepth] reading up to 10 units from the screenlog
        m = 'm' + str(u)
        unit = 'Unit_' + str(u)
        if pd.notnull(sample[unit]):  
            ulist = []          
            ulist.append(sample[unit])                  # unit name
            ulist.append(s_depth)                       # start depth
            e_depth = sample[m] + s_depth
            ulist.append(e_depth)                       # end depth
            unitslist.append(ulist)
            s_depth = e_depth 
    #unitslist = list of lists
    return unitslist
        
def process_trend(id, units): 
    # id is the sample_id, units is the dict of units, start and end depths
    trendsfile = id + '.xlsx'
    try:
        ttdata = pd.read_excel(holedatadir + trendsfile, skiprows=1, usecols=[
            'Relative Time\nhh:mm:ss',
            'Rack Drive Lowering\nKN',
            'Actual Drill Torque\nKN*m',
            'Feed Result\nmm/min',
            'Drill Depth\nmm',
            ])
    except Exception as e:
        print(e, '\nError reading holedata file:', trendsfile)
        sys.exit(0)
    drillspeed = []
    for i, r in ttdata.iterrows():
        if i == 0:
            drillspeed.append(0)
        else:
            rate = (ttdata.loc[i,'Drill Depth\nmm'] - ttdata.loc[i-1,'Drill Depth\nmm']) / 2000
            drillspeed.append(rate)
    ttdata['DrillSpeed_m/s'] = drillspeed
    ttdata = ttdata[ttdata['Drill Depth\nmm'] >= 0]
    maxdepthindex = (ttdata[ttdata['Drill Depth\nmm'] == ttdata['Drill Depth\nmm'].max()].index.values[-1])
    ttdata = ttdata[ttdata.index <= maxdepthindex]
    ttdata['Drill Depth\nmm'] = ttdata['Drill Depth\nmm'] / 1000
    results = []
    for lst in units:
        l, start, end = lst[0], lst[1], lst[2]
        lith = lithologies.loc[lithologies['CODE'] == l, 'LITHOLOGY'].iloc[0]
        curve = ttdata[ttdata['Drill Depth\nmm'].between(start, end)]
        # drop zero force values, drilling has stopped
        curve = curve[curve['Actual Drill Torque\nKN*m'] != 0]
        avg_rate = curve['DrillSpeed_m/s'].mean()
        avg_torque = curve['Actual Drill Torque\nKN*m'].mean()
        avg_force = curve['Rack Drive Lowering\nKN'].mean()
        results.append([id, lith, start, end, round(end-start,2), avg_rate, avg_torque, avg_force]) # list of lists
    return results

def export(output_data, date):
    output = pd.DataFrame(data = output_data, columns = [
        'Sample_ID', 
        'Lithology', 
        'Depth_Start_m', 
        'Depth_End_m', 
        'Thickness_Unit_m', 
        'Penetration_Rate_m/s', 
        'Torque_kNm', 
        'Force_kN'
        ])
    filename = outputdir + 'ToolTrend_Analysis_' + date + '.csv'
    output.to_csv(filename, index = False)
    print('Results exported as', filename)

def main():
    with open(settings + 'heading.txt','r') as title:
        print(title.read(), '\n\n')
    time.sleep(2)
    trendlist, screenlogfile = read_trends()
    print(len(trendlist), 'Hole Data files found')
    if continue_prompt('Continue? [Y/N] (type "help" for instructions) ', h=True) == 'HELP':
        with open(settings + 'instructions.txt','r') as instr:
            print(instr.read(), '\n\n')
        continue_prompt()
    if continue_prompt("use yesterday's date " + yesterday + " ? [Y/N] ") == True:
        date = yesterday
    else:
        date = get_date()
    print('Reading Screenlog . . .')
    screenlog = read_screenlog(screenlogfile, date)
    print(len(screenlog), 'samples found for', date)
    print('Checking Hole Data files . . .')
    samplelist = screenlog['Sample_ID'].values.tolist()
    matched = check_files(samplelist, trendlist)
    if matched == len(samplelist):
        if continue_prompt() == False:
            sys.exit(0)
    else:
        if continue_prompt('Run analysis anyway on ' + str(matched)  + ' matching samples? [Y/N]') == False:
            sys.exit(0)
    print('Running Tooltrend analysis . . .')
    init_time = dt.datetime.now()
    outdata = []
    for i, s in screenlog.iterrows():
        if s['Sample_ID'] not in trendlist:
            print(s['Sample_ID'], 'skipped (no matching Hole Data)')
        else:
            print('Processing', s['Sample_ID'], end='', flush=True)
            units = get_units(s)
            if len(units) != 0:
                for l in process_trend(s['Sample_ID'], units):
                    outdata.append(l)
                print('\t\t\t\t[OK]')
            else:
                print('\t\t\t\t[NO UNITS]')
    try:
        export(outdata, date)
    except Exception as e:
        print(e, '\nOutput data not exported due to above error')
    dur = dt.datetime.now() - init_time
    dur = dur.seconds + (dur.microseconds / 1000000)
    print('Analysis completed in', round(dur, 2), 'seconds')
    if disposable_screenlog == True:
        if continue_prompt('Delete Screenlog from folder? [Y/N] ') == True:
            if outputdir in screenlogfile: # DOUBLE CHECK TO NOT DELETE ORIGINAL SCREENLOG
                os.remove(screenlogfile)
                print(screenlogfile, 'deleted')
    if continue_prompt('Delete Hole Data files from folder? [Y/N] ') == True:
        for f in trendlist:
            os.remove(holedatadir + f + '.xlsx')
            print(f + '.xlsx deleted')
    exit = input('FINISHED! Press enter to exit')

main()
# ToolTrend sheet in GeoT DB:
# Sample_ID     Lithology   Depth_Start_m   Depth_End_m     Thickness_Unit_m    Penetration_Rate_m/s    Torque_kNm      Force_kN
