import os
import pandas as pd
import datetime as dt
import time

outdata = []
holedatadir = 'C:/Users/Geo/Desktop/TTanalysis/HoleData/'
lithologies = pd.read_csv('Z:/IMDH App Settings/TTanalysis/Lithologies.csv')

def read_trends():
    trendlist = []
    for file in sorted(os.listdir(holedatadir)):
        if '.xlsx' in file:
            trendlist.append(file[:-5])
        elif 'Screenlog' in file and '.XLSB' in file:
            screenlogfile = file
        elif '~$' in file:
            print(file, 'is currently open in another program. Please close all files in the folder and try again')
            exit = input('Press Enter to exit')
            exit()
    return trendlist, screenlogfile

def check_files(slist, tlist):
    if slist == tlist:
        return True
    else:
        print('Mismatch in screenlog samples and tooltrend files')
        print(len(slist), 'samples in Screenlog,', len(tlist), 'sample hole data files for', date, '\nCHECK:')
        for s in slist:
            if s not in tlist:
                print(s, 'occurs in screenlog, but not in hole data')
        for t in tlist:
            if t not in slist:
                print(t, 'occurs in hole data, but not in screenlog')
        return False

def read_screenlog(screenlogfile, date): # date in YYYY-MM-DD
    screenlog = pd.read_excel(holedatadir + screenlogfile, engine='pyxlsb', usecols=[
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
        'Unit_10', 'm10',
        ])
    # convert dates in screenlog from float format to datetime format
    screenlog['Start_Date'] = pd.TimedeltaIndex(screenlog['Start_Date'], unit = 'd') + dt.datetime(1899, 12, 30)
    screenlog = screenlog[screenlog['Start_Date'] == date]
    print(len(screenlog))
    del screenlog['Start_Date']
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
    ttdata = pd.read_excel(holedatadir + trendsfile, skiprows=1, usecols=[
        'Relative Time\nhh:mm:ss',
        'Rack Drive Lowering\nKN',
        'Actual Drill Torque\nKN*m',
        'Feed Result\nmm/min',
        'Drill Depth\nmm',
        ])
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
    filename = "ToolTrend_Analysis_" + date + ".csv"
    output.to_csv(filename, index = False)
    print("Results exported as", filename)

def main():
    print('IMDH GEOTECHNICAL SAMPLE TOOLTREND ANALYSIS')
    trendlist, screenlogfile = read_trends()
    date = input('Enter filter date YYYY-MM-DD:\n')
    print('Reading Screenlog . . .')
    screenlog = read_screenlog(screenlogfile, date)
    print('Checking Hole Data files . . .')
    samplelist = screenlog['Sample_ID'].values.tolist().sort()
    trendlist = trendlist.sort()
    if check_files(samplelist, trendlist) == False:
        exit = input('Check Hole Data files and Screenlog entries and try again\nPress Enter to exit')
        exit()
    print('Running Tooltrend analysis . . .')

    for i, s in screenlog.iterrows():
        print('Processing', s['Sample_ID'])
        units = get_units(s)
        for l in process_trend(s['Sample_ID'], units):
            outdata.append(l)
    try:
        export(outdata, date)
        exit = input("Enter to exit")
    except Exception as err:
        print(err)

main()
# ToolTrend sheet in GeoT DB:
# Sample_ID     Lithology   Depth_Start_m   Depth_End_m     Thickness_Unit_m    Penetration_Rate_m/s    Torque_kNm      Force_kN
