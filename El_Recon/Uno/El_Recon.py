'''
read diamond data from 4 sources and attempt to reconcile:
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

Files to be copied to script dir

Filter/trim by DATE

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

epoch = dt.datetime(1899, 12, 30)

# even need this?
lists = pd.read_csv('lists.csv')
shapes = lists['Shape'].values.tolist()
colours = lists['Colour'].values.tolist()
clarities = lists['Clarity'].values.tolist()

props = {0:'Carat', 1:'Shape', 2:'Colour', 3:'Clarity'}

log = [] # list of string feedback messages

def prt(printarg): # PRINT RIGHT THERE
    print(printarg, end='', flush=True)

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
            print('Epic fail! Please respond with Y or N, not "' + proceed + '"')
            time.sleep(1.2)

def get_files():
    print('\nChecking files . . .')
    for file in os.listdir():
        if '~$' in file:
            print(file, 'is open in another program.\nPlease close all files in', os.getcwd())
            if continue_prompt('Retry? [Y/N] ') == True:
                drs, sl, sr = get_files()
        elif 'DiamondRecordSheet' in file:
            drs = file
            print('Found', file)
        elif 'Screenlog' in file:
            sl = file
            print('Found', file)
        elif 'SamplingResults' in file:
            sr = file
            print('Found', file)
        else:
            log.append(file + ' ignored')
    return drs, sl, sr

def get_date(label):
    valid_date = False
    while valid_date == False:
        date = input(label + ':\n ')
        try: 
            date_obj = dt.datetime.strptime(date,'%Y-%m-%d')
            valid_date = True
            break
        except ValueError:
            print('Unable to use', date, '\nPlease use the format YYYY-MM-DD')
    return [date, date_obj]

def bounding_dates():
    print('\nEnter bounding values for DATE DRILLED in YYYY-MM-DD format')
    allgood = False
    while allgood == False:
        start = get_date('Start Date:')
        end = get_date('End Date:')
        if start[1] <= end[1]:
            days = (end[1] - start[1]).days + 1
            print('Diamond data for', days, 'days will be checked')
            allgood = True
        else:
            print('Start Date cannot be after End Date!\nPlease try again')
    return start[0], end[0]

def read_drs(name, dates):
    print('\nReading Diamond Record Sheet . . .')
    read_init = dt.datetime.now()
    #try:
    sheets = pd.ExcelFile(name).sheet_names
    for sh in sheets:
        if 'DiamondSortingLog' in sh:
            sheetname = sh
            break
    drs = pd.read_excel(name, sheet_name=sheetname, skiprows=1, usecols='A,C,D,H,AM:AO,AV:AOM')
    drs.dropna(how='all', inplace=True)
    drs.dropna(axis=1, how='all', inplace=True)
    drs.columns = np.arange(0,drs.shape[1])
    drs.rename({
        0:'Seq No.',
        1:'Sample ID',
        2:'Date Drilled',
        3:'Date Sorted (start)', # for Audits, use date sorted as date
        4:'Total Stones',
        5:'Total Carats',
        6:'Group Weight'}, 
        axis=1, inplace=True)
    drs['Date Drilled'].fillna(drs['Date Sorted (start)'], inplace=True)
    drs['Group Weight'].fillna(drs['Total Carats'], inplace=True)
    drs['Date Drilled'] = pd.to_datetime(drs['Date Drilled'], errors='coerce')
    drs = drs[drs['Date Drilled'].between(dates[0], dates[1])] # ev dropped here
    drs.dropna(axis=1, how='all', inplace=True) # drop empty cols
    col_dict = {}
    props = {1:'Carat', 2:'Shape', 3:'Colour', 4:'Clarity'}
    max_stones = (drs.shape[1] - 7) / 4 # check that == int
    if max_stones != int(max_stones):
        print('Error in Diamond Record Sheet! Max number of stones evaluates to', max_stones)
    # stones start at column 7
    for d in range(1, int(max_stones) + 1): # 1 to 19
        for p in range(1,5):
            col_dict[d*4 + p + 2] = 'D' + str(d) + props[p]
    drs.rename(col_dict, axis=1, inplace=True)
    drs.reset_index(inplace = True, drop=True)
    print(drs.shape)
    dur = dt.datetime.now() - read_init
    dur = dur.seconds + (dur.microseconds / 1000000)
    print('Diamond Record Sheet processed in', round(dur, 2), 'seconds (', len(drs), ' records)')
    return drs, max_stones
#except Exception as e:
    #print(e, '\nError reading Diamond Record Sheet')
    #return None


def read_sl(name, dates):
    print('\nReading Screenlog . . .')
    read_init = dt.datetime.now()
    #try:
    sl = pd.read_excel(name, engine='pyxlsb', usecols=[
        'Seq',
        'Sample_ID',
        'Start_Date',
        'Stones',
        'Group_wt',
        'Est_Carats',
        'Est_Brkdwn',
        'Shape',
        'Colour',
        'Clarity'
        ])
    #sl.dropna(how='all', subset=['Stones'], inplace=True)
    sl['Stones'].fillna(0, inplace=True)
    sl['Start_Date'] = pd.TimedeltaIndex(sl['Start_Date'], unit = 'd') + epoch
    sl = sl[sl['Start_Date'].between(dates[0], dates[1])]

    dur = dt.datetime.now() - read_init
    dur = dur.seconds + (dur.microseconds / 1000000)
    print('Screenlog processed in', round(dur, 2), 'seconds (', len(sl), ' records)')
    return sl
    #except Exception as e:
    #    print(e, '\nError reading Screenlog')
    #    return None

def read_sr(name, dates): # works for guinea file
    print('\nReading Sampling Results . . .')
    read_init = dt.datetime.now()
    #try:
    sr = pd.read_excel(name, skiprows=4, usecols=[
        'Date',
        'Sample N.\n&\nBarcode N.',
        'Individual  Stone Count Offshore',
        'IndividualCarats Offshore ',
        'Individual Stone Shape',
        'Individual Stone Colour',
        'Individual Stone Clarity'
        ])
    sr.dropna(how='all', inplace=True)
    sr.reset_index(drop=True, inplace=True)
    sr.drop(0, inplace=True)
    sr['Date'] = pd.to_datetime(sr['Date'], errors='coerce')
    sr['Date'].fillna(method='ffill', inplace=True)
    sr = sr[sr['Date'].between(dates[0], dates[1])]
    sr['Sample N.\n&\nBarcode N.'].fillna(method='ffill', inplace=True)
    sr.drop(sr[sr['Sample N.\n&\nBarcode N.'].str.contains('Samples')].index, inplace=True)
    sr.reset_index(drop=True, inplace=True)
    # now have a clean dataset per stone
    dur = dt.datetime.now() - read_init
    dur = dur.seconds + (dur.microseconds / 1000000)
    print('Sampling Results processed in', round(dur, 2), 'seconds (', len(sr), ' records)')
    return sr
    #except Exception as e:
    #    print(e, '\nError reading Sampling Results sheet')
    #    return None

def compare_lists(list1, list2, label): # use df1.eq(df2)
    mismatches = []
    for i, (d, s) in enumerate(zip(list1, list2)):
        if str(d) != str(s):
            try:
                if round(float(d), 2) != round(float(s), 2):
                    issue = ('\t' + label + ':\t"' + str(d) + '"\t\t\t"' + str(format(s, ':4.2f')) + '"')
                    mismatches.append([i, issue])
            except Exception as e:
                log.append(e)
                issue = ('\t' + label + ':\t"' + str(d) + '"\t\t\t"' + str(s) + '"')
                mismatches.append([i, issue])
    return mismatches # a list of lists

def compare_stones(list1, list2, sid): # use df1.eq(df2)
    mismatches = []
    for i, (sl, sr) in enumerate(zip(list1, list2)):
        if str(sl) != str(sr):
            try:
                if round(float(sl), 2) != round(float(sr), 2):
                    issue = ('\t' + props.get(i) + ':\t"' + str(sl) + '"\t\t\t"' + str(format(sr, ':4.2f')) + '"')
                    mismatches.append([sid, issue])
            except Exception as e:
                log.append(e)
                issue = ('\t' + props.get(i) + ':\t"' + str(sl) + '"\t\t\t"' + str(sr) + '"')
                mismatches.append([sid, issue])
    return mismatches # a list of lists

def compare_series(series1, series2, label, ids): # ids is a series of SampleIDs
    if series1.eq(series2).any() == False:
        mm = compare_lists(series1.values.tolist(), series2.values.tolist(), label)
        if len(mm) == 0:
            print(label, 'match')
        else:
            for m in mm:
                print('Check', label, 'SampleID:\t', ids.loc[m[0]])
    else:
        print('No mismatches in', label)

def sl_vs_drs(sl, drs, max_stones):
    print('Comparing Screenlog and Diamond Record Sheet data (Audits not included) . . .')
    drs.dropna(how='all', subset=['Seq No.'], inplace=True)
    drs.reset_index(drop=True, inplace=True) # no audits in screenlog
    # lines to attempt to read in DRS:
    if len(drs) != len(sl):
        print('Number of samples in Screenlog (', len(sl), ') and DRS (', len(drs), ') do not match')
    else:
        print(len(sl), 'samples in Screenlog match', len(drs), 'samples in DRS')

    collec_cts, collec_shps, collec_cols, collec_cla = [], [], [], []
    for i, r in drs.iterrows():
        #prt('\nrow: ' + str(i) + '\t')     iterator tracker
        drs_carats, drs_count, drs_shapes, drs_colours, drs_clarities = [], [], [], [], []
        for n in range(7,(max_stones*4)+6,4): # (max_stones*4)+6   +6??? +4??? +7/8??
            #print(n, end='', flush=True)
            if pd.notnull(drs.iloc[i, n]): # insert validation here
                drs_carats.append(format(drs.iloc[i, n], '.2f'))
                drs_shapes.append(format(drs.iloc[i, n+1], '.0f'))
                drs_colours.append(str(drs.iloc[i, n+2]))
                drs_clarities.append(str(drs.iloc[i, n+3]))
            #prt('.')
        collec_cts.append(', '.join(drs_carats))
        collec_shps.append(', '.join(drs_shapes))
        collec_cols.append(', '.join(drs_colours))
        collec_cla.append(', '.join(drs_clarities))

    sl.fillna('', inplace=True)
    drs['Group Weight'].fillna('0.0', inplace=True)
    drs['Total Carats'].fillna('0.0', inplace=True)
    mm_cts = compare_lists(collec_cts, sl['Est_Brkdwn'].values.tolist(),'Carat')
    mm_shps = compare_lists(collec_shps, sl['Shape'].values.tolist(),'Shape')
    mm_cols =  compare_lists(collec_cols, sl['Colour'].values.tolist(),'Colour')
    mm_cla = compare_lists(collec_cla, sl['Clarity'].values.tolist(),'Clarity')

    
    mms = (mm_cts + mm_shps + mm_cols + mm_cla)
    if len(mms) > 0:
        print('Seq:\tSampleID:\tIssue\t\tDRS_val:\t\t\t\tSL_val:')
    else:
        print('No mismatches in Carats, Shapes, Colours, Clarities')
    for mm in mms:
        print(drs.loc[mm[0], 'Seq No.'], '\t', drs.loc[mm[0], 'Sample ID'], mm[1])

    compare_series(drs['Total Stones'], sl['Stones'],'Total Stones',drs['Sample ID'])
    compare_series(drs['Group Weight'], sl['Group_wt'],'Group Weight',drs['Sample ID'])
    compare_series(drs['Total Carats'], sl['Est_Carats'],'Total Carats',drs['Sample ID'])

    
    print('\n\n\n[ DRS vs SL complete ]')

def sl_vs_sr(sl, sr):
    print('Comparing Screenlog and SamplingResults data (Audits not included) . . .')
    # convert sl data to per stone i.e. for each sample create 1 line, and 1 additional line/stone > 1
    # create a table similar to the SR sheet
    #sl_cts, sl_shps, sl_cols, sl_cla = [], [], [], []
    sl_per_stone = []
    sl_count = 0
    samples = sl['Sample_ID'].values.tolist()
    for i, r in sl.iterrows():
        stonecount = int(r['Stones'])
        sl_count += stonecount
        if stonecount == 0:
            sl_per_stone.append([format(0, '.2f'), '', '', ''])
        elif stonecount == 1:
            sl_per_stone.append([r['Est_Brkdwn'], r['Shape'], r['Colour'], r['Clarity']])
        else:
            for s in range(stonecount):
                sl_per_stone.append([
                    (r['Est_Brkdwn'].split(', '))[s],  # ct
                    (r['Shape'].split(', '))[s],       # shape
                    (r['Colour'].split(', '))[s],      # colour
                    (r['Clarity'].split(', '))[s]      # clarity
                ])


    #drop audits, sterilizations from sr
    sr.drop(sr[
        (sr['Sample N.\n&\nBarcode N.'].str.contains('ADT')) | (sr['Sample N.\n&\nBarcode N.'].str.contains('Steril'))
        ].index, inplace=True)
    # reset index sr
    sr.reset_index(drop=True, inplace=True)
    sr_count = sr['Individual  Stone Count Offshore'].sum()
    sr.fillna('', inplace=True)
    print(f'Screenlog shows {sl_count} total stones, SR shows {sr_count}')
    print('SampleID\tProperty\tScreenlog value\t\tSR value')
    for i, r in sr.iterrows():
        sr_stone = [r['IndividualCarats Offshore '], r['Individual Stone Shape'], r['Individual Stone Colour'], r['Individual Stone Clarity']]
        mms = compare_stones(sl_per_stone[i], sr_stone, r['Sample N.\n&\nBarcode N.'])
        for mm in mms:
            print(mm[0], mm[1])

    print('\n\n\n[ SL vs SR complete ]')


def main():
    #dates = ['2020-11-24', '2020-11-25']
    dates = bounding_dates()
    drs_filename, sl_filename, sr_filename = get_files()
    # return dataframes with diamond data per sample
    sl = read_sl(sl_filename, dates)
    if continue_prompt('Skip DRS check? [Y/N] ') == False:
        drs, max_stones = read_drs(drs_filename, dates)
        if continue_prompt('Max stone count for this date range: ' + str(max_stones) + ', correct? [Y/N] ') == False:
            print('Check fields in Diamond Record sheet and retry')
            sys.exit(0)
        print('\nReconciling Screenlog with Diamond Record Sheet . . .')
        sl_vs_drs(sl, drs, int(max_stones))
        if continue_prompt() == False:
            sys.exit(0)

    # return dataframe with diamond data per stone
    sr = read_sr(sr_filename, dates)


    print('Reconciling Screenlog with Sampling Results . . .')
    sl_vs_sr(sl, sr)

    viewlog = input("\nPress Enter to exit, or L to view the log\n")
    if viewlog.upper() == "L":
        print('')
        for i in log:
            print('\t', i)
        print('')
        exit = input("\nPress Enter to exit")
    else:
        sys.exit(0)
    


main()