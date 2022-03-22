# === AKO update === v1.2.4
#. . . . 1 . . . . 2 . . . . 3 . . . . 4 . . . . 5 . . . . 6 . . . . 7 . . . . 8
import os
import datetime as dt
import socket
import copy

import numpy as np
import pandas as pd
from bokeh.plotting import figure, curdoc
from bokeh.layouts import layout
from bokeh.io import export_png, curdoc
from bokeh.models import (
    Range1d, 
    LinearAxis, 
    Label, 
    Span, 
    ColumnDataSource, 
    ImageURL, 
    SingleIntervalTicker
)

moving_average = 5 # for penetration rate. Must be an odd integer. Value of 1 will plot raw penetration rate
number_units = 10 # number of units to attempt to read from the screenlog (default 10)
pi = 3.14159265359
epoch = dt.datetime(1899, 12, 30)
init_time = dt.datetime.now()
host = socket.gethostname()
if host.upper() == 'GEOTECH-PC':
    workingdir = 'C:/Users/Geo/Desktop/TTplot/'
    settings = 'Z:/IMDH App Settings/TTplot/'
    if not os.path.exists(settings):
        settings = 'C:/Users/Geo/Documents/Python Scripts/TTplot/backup settings/'
elif host.upper() == 'GEO-CAPTURE':
    workingdir = 'C:/Users/User/Desktop/TTplotGeo/'
    settings = 'C:/Users/User/Desktop/SHARE/IMDH App Settings/TTplot/'
imdh = settings[0:-7] + 'logo/Small keyboard.txt'
logo_url = "file://" + settings + "IMDSA.png"

gravels = pd.read_csv(settings + 'Gravels.csv')
gravels = gravels['GravelUnits'].values.tolist()
# --- if the folder doesn't exist, create a folder to export plots into
if not os.path.exists('plots'):
    plots = os.mkdir('plots')
logmessages = []

ylocs = [70, 65, 60, 55, 50, 45, 37, 32, 27, 22, 17, 12, 7, 2] # Spacing 

url = logo_url
source = ColumnDataSource(dict(
url = [url],
x1  = np.linspace(50, 50, 1),
y1  = np.linspace(90, 90, 1),
w1  = np.linspace(50, 50, 1),
h1  = np.linspace(50.3*0.35, 50.30*0.35, 1), 
))
logo = ImageURL(url="url", x="x1", y="y1", w='w1', h="h1", anchor="center")

def check_dir():
# --- check all files in directory
    tooltrendlist = []
    for file in os.listdir():
        if "Screenlog" in file:
            screenlogfile = file
        elif '.xlsx' in file:
            tooltrendlist.append(file[:-5])
    return screenlogfile, tooltrendlist

def read_screenlog(screenlogfile):
    sl = pd.read_excel(
        screenlogfile, 
        engine='pyxlsb', usecols=[
            'Sample_ID',
            'Concession',
            'Feature',
            'TS_Act',
            'WD_act',
            'Start_Date',
            'Start_Time',
            'End_Time',
            'E_actual',
            'N_actual',
            'Tool',
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
    # --- trim screenlog dataframe to only contain samples for which tool trend files are present
    trim = sl.Sample_ID.isin(tooltrendlist)
    sl = sl[trim]
    sl['Start_Date'] = pd.TimedeltaIndex(sl['Start_Date'], unit = 'd') + epoch
    sl['Start_Time'] = pd.TimedeltaIndex(sl['Start_Time'], unit = 'd') + epoch
    sl['End_Time'] = pd.TimedeltaIndex(sl['End_Time'], unit = 'd') + epoch
    return sl

def read_holedata(filename): # --- extract tool trend data from tool trend file
    tooltrend = pd.read_excel(
        filename + '.xlsx', 
        sheet_name='HoleData', 
        skiprows=1, 
        usecols=[
            'Actual Drill Torque\nKN*m', 
            'Tower Feed Setpoint\n%',
            'Drill Depth\nmm',
            'Relative Time\nhh:mm:ss',
            'Pitch\n°',
            'Roll\n°',
            'Rack Drive Lowering\nKN',
            'Feed Result\nmm/min',
            'Swivel RPM\nrpm'
    ])
    depth = tooltrend['Drill Depth\nmm']
    maxdepth = depth.max()
    maxdepthindex = (
        tooltrend[tooltrend['Drill Depth\nmm'] == maxdepth].index.values[-1])
    tooltrend = tooltrend[tooltrend['Drill Depth\nmm'] >= 0]
    tooltrend = tooltrend[tooltrend.index <= maxdepthindex]
    if moving_average > 1:
        tooltrend.reset_index(inplace=True, drop=True)
        buff = int(moving_average / 2)
        speed = []
        for i in range(buff):
            speed.append(tooltrend.at[i, 'Feed Result\nmm/min'])
        for i in range(buff,len(tooltrend) - buff):
            rmean = tooltrend.loc[i-buff:i+buff,'Feed Result\nmm/min'].sum() / moving_average
            speed.append(rmean)
        for i in range(buff):
            speed.append(tooltrend.at[len(tooltrend)-1, 'Feed Result\nmm/min'])
        tooltrend['Penetration Rate\nmm/min'] = speed
    else:
        tooltrend['Penetration Rate\nmm/min'] = tooltrend['Feed Result\nmm/min']

    tooltrend = tooltrend[tooltrend['Feed Result\nmm/min'] > 0]
    tooltrend = tooltrend[tooltrend['Tower Feed Setpoint\n%'] <= 100]
    (drilldepth, torque, towerfeed, time, pitch, roll, force, mse, penrate, swivel, penratemm) = \
        ([], [], [], [], [], [], [], [], [], [], [])
    (h1, m1, s1) = str(tooltrend['Relative Time\nhh:mm:ss'].iloc[0]).split(':')
    result1 = (int(h1) * 60) + (int(m1)) + (int(s1) / 60)
    for i, tt in tooltrend.iterrows():
        drilldepth.append(tt['Drill Depth\nmm'])
        torque.append(tt['Actual Drill Torque\nKN*m'])
        towerfeed.append(tt['Tower Feed Setpoint\n%'])
        pitch.append(tt['Pitch\n°'])
        roll.append(tt['Roll\n°'])
        force.append(tt['Rack Drive Lowering\nKN'])
        swivel.append(tt['Swivel RPM\nrpm'])
        #penrate.append(tt['Feed Result\nmm/min'] / 1000) # speed value used in MSE calculation
        mse.append(((tt['Rack Drive Lowering\nKN'] / 5) + (0.4*pi)*(tt['Swivel RPM\nrpm']*tt['Actual Drill Torque\nKN*m']/(tt['Feed Result\nmm/min'] / 1000)))/1000)
        penratemm.append(tt['Penetration Rate\nmm/min']) # speed value plotted 
        (h, m, s) = str(tt['Relative Time\nhh:mm:ss']).split(':')
        result = (int(h) * 60) + (int(m)) + (int(s) / 60)
        time.append(result - result1)
    return drilldepth, torque, towerfeed, time, pitch, roll, force, mse, penrate, swivel, penratemm #11 items

def template_layout():
    # p1 force and torque
    p1 = figure(
        x_axis_label = 'Torque (kNm)',
        y_axis_label = 'Drill Depth (mm)',
        plot_width = 500,
    )
    p1.outline_line_color = 'black'
    p1.outline_line_width = 1
    p1.add_layout(LinearAxis(x_range_name = 'force', axis_label = 'Force (kN)'), 'above')
    p1.toolbar_location = None
    #p1.legend.location = 'top_left'  #default top_right is less obstructive

    # p2 units
    p2 = figure(
        x_axis_label = 'Total Drill Time (min)',
        y_axis_label = 'Total Depth (mm)',
        plot_width = 300
    )
    p2.toolbar_location = None
    p2.extra_x_ranges = {'tower' : Range1d(start = 0, end = 104)}
    p2.add_layout(LinearAxis(x_range_name = 'tower', axis_label = 'Tower Feed Setpoint (%)'), 'above')

    # p3 MSE
    p3 = figure(
        x_axis_label = 'Mechanical Specific Energy (MPa)',
        y_axis_label = 'Drill Depth (mm)',
        plot_width = 300,
    )
    p3.outline_line_color = 'black'
    p3.outline_line_width = 1
    p3.toolbar_location = None

    # p4 info table
    p4 = figure(
        background_fill_color = 'white',
        plot_width = 350,
        title = ('SAMPLE DRILL REPORT'),
    )
    p4.min_border_top = 75
    p4.min_border_left = 55
    p4.title.text_font_style = 'bold'
    p4.title.text_font_size = '25px'
    p4.title.align = 'center'
    p4.title.text_font_style = 'bold'
    p4.y_range = Range1d(0, 100)
    p4.x_range = Range1d(0, 100)
    p4.toolbar_location = None
    p4.outline_line_color = 'black'
    p4.outline_line_width = 1
    p4.xgrid.grid_line_color = None
    p4.ygrid.grid_line_color = None
    p4.xaxis.visible = False
    p4.yaxis.visible = False
    p4.add_layout(Label(
        x = 50, 
        y = 75, 
        text = 'THE EXPLORER', 
        text_color = 'black', 
        text_font_size = '20px', 
        text_align = 'center'))
    p4.add_glyph(source, logo)

    # p5 penrate
    p5 = figure(
        x_axis_type = None,
        x_axis_label = 'Penetration Rate (mm/min)',
        y_axis_label = 'Total Depth (mm)',
        plot_width = 150,
    )
    ticker = SingleIntervalTicker(interval = 100, num_minor_ticks = 5)
    xaxis = LinearAxis(ticker = ticker)
    p5.add_layout(xaxis, 'below')
    p5.xaxis.axis_label = 'Penetration (mm/min)'
    p5.outline_line_color = 'black'
    p5.outline_line_width = 1
    p5.toolbar_location = None

    template = [p1, p2, p3, p4, p5]
    return template

def plot_force_torque(p1, drilldepth, torque, force):
    # --- generate a plot of tool trends versus drill depth
    p1.y_range = Range1d(max(drilldepth), 0)
    p1.x_range = Range1d(0, max(torque) + 25)
    p1.extra_x_ranges = {'force' : Range1d(start = min(force), end = max(force)+ 25)}
    p1.line(force, drilldepth, line_width = 2, color = 'blue', x_range_name = 'force', legend_label = "Force")
    p1.line(torque, drilldepth, line_width = 2, color = 'red', legend_label = 'Torque')
    return p1

def plot_units(p2, depth, time, toolsink, towerfeed):
    # --- generate a plot of tool sink, units and total depth
    p2.y_range = Range1d(max(depth), 0)
    p2.x_range = Range1d(0, max(time))
    p2.multi_polygons(
        fill_color = (242,242,242),
        line_color = 'black',
        line_width = 1,
        line_alpha = 1,
        xs=[[[[0, 0, max(time), max(time)]]]],
        ys=[[[[(toolsink), 0, 0, (toolsink)]]]],
        hatch_pattern = '\\',
        hatch_color = 'lightgrey',
        hatch_weight = 0.5,
        hatch_scale = 10,
        legend_label = 'Tool Sink'
    )
    p2.multi_polygons(
        fill_alpha = 0,
        line_color = 'black',
        line_width = 1,
        line_alpha = 1,
        xs=[[[[0, 0, max(time), max(time)]]]],
        ys=[[[[max(depth), min(depth), min(depth), max(depth)]]]],
    )
    p2.line(time, depth, color = 'black', line_width = 2, legend_label = "Drill Time")
    p2.line(towerfeed, depth, 
        line_width = 2, 
        color = 'black', 
        x_range_name = 'tower', 
        line_dash = 'dashed', 
        legend_label = "Tower Feed")
    return p2

def plot_mse(p3, mse, drilldepth):
    # --- generate a plot that displays mechanical specific energy
    p3.line(mse, drilldepth, line_width = 2, color = 'orange')
    p3.y_range = Range1d(max(drilldepth), 0)
    p3.x_range = Range1d(0, max(mse))
    return p3

def plot_info(p4, info, pitch, roll):
    # --- generate a metadata table
    reqtext = [  # 14 details, 28 items
        'Sample ID:',
        info[0],
        'Concession:',
        info[1],
        'Feature:',
        info[2],
        'Date Drilled:',
        str(info[3].date()),
        'Easting:',
        str(info[4]),
        'Northing:',
        str(info[5]),
        'Start Time:',
        str(info[6].time().strftime('%H:%M')),
        'Stop Time:',
        str(info[7].time().strftime('%H:%M')),
        'Tool:',
        info[8],
        'Water Depth:',
        str(int(info[9])) + ' m',
        'Max Pitch:',
        str(round(max(pitch), 2)) + '°',
        'Max Roll:',
        str(round(max(roll), 2)) + '°',
        'Average Pitch:',
        str(round(np.mean(pitch), 2)) + '°',
        'Average Roll:',
        str(round(np.mean(roll), 2)) + '°',
        ]
    for i in range(0,28):
        if (i % 2) == 0:
            xloc = 20
            fstyle = 'bold'
        else:
            xloc = 60
            fstyle = None
        yloc = ylocs[int(i / 2)]
        p4.add_layout(Label(
            x = xloc, 
            y = yloc, 
            text = reqtext[i], 
            text_color = 'black', 
            text_font_size = '12px', 
            text_font_style = fstyle))
    return p4

def plot_speed(p5, depth, penratemm, toolsink):
    # --- generate a plot that displays drill penetration rate
    p5.line(penratemm, depth, line_width = 2, color = 'green')
    p5.y_range = Range1d(max(depth), 0)
    p5.x_range = Range1d(0, max(penratemm))
    p5.multi_polygons(
        fill_color = (242,242,242),
        line_color = 'black',
        line_width = 1,
        line_alpha = 1,
        xs=[[[[0, 0, max(penratemm), max(penratemm)]]]],
        ys=[[[[(toolsink), 0, 0, (toolsink)]]]],
        hatch_pattern = '\\',
        hatch_color = 'lightgrey',
        hatch_weight = 0.5,
        hatch_scale = 10,
    )
    return p5

# --- run a loop to create a dataframe of each sample
def create_plot(ttfile, template):
    # --- isolate the screenlog data of specific sample
    data = screenlog.loc[screenlog['Sample_ID'] == ttfile].iloc[0]
    units = []
    lines = []
    for u in range(1,number_units + 1): # attempt to read n number of units
        m = 'm' + str(u)
        unit = 'Unit_' + str(u)
        if pd.notnull(data[unit]):
            lines.append(int((data[m]) * 1000))
            units.append(data[unit])
    toolsink = (data['TS_Act'] * 1000)
    if toolsink < 0:
        toolsink = 0
    sampleid = data['Sample_ID']
    concession = data['Concession']
    feature = (data['Feature'])
    try:
        feature = str(int(feature))
    except:
        pass
    finally:
        feature = str(feature)
    datedrilled = data['Start_Date']
    easting = data['E_actual']
    northing = data['N_actual']
    timestart = data['Start_Time']
    timeend = data['End_Time']
    tool = data['Tool']
    waterdepth = data['WD_act']
    #lists of values from tooltrend file (12 items):
    drilldepth, torque, towerfeed, time, pitch, roll, force, mse, penrate, swivel, penratemm = read_holedata(ttfile)
    depth = [d + toolsink for d in drilldepth]
    info = [sampleid, concession, feature, datedrilled, easting, northing, timestart, timeend, tool, waterdepth, pitch, roll]
    p1 = plot_force_torque(template[0], drilldepth, torque, force)
    p2 = plot_units(template[1], depth, time, toolsink, towerfeed)
    p3 = plot_mse(template[2], mse, drilldepth)
    p4 = plot_info(template[3], info, pitch, roll)
    p5 = plot_speed(template[4], depth, penratemm, toolsink)

    # plot lines and units
    for i, line in enumerate(lines):
        if units[i] in gravels:
            (text_col, text_style, layeralpha)  = ('teal', 'bold', 0.1)
        else:
            (text_col, text_style, layeralpha)  = ('grey', 'italic', 0)
        linelocationp1 = (sum(lines[0:i+1]))
        linelocationp2 = (toolsink + sum(lines[0:i+1]))
        labellocationp1 = linelocationp1 - 0.5*lines[i]
        labellocationp2 = linelocationp2 - 0.5*lines[i]
        p1.multi_polygons(
        fill_color = 'teal',
        line_alpha = 0,
        fill_alpha = layeralpha,
        xs=[[[[0, 0, max(torque) + 25, max(torque) + 25]]]],
        ys=[[[[linelocationp1, (sum(lines[0:i])), (sum(lines[0:i])), linelocationp1]]]],
        )
        p2.multi_polygons(
        fill_color = 'teal',
        line_alpha = 0,
        fill_alpha = layeralpha,
        xs=[[[[0, 0, max(time), max(time)]]]],
        ys=[[[[linelocationp2, (sum(lines[0:i])) + toolsink, (sum(lines[0:i])) + toolsink, linelocationp2]]]],
        )
        p5.multi_polygons(
        fill_color = 'teal',
        line_alpha = 0,
        fill_alpha = layeralpha,
        xs=[[[[0, 0, max(penratemm), max(penratemm)]]]],
        ys=[[[[linelocationp2, (sum(lines[0:i])) + toolsink, (sum(lines[0:i])) + toolsink, linelocationp2]]]],
        )
        p3.multi_polygons(
        fill_color = 'teal',
        line_alpha = 0,
        fill_alpha = layeralpha,
        xs=[[[[0, 0, max(mse) + 25, max(mse) + 25]]]],
        ys=[[[[linelocationp1, (sum(lines[0:i])), (sum(lines[0:i])), linelocationp1]]]],
        )
        p1.add_layout(Span(location = linelocationp1, dimension = 'width', line_color = 'black', line_dash = 'dashed'))
        p2.add_layout(Span(location = linelocationp2, dimension = 'width', line_color = 'black', line_dash = 'dashed'))
        p3.add_layout(Span(location = linelocationp1, dimension = 'width', line_color = 'black', line_dash = 'dashed'))
        p5.add_layout(Span(location = linelocationp2, dimension = 'width', line_color = 'black', line_dash = 'dashed'))
        p1.add_layout(Label(x = 0.9*max(torque), y = labellocationp1, text = units[i], text_color = text_col, text_font_style = text_style, text_align = 'left', text_baseline = 'middle'))
        p2.add_layout(Label(x = max(time)/2, y = labellocationp2, text = units[i], text_color = text_col, text_font_style = text_style, text_align = 'center', text_baseline = 'middle'))
        p2.legend.location = 'top_center'

    # --- combine the three plots in one layout
    allplots = layout([[p4, p5, p2], [p1, p3]])
    allplots.width = 850
    allplots.height = 1225
    export_png(allplots, filename = 'plots/ttplot_' + file + '.png')
    success = '\nSuccessfully exported as: plots/ttplot_' + file + '.png\n\n\n'
    return success

# MAIN
with open(imdh, 'r') as heading:
    print('\n', heading.read())
print('\nTTPLOT SAMPLE TOOLTREND PLOT version 1.2.5')
print('Reading files . . .')
screenlogfile, tooltrendlist = check_dir()
print(len(tooltrendlist), 'Hole Data Files found')
print('Reading screenlog . . .')
screenlog = read_screenlog(screenlogfile)
for file in tooltrendlist:
    print('\nPlotting',file, '\n')
    template = template_layout()
    print(create_plot(file, template))
dur = dt.datetime.now() - init_time
dur = dur.seconds + (dur.microseconds / 1000000)
print(len(tooltrendlist), 'sample tooltrends plotted in', round(dur, 2), 'seconds')
print('\n\n[FINISHED]')

#for file in tooltrendlist:
#    os.remove(file + '.xlsx')
#for file in os.listdir():
#    if 'Screenlog' in file:
#        os.remove(file)

"""
FUNCTIONS:
remove files
"""