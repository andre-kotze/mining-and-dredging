# --- import required modules
import numpy as np
import pandas as pd
from bokeh.plotting import figure, curdoc
from bokeh.layouts import layout
from bokeh.io import export_png, curdoc
from bokeh.models import Range1d, LinearAxis, Label, Span, ColumnDataSource, ImageURL, SingleIntervalTicker
import os
import datetime as dt
from math import pi

# --- if the folder doesn't exist, create a folder to export plots into
if not os.path.exists('plots'):
    plots = os.mkdir('plots')

# --- define a list of gravel units
gravels = [
    'BLDG',
    'BLDRBL',
    'CBLG',
    'CBLGRBL',
    'CBLRBL',
    'CBLS',
    'CYG',
    'CYRBL',
    'G',
    'GRBL',
    'GS',
    'GSCY',
    'GSH',
    'PBLG',
    'PBLRBL',
    'PBLS',
    'PBLSHS',
    'PBLSST',
    'RBL',
    'RBLG',
    'RBLS',
    'RBLSH',
    'RBLSHS',
    'SG',
    'SHG',
    'SHGRBL',
    'SHPBLS',
    'SHRBL',
    'SHRBLG',
    'SHRBLS',
    'SRBL',
    'SSTRBL',
]

# --- check all files in directory
tooltrendlist = []
for file in os.listdir():
    # --- if the file is the screenlog, read relevant fields into memory
    if "Screenlog" in file:
        screenlog = pd.read_excel(
            file, engine='pyxlsb', usecols=[
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
                'Unit_1',
                'm1',
                'Unit_2',
                'm2',
                'Unit_3',
                'm3',
                'Unit_4',
                'm4',
                'Unit_5',
                'm5',
                'Unit_6',
                'm6',
                'Unit_7',
                'm7',
                'Unit_8',
                'm8',
                'Unit_9',
                'm9',
                'Unit_10',
                'm10'
                ])
    # --- if the file is the Python script, ignore it
    elif '.xlsx' in file:
        tooltrendlist.append(file[:-5])
    # --- if anything else, ignore
    else:
        pass
# --- trim screenlog dataframe to only contain samples for which tool trend files are present
trim = screenlog.Sample_ID.isin(tooltrendlist)
trimmedscreenlog = screenlog[trim]
# --- run a loop to create a dataframe of each sample
for file in tooltrendlist:
    # --- isolate the screenlog data of specific sample
    screenlogdata = trimmedscreenlog.loc[trimmedscreenlog['Sample_ID'] == file]
    # --- extract the tool sink of the sample
    toolsink = screenlogdata['TS_Act'].iloc[0]
    # --- extract unit thickness and depth
    units = []
    lines = []
    for u in range(1,11):
        m = 'm' + str(u)
        unit = 'Unit_' + str(u)
        if pd.notnull(screenlogdata[unit].iloc[0]):
            lines.append(int((screenlogdata[m].iloc[0]) * 1000))
            units.append(screenlogdata[unit].iloc[0])
    sampleid = screenlogdata['Sample_ID'].iloc[0]
    concession = screenlogdata['Concession'].iloc[0]
    feature = screenlogdata['Feature'].iloc[0]
    datedrilled = pd.TimedeltaIndex(screenlogdata['Start_Date'], unit = 'd') + dt.datetime(1899, 12, 30)
    easting = screenlogdata['E_actual'].iloc[0]
    northing = screenlogdata['N_actual'].iloc[0]
    timestart = pd.TimedeltaIndex(screenlogdata['Start_Time'], unit = 'd') + dt.datetime(1899, 12, 30)
    timeend = pd.TimedeltaIndex(screenlogdata['End_Time'], unit = 'd') + dt.datetime(1899, 12, 30)
    tool = screenlogdata['Tool'].iloc[0]
    waterdepth = screenlogdata['WD_act'].iloc[0]
       

    # --- extract tool trend data from tool trend file
    tooltrend = pd.read_excel(file + '.xlsx', sheet_name='HoleData', skiprows=1, usecols=[
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
    maxdepthindex = (tooltrend[tooltrend['Drill Depth\nmm'] == maxdepth].index.values[-1])
    moving_average = 5 # must be an odd integer
    if moving_average > 1:
        buff = int(moving_average / 2)
        speed = []
        for i in range(buff):
            speed.append(0)
        for i in range(buff,len(tooltrend) - buff):
            rmean = tooltrend.loc[i-buff:i+buff,'Feed Result\nmm/min'].sum() / moving_average
            speed.append(rmean)
        for i in range(buff):
            speed.append(0)
        tooltrend['Feed Result\nmm/min'] = speed
    tooltrend = tooltrend[tooltrend['Drill Depth\nmm'] >= 0]
    tooltrend = tooltrend[tooltrend.index <= maxdepthindex]
    tooltrend = tooltrend[tooltrend['Feed Result\nmm/min'] > 0]
    tooltrend = tooltrend[tooltrend['Tower Feed Setpoint\n%'] <= 100]
    (depth, drilldepth, torque, towerfeed, time, pitch, roll, force, mse, penrate, swivel, penratemm) = ([], [], [], [], [], [], [], [], [], [], [], [])
    counter = 0
    (h1, m1, s1) = str(tooltrend['Relative Time\nhh:mm:ss'].iloc[0]).split(':')
    result1 = (int(h1) * 60) + (int(m1)) + (int(s1) / 60)
    while counter < len(tooltrend.index):
        drilldepth.append(tooltrend['Drill Depth\nmm'].iloc[counter])
        depth.append(tooltrend['Drill Depth\nmm'].iloc[counter] + (toolsink * 1000))
        torque.append(tooltrend['Actual Drill Torque\nKN*m'].iloc[counter])
        towerfeed.append(tooltrend['Tower Feed Setpoint\n%'].iloc[counter])
        pitch.append(tooltrend['Pitch\n°'].iloc[counter])
        roll.append(tooltrend['Roll\n°'].iloc[counter])
        force.append(tooltrend['Rack Drive Lowering\nKN'].iloc[counter])
        swivel.append(tooltrend['Swivel RPM\nrpm'].iloc[counter])
        penrate.append(tooltrend['Feed Result\nmm/min'].iloc[counter] / 1000)
        mse.append(
            ((force[counter] / 5) + (0.4*pi)*(swivel[counter]*torque[counter]/(penrate[counter])))/1000
        )
        penratemm.append(tooltrend['Feed Result\nmm/min'].iloc[counter])
        t = str(tooltrend['Relative Time\nhh:mm:ss'].iloc[counter])
        (h, m, s) = t.split(':')
        result = (int(h) * 60) + (int(m)) + (int(s) / 60)
        time.append(result - result1)
        counter += 1


    # --- generate a plot of tool trends versus drill depth
    p1 = figure(
        x_axis_label = 'Torque (kNm)',
        y_axis_label = 'Drill Depth (mm)',
        plot_width = 500,
    )
    p1.outline_line_color = 'black'
    p1.outline_line_width = 1
    p1.line(torque, drilldepth, line_width = 2, color = 'red', legend_label = 'Torque')
    p1.y_range = Range1d(max(drilldepth), 0)
    p1.x_range = Range1d(0, max(torque) + 25)
    p1.extra_x_ranges = {'force' : Range1d(start = min(force), end = max(force)+ 25)}
    p1.line(force, drilldepth, line_width = 2, color = 'blue', x_range_name = 'force', legend_label = "Force")
    p1.add_layout(LinearAxis(x_range_name = 'force', axis_label = 'Force (kN)'), 'above')
    p1.toolbar_location = None
    p1.legend.location = 'top_left'


    # --- generate a plot of tool sink, units and total depth
    p2 = figure(
        x_axis_label = 'Total Drill Time (min)',
        y_axis_label = 'Total Depth (mm)',
        plot_width = 300
    )
    p2.toolbar_location = None
    p2.y_range = Range1d(max(depth), 0)
    p2.x_range = Range1d(0, max(time))
    p2.multi_polygons(
        fill_color = (242,242,242),
        line_color = 'black',
        line_width = 1,
        line_alpha = 1,
        xs=[[[[0, 0, max(time), max(time)]]]],
        ys=[[[[(toolsink*1000), 0, 0, (toolsink*1000)]]]],
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
    p2.extra_x_ranges = {'tower' : Range1d(start = 0, end = 104)}
    p2.line(towerfeed, depth, line_width = 2, color = 'black', x_range_name = 'tower', line_dash = 'dashed', legend_label = "Tower Feed")
    p2.add_layout(LinearAxis(x_range_name = 'tower', axis_label = 'Tower Feed Setpoint (%)'), 'above')

    # --- generate a plot that displays mechanical specific energy
    p3 = figure(
        x_axis_label = 'Mechanical Specific Energy (MPa)',
        y_axis_label = 'Drill Depth (mm)',
        plot_width = 300,
    )
    p3.outline_line_color = 'black'
    p3.outline_line_width = 1
    p3.line(mse, drilldepth, line_width = 2, color = 'orange')
    p3.y_range = Range1d(max(drilldepth), 0)
    p3.x_range = Range1d(0, max(mse))
    p3.toolbar_location = None

    # --- generate a plot that displays drill penetration rate
    p5 = figure(
        x_axis_type = None,
        x_axis_label = 'Penetration Rate (mm/min)',
        y_axis_label = 'Total Depth (mm)',
        plot_width = 150,
    )
    ticker = SingleIntervalTicker(interval = 100, num_minor_ticks = 5)
    xaxis = LinearAxis(ticker = ticker)
    p5.add_layout(xaxis, 'below', label = 'Penetration (mm/min)')
    p5.outline_line_color = 'black'
    p5.outline_line_width = 1
    p5.line(penratemm, depth, line_width = 2, color = 'green')
    p5.y_range = Range1d(max(depth), 0)
    p5.x_range = Range1d(0, max(penrate) * 1000)
    p5.toolbar_location = None
    p5.multi_polygons(
        fill_color = (242,242,242),
        line_color = 'black',
        line_width = 1,
        line_alpha = 1,
        xs=[[[[0, 0, max(penratemm), max(penratemm)]]]],
        ys=[[[[(toolsink*1000), 0, 0, (toolsink*1000)]]]],
        hatch_pattern = '\\',
        hatch_color = 'lightgrey',
        hatch_weight = 0.5,
        hatch_scale = 10,
    )

    for i, line in enumerate(lines):
        if units[i] in gravels:
            (text_col, text_style, layeralpha)  = ('teal', 'bold', 0.1)
        else:
            (text_col, text_style, layeralpha)  = ('grey', 'italic', 0)
        linelocationp1 = (sum(lines[0:i+1]))
        linelocationp2 = (toolsink * 1000 + sum(lines[0:i+1]))
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
        ys=[[[[linelocationp2, (sum(lines[0:i])) + toolsink*1000, (sum(lines[0:i])) + toolsink*1000, linelocationp2]]]],
        )
        p5.multi_polygons(
        fill_color = 'teal',
        line_alpha = 0,
        fill_alpha = layeralpha,
        xs=[[[[0, 0, max(penratemm), max(penratemm)]]]],
        ys=[[[[linelocationp2, (sum(lines[0:i])) + toolsink*1000, (sum(lines[0:i])) + toolsink*1000, linelocationp2]]]],
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


    # --- generate a plot that acts as a table containing meta data
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
    xlocs = [20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60, 20, 60]
    ylocs = [70, 70, 65, 65, 60, 60, 55, 55, 50, 50, 45, 45, 37, 37, 32, 32, 27, 27, 22, 22, 17, 17, 12, 12, 7, 7, 2, 2]
    reqtext = [
        'Sample ID:',
        str(sampleid),
        'Concession:',
        str(concession),
        'Feature:',
        str(feature),
        'Date Drilled:',
        str(datedrilled[0].date()),
        'Easting:',
        str(easting),
        'Northing:',
        str(northing),
        'Start Time:',
        str(timestart[0].time().strftime('%H:%M')),
        'Stop Time:',
        str(timeend[0].time().strftime('%H:%M')),
        'Tool:',
        str(tool),
        'Water Depth:',
        str(int(waterdepth)) + ' m',
        'Max Pitch:',
        str(round(max(pitch), 2)) + '°',
        'Max Roll:',
        str(round(max(roll), 2)) + '°',
        'Average Pitch:',
        str(round((sum(pitch) / len(pitch)), 2)) + '°',
        'Average Roll:',
        str(round((sum(roll) / len(roll)), 2)) + '°',
        ]
    for i, entry in enumerate(xlocs):
        if xlocs[i] == 20:
            fstyle = 'bold'
        else:
            fstyle = None
        p4.add_layout(Label(x = xlocs[i], y = ylocs[i], text = reqtext[i], text_color = 'black', text_font_size = '12px', text_font_style = fstyle))
        p4.add_layout(Label(x = 50, y = 75, text = 'THE EXPLORER', text_color = 'black', text_font_size = '20px', text_align = 'center'))
    
    
    url = "file:///Z:/Letterheads/IMDSA.png"
    source = ColumnDataSource(dict(
    url = [url],
    x1  = np.linspace(50, 50, 1),
    y1  = np.linspace(90, 90, 1),
    w1  = np.linspace(50, 50, 1),
    h1  = np.linspace(50.3*0.35, 50.30*0.35, 1), 
    ))
    logo = ImageURL(url="url", x="x1", y="y1", w='w1', h="h1", anchor="center")
    p4.add_glyph(source, logo)

    # --- combine the three plots in one layout
    allplots = layout([[p4, p5, p2], [p1, p3]])
    allplots.width = 850
    allplots.height = 1225
    export_png(allplots, filename = 'plots/ttplot_' + file + ".png")

#for file in tooltrendlist:
#    os.remove(file + '.xlsx')
#for file in os.listdir():
#    if 'Screenlog' in file:
#        os.remove(file)

