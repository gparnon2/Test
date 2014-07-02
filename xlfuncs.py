import xlsxwriter
import pandas as pd
import numpy as np
import os
import re


def custom_tabs(tab_dict, filename="output.xslx",finish=True,autosize=False, taborder = [], percentlist = [], currencylist = [], graphlist = []):
    # tab_dict is a dictionary of dictionaries
    # [{table_name: tablename, tab_name: tabname, value: df}]
    # currencylist is a list of the headers which should have currency formatting
    # percentlist is a list of the headers which should have percent formatting
    # graphlist is a list of dictionaries describing what graphs to make. The dictionaries are as follows: 
    # {tabname: tab name, tablename: name of table, xvar: x variable, type: graph type, series: list of series, number:entries to use, entered in "x:y", 
    # "x:", or ":y", where x and y are integers. x:y will get from x to y, x: will get from x to end, :y from start to y (optional input), 'gname': optional input for graph name}
    #woop woop
    
    workbook = xlsxwriter.Workbook(filename)
    graphcounter = 0
    mergedheaderformats = format_merged_header(workbook)
    ws2 = workbook.add_worksheet('Insights')
    ws2.hide_gridlines(option=2)
    ws2.set_column(0,0,2)
    tabcounter = [] #create a list of all the tabs in the excel doc to check validity
    for i in tab_dict:
        tabcounter.append(i['tab_name'])
    tabs = list_of_tabs(tab_dict, taborder, tabcounter) #Create a list of the tab names
    for tab in tabs:
        usedtitles = [] #used to keep track of which headers have been used
        tablenum = tabcounter.count(tab) #Check for multiple tables on a tab
        arraylength = 50 #set the length of the array of columns to check for sizing (arbitrary)
        lengths = [1]*arraylength
        ws = workbook.add_worksheet(tab)
        ws.hide_gridlines(option=2) #option 1, normal: option 2, no gridlines
        if not tablenum == 1:
            offset = 2
        else:
            offset = 1
        ws.set_column(0,0,2)
        for val in tab_dict:
            if val['tab_name'] == tab:
                hastotals = False #Tracker for if there is a totals row
                tbl = cleanDataFrame(val['value'])
                headers = tbl.columns
                if not tablenum == 1:
                    usedtitles = merge_header(ws, val, usedtitles, headers, offset, mergedheaderformats) #Merge the central header
                lengths = format_column_headers(headers, tablenum, workbook, ws, offset, lengths) #Format the column headers
                lengths = format_cells(tbl, workbook, headers, currencylist, percentlist, ws, lengths, offset) #format the cells
                for row in range(len(tbl)):
                    if tbl.iloc[row][0] == 'Total':  
                        hastotals = True #Set hastotals to true if there is a totals column
                for i in graphlist:
                    #Check whether the number input has been entered and sets the variable hasnumentries
                    #to 1 if there is an entry in "x:y" form, 2 if it is in ":x", and 3 if it is in "x:"
                    hasnumentries = 0
                    try: 
                        number = i['number']
                        colonloc = number.index(':')
                        hasnumentries = 1
                        beginning = number[:colonloc]
                        end = number[(colonloc + 1):]
                        #If the entry is in ":x" form
                        if len(beginning) == 0:
                            hasnumentries = 2
                        #If the entry is in "x:" form
                        if len(end) == 0:
                            hasnumentries = 3
                    except KeyError:
                        hasnumentries = 0
                    x = 0 #Create a variable to change the placement of the chart target data based on whether there is a totals row or not
                    if hastotals == False:
                        x = 1
                    #If there is one table on the page, just check for the correct tab and check if the table data needs to be altered. 
                    if tablenum == 1:
                        if (i['tabname'] == tab):
                            if hasnumentries == 0:
                                create_chart(headers,len(headers), offset+1, offset+row + 1 + x, workbook, ws2, graphcounter, i)
                                graphcounter = graphcounter + 1
                            elif hasnumentries == 1:
                                create_chart(headers,len(headers), offset + int(beginning) +1, offset+ 1 +int(end), workbook, ws2, graphcounter, i)
                                graphcounter = graphcounter + 1
                            elif hasnumentries == 2:
                                create_chart(headers,len(headers), offset + 1, offset+1 + int(end), workbook, ws2, graphcounter, i)
                                graphcounter = graphcounter + 1
                            else:
                                create_chart(headers,len(headers), offset + row - int(beginning), offset+ row + 1 + x, workbook, ws2, graphcounter, i)
                                graphcounter = graphcounter + 1
                    #Otherwise have to check tablename and tabname. If tablename is not specified catch the keyerror and use the header holder to 
                    #check the default table name. Also, check to see if the data needs to be altered to create the graph.
                    #Form is: if tab and table name match, if there are totals, if there is a numentries parameter: build this chart
                    else:
                        try:   
                            if (i['tabname'] == tab and i['tablename'] == val['table_name']):
                                if hasnumentries == 0:
                                    create_chart(headers,len(headers), offset+1, offset+row + 1 + x, workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                                elif hasnumentries == 1:
                                    create_chart(headers,len(headers), offset +int(beginning) + 1, offset+ 1 +int(end), workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                                elif hasnumentries == 2:
                                    create_chart(headers,len(headers), offset + 1, offset+ 1 +int(end), workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                                else:
                                    create_chart(headers,len(headers), offset+ row - int(beginning), offset+ row + 1 + x, workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                        except KeyError:
                            if (i['tabname'] == tab and i['tablename'] == usedtitles[-1]):
                                if hasnumentries == 0:
                                    create_chart(headers,len(headers), offset+1, offset+row + 1 + x, workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                                elif hasnumentries == 1:
                                    create_chart(headers,len(headers), offset +int(beginning) + 1, offset+ int(end) + 1, workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                                elif hasnumentries == 2:
                                    create_chart(headers,len(headers), offset + 1, offset+ int(end), workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                                else:
                                    create_chart(headers,len(headers), offset + 1 + row - int(beginning), offset+ row + 1 + x, workbook, ws2, graphcounter, i)
                                    graphcounter = graphcounter + 1
                offset = offset + row + 5
                # Length of full screen is 210
                columnnums = len(headers)
                columnsize = 210/columnnums
        if autosize == True:
            for x in range(arraylength):
                if lengths[x] < columnsize - 1:
                    ws.set_column(x+1,x+1,lengths[x] + 2)
                else:
                    ws.set_column(x+1,x+1,columnsize)             
    if finish == True: 
        workbook.close()
        os.system("open "+filename)
    else: return workbook
#Merge and center the header: format is from the format_merged_header function. If no table name is specified the
#error is caught and the leftmost column is set as the header unless it has been used already. In that case, the
#header is the left column plus the number of times that header has been used already
def merge_header(ws, val, usedtitles, headers, offset, mergedheaderformats):
    try:
        ws.merge_range(offset-1, 1, offset-1, len(headers), val['table_name'], mergedheaderformats['header'])
        usedtitles.append(val['table_name'])
    except KeyError:
        if not headers[0] in usedtitles:
            ws.merge_range(offset-1, 1, offset-1, len(headers), headers[0], mergedheaderformats['header'])
            usedtitles.append(headers[0])
        else:
            i = 1
            while i < 100:
                if not headers[0] + '-' + str(i) in usedtitles:
                    ws.merge_range(offset-1, 1, offset-1, len(headers), headers[0] + '-' + str(i), mergedheaderformats['header'])
                    usedtitles.append(headers[0] + '-' + str(i))
                    break
                else:
                    i = i + 1
    return usedtitles


def format_cells(tbl, workbook, headers, currencylist, percentlist, ws, lengths, offset):
    for row in range(len(tbl)):
        for col in range(len(tbl.iloc[row])):
            fmt = format_table(workbook)['table']
            value = tbl.iloc[row][col]
            #Add money format if the header is in the currency list
            if headers[col] in currencylist:
                if value >= 1000:
                    fmt.set_num_format('$#,##0')
                else:
                    fmt.set_num_format('$0.00')
            #Add percent format if the header is in the percent list
            elif headers[col] in percentlist:
                fmt.set_num_format('0.00%')
            elif value >= 1000: 
                fmt.set_num_format('#,##0')
            elif value < 1000 and value > 10:
                fmt.set_num_format('0.00')
            else: 
                fmt.set_num_format('0.00')
            #Add the length of the strings to an array to check for autosize length
            if len(str(value)) > lengths[col]:
                lengths[col] = len(str(value))  
            if tbl.iloc[row][0] == 'Total':  
                fmt.set_bg_color('silver')
            #Set hastotals to true if there is a totals column
                # hastotals = True
            #Put a left border on the leftmost column
            if col == 0:
                fmt.set_left()
            #Put a right border on the rightmost column
            if col == len(headers)-1:
                fmt.set_right()
            ws.write(offset+row+1,col + 1,value, fmt)
    return lengths
#This is the header formatting: Checks to see if it is the only table on the tab. If it is, the format uses a blue
#background and white font. Otherwise it uses the silver background and black font. There is also a border put on the
#far left and far right cells in order to complete the entire table border
def format_column_headers(headers, tablenum, workbook, ws, offset, lengths):
    for h in range(len(headers)):
        headerfmt = format_header(workbook)['header']
        #If there are multiple tables on a sheet, each one has a silver header format
        if not tablenum == 1:
            if h == 0:
                headerfmt.set_left()
                ws.write(offset,h+1,headers[h],headerfmt)
            elif h == len(headers) - 1:
                headerfmt.set_right()
                ws.write(offset,h+1,headers[h],headerfmt)
            else:        
                ws.write(offset,h+1,headers[h],headerfmt)
        else:
            #Otherwise the single table has a blue header format
            headerfmt.set_bg_color('335599')
            headerfmt.set_font_color('white')    
            if h == 0:
                headerfmt.set_left()
                ws.write(offset,h+1,headers[h],headerfmt)
            elif h == len(headers) - 1:
                headerfmt.set_right()
                ws.write(offset,h+1,headers[h],headerfmt)
            else:    
                ws.write(offset,h+1,headers[h],headerfmt)
        lengths[h] = len(headers[h])
    return lengths

#create a list of the tabnames (with multiples) to check if there are multiple countries on one tab. If there is a user input for tab order,
#put the list of tabs in that order, making sure that all of the user inputted tabs are actual tabs. Otherwise, go in an order decided by 
#XLSX writer. 

def list_of_tabs(tab_dict, taborder, tabcounter):
    if not taborder == []:
        taborder = inorder(taborder)
        taborder = [x for x in taborder if x in tabcounter]
        tabs = taborder
    else:
        tabs = list(set([i['tab_name'] for i in tab_dict]))
    return tabs

#Types of graphs are area, bar, column, line, pie, radar, scatter, stock.
def create_chart(headers, cols, startrow, endrow, workbook, ws2, numberofgraphs, namedict, filename="output.xslx",show=True, gname = None):
    letter = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O',
    'P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE',
    'AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS',
    'AT','AU','AV','AW','AX','AY','AZ']
    chart = workbook.add_chart({'type': namedict['type']})
    xcounter = 0
    while xcounter < cols:
        if headers[xcounter] == namedict['xvar']:
            break
        xcounter = xcounter + 1
    l2 = letter[xcounter]
    length3 = len(namedict['series'])
    for i in range(length3):
        counter = 0
        while counter < cols:
            if headers[counter] == namedict['series'][i]:
                break
            counter = counter + 1
        #Use the letter reference variable to convert the column number into the column letter
        l1 = letter[counter]
        #Add in each series from the tab inputted through a parameter, and from the start row to the end row, all submitted through parameters
        chart.add_series({'categories': '=' + namedict['tabname'] + '!' + l2 + str(startrow + 1) + ':' + l2 + str(endrow),
            'values':'=' + namedict['tabname'] + '!'+l1+ str(startrow + 1) + ':'+l1+str(endrow),
            'name': namedict['series'][i]})
    #Set the chart x axis name
    chart.set_x_axis({'name' : namedict['xvar']})
    #Set the chart size
    chart.set_size({'width':650, 'height':350})
    length8 = len(namedict['series'])
    try:
        gname = namedict['gname']
    except KeyError:
        gname = None
    if gname == None:
        l = 1
        title = namedict['series'][0]
        while l < length8:
            title = title + ', ' + namedict['series'][l]
            l = l+1
            chart.set_title({'name': title + ' over ' + namedict['xvar']})
    else:
        chart.set_title({'name': gname})
    #Graph spacing: for every even numbered graph put it in the B column, every odd numbered graph in the M column. Then put the graphs
    #2 accross, with every pair moving down 20 rows
    if numberofgraphs%2 == 0:
        ws2.insert_chart('B' + str((numberofgraphs*10) + 2), chart)
    else:
        ws2.insert_chart('M' + str(((numberofgraphs - 1)*10) + 2), chart)
    return
#Ground level format for each column header. Left aligned and vertically centered, bolded, 10 pt Times New Roman, silver background, top border, wrapped
def format_header(workbook):
    header = workbook.add_format()
    header.set_align('left')
    header.set_align('vcenter')
    header.set_bold()
    header.set_font_name('Times New Roman')
    header.set_bg_color('silver')
    header.set_font_size(10)
    header.set_text_wrap()
    header.set_top()
    formats = {'header':header}
    return formats
#Format for the merged header. Horizontally and vertically centered, 10 pt Times New Roman, blue background, bolded, white font color, full border, wrapped
def format_merged_header(workbook):
    header2 = workbook.add_format()
    header2.set_align('center')
    header2.set_align('vcenter')
    header2.set_font_name('Times New Roman')
    header2.set_bg_color('335599')
    header2.set_bold()
    header2.set_font_size(10)
    header2.set_font_color('white')
    header2.set_text_wrap()
    header2.set_border()
    formats = {'header':header2}
    return formats
#Ground level format for the data entries in the table. Left aligned, vertically centered, 10 pt Times New Roman, top and bottom border, text-wrapped
def format_table(workbook):
    fmt = workbook.add_format()
    fmt.set_align('left')
    fmt.set_align('vcenter')
    fmt.set_font_name('Times New Roman')
    fmt.set_font_size(10)
    fmt.set_top() #top border
    fmt.set_bottom() #bottom border
    fmt.set_text_wrap()
    formats = {'table':fmt}
    return formats

def cleanDataFrame(df):
    df = df.apply(lambda x: x.fillna(0))
    df = df.apply(lambda x: x.replace([np.inf, -np.inf], 0))
    for col in df:
        if type(df[col].iloc[0]) == str: 
            df[col] = df[col].apply(lambda x: removeNonAscii(x))
    return df

def tab_name(tab):
    return re.sub('[\[\]:*"?/]', '', str(tab))[:30]

def removeNonAscii(s):
    try: return "".join(i for i in s if ord(i)<128)
    except: return s

def inorder(seq):
    seen = set()
    seen_add = seen.add
    return [ x for x in seq if x not in seen and not seen_add(x)]
