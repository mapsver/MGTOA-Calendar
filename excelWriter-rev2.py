#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      shiv
#
# Created:     29/09/2014
# Copyright:   (c) shiv 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import ctypes
import datetime
import xlwt

def Mbox(title, text, style):
    ctypes.windll.user32.MessageBoxW(0, text, title, style)

globRow = 1
##### CHANGE globCol EVERY YEAR. Sunday = -1.
globCol = 2 #Jan 1,2020 is a WED (so globCol = 2)
currMonthName = ""
currDay = ""
globTithiInfo = ""
globNakInfo = ""
globSkipTithiInfo = ""
globSkipNakInfo = ""
globTithiTime = ""
globNakTime = ""
globSkipTithiTime = ""
globSkipNakTime = ""
datestyle = None
nakstyle = None
tithistyle = None
eventstyle = None
firstSheetRow = 2

wbook = xlwt.Workbook()
sh1 = wbook.add_sheet("First")
firstSheet = sh1

def main():
    global globTithiInfo, globNakInfo, globSkipTithiInfo, globSkipNakInfo, currDay
    global globTithiTime, globNakTime, globSkipTithiTime, globSkipNakTime
    # Get started reading the csv
    fname = 'drikCalendarPHX-Vakyam.txt'
    ##fname = 'drikCalendarPHX-TGanita.txt'

    with open(fname, 'r') as inf:
        data = inf.readlines()

        SetupFonts()
        i = 0
        for line in data:
            i+=1
            if (i < 4):
                continue
            words = line.split(", ")
            try:
                dateObj = datetime.datetime.strptime(words[0], "%d/%m/%Y")
            except ValueError:
                break

            getCellLocation(dateObj)
            currDay = dateObj.day

            cellData = ""

            tithiInfo = words[1].split(" ")
            tithiInfo = list(filter(None, tithiInfo))
            globTithiInfo = tithiInfo[0].strip()
            globTithiTime = getFormattedTime(tithiInfo[2])

            nakInfo = words[2].split(" ")
            nakInfo = list(filter(None, nakInfo))
            globNakInfo = nakInfo[0].strip() # + " until " +
            globNakTime = getFormattedTime(nakInfo[2])

            if (words[3].strip()):
                sktithiInfo = words[3].split(" ")
                sktithiInfo = list(filter(None, sktithiInfo))
                globSkipTithiInfo = sktithiInfo[0].strip()
                globSkipTithiTime = getFormattedTime(sktithiInfo[2])

            if (words[4].strip()):
                sknakInfo = words[4].split(" ")
                sknakInfo = list(filter(None, sknakInfo))
                globSkipNakInfo = sknakInfo[0].strip()
                globSkipNakTime = getFormattedTime(sknakInfo[2])

            #cellData = globTithiInfo + "\n" + globNakInfo + "\n" + globSkipTithiInfo + "\n" + globSkipNakInfo
            # WriteToExcelOld(cellData)
            #if WriteToExcel() == False:
            #    break;

            globTithiInfo = ConvertInfoToSanskrit(globTithiInfo)
            globNakInfo = ConvertInfoToSanskrit(globNakInfo)
            globSkipTithiInfo = ConvertInfoToSanskrit(globSkipTithiInfo)
            globSkipNakInfo = ConvertInfoToSanskrit(globSkipNakInfo)

            WriteToExcel()

    inf.close()
    SetColumnWidth()
    ##wbook.save("excel-2015-TGanita.xls")
    wbook.save("excel-2020-Vakyam.xls")

def SetColumnWidth():
    colWidth = 4699.53
    for i in range(0,13):
        sht = wbook.get_sheet(i)
        for j in range(0,7):
            sht.col(j).width = 5585 #5205


def getCellLocation(dateObj):
    global globRow, globCol, sh1, currMonthName

    globCol += 1
    if (globCol > 6):
        globRow += 6
        globCol = 0

    if dateObj.day == 1:
        getNextMonth()
        sh1 = wbook.add_sheet(currMonthName, cell_overwrite_ok=True)
        globRow = 1

def getFormattedTime(timeStr):
    timeObj = timeStr.strip('+ ').split(':')

    if len(timeObj) != 2:
        return timeStr

    if (isinstance(timeObj[0], str) == False) or (isinstance(timeObj[1], str) == False):
        return timeStr

    newtimeStr = ""
    hr = int(timeObj[0])
    min = int(timeObj[1])
    nextDayStr = ""

    amPm = "AM"
    if hr > 23:
        hr -= 24
        amPm = "AM"
        nextDayStr = getDayOfWeek(globCol + 1)
    elif hr > 12:
        hr -= 12
        amPm = "PM"
    elif hr == 12:
        amPm = "PM"

    if hr == 0:
        hr = 12

    #newtimeStr = "{0}:{1} {2}".format(hr, min, amPm)
    newtimeStr = "%(hrStr)02d:%(minStr)02d %(merStr)s" % {'hrStr':hr, 'minStr':min, 'merStr':amPm}
    if (nextDayStr):
        newtimeStr += " " + nextDayStr

    return newtimeStr

def getDayOfWeek(dayNum):
    if dayNum == 0:
        return "Sun"
    if dayNum == 1:
        return "Mon"
    if dayNum == 2:
        return "Tue"
    if dayNum == 3:
        return "Wed"
    if dayNum == 4:
        return "Thu"
    if dayNum == 5:
        return "Fri"
    if dayNum == 6:
        return "Sat"
    if dayNum == 7:
        return "Sun"

def WriteToExcelOld(cellData):
    global wbook, sh1
    style = xlwt.XFStyle()
    style.alignment.wrap = 1
    sh1.write(globRow, globCol, cellData, style)

def SetupFonts():
    global datestyle, nakstyle, tithistyle

    datestyle = xlwt.XFStyle()
    dateFont = xlwt.Font()
    dateFont.name = 'Century Gothic'
    dateFont.height = 0x0118
    dateBorder = xlwt.Borders()
    dateBorder.top = xlwt.Borders.DOTTED
    dateBorder.left = xlwt.Borders.DOTTED
    dateBorder.right = xlwt.Borders.DOTTED
    dateAlign = xlwt.Alignment()
    dateAlign.horz = xlwt.Alignment.HORZ_LEFT
    datestyle.font = dateFont
    datestyle.alignment = dateAlign
    datestyle.borders = dateBorder

    nakFont = xlwt.Font()
    nakFont.name = 'Arial Narrow'
    nakFont.bold = True
    nakFont.height = 0x00C8
    nakAlign = xlwt.Alignment()
    nakAlign.horz = xlwt.Alignment.HORZ_CENTER
    nakBorder = xlwt.Borders()
    nakBorder.left = xlwt.Borders.DOTTED
    nakBorder.right = xlwt.Borders.DOTTED
    nakstyle = xlwt.XFStyle()
    nakstyle.font = nakFont
    nakstyle.alignment = nakAlign
    nakstyle.borders = nakBorder

    tithiFont = xlwt.Font()
    tithiFont.name = 'Arial Narrow'
    tithiFont.bold = True
    tithiFont.height = 0x00C8
    tithiFont.colour_index = 0x3E
    tithiAlign = xlwt.Alignment()
    tithiAlign.horz = xlwt.Alignment.HORZ_CENTER
    tithiBorder = xlwt.Borders()
    tithiBorder.left = xlwt.Borders.DOTTED
    tithiBorder.right = xlwt.Borders.DOTTED
    tithistyle = xlwt.XFStyle()
    tithistyle.font = tithiFont
    tithistyle.alignment = tithiAlign
    tithistyle.borders = tithiBorder

    eventFont = xlwt.Font()
    eventFont.name = 'Arial Narrow'
    eventFont.bold = True
    eventFont.height = 0x00C8
    eventFont.colour_index = 0x3A
    eventAlign = xlwt.Alignment()
    eventAlign.horz = xlwt.Alignment.HORZ_CENTER
    eventBorder = xlwt.Borders()
    eventBorder.left = xlwt.Borders.DOTTED
    eventBorder.right = xlwt.Borders.DOTTED
    eventstyle = xlwt.XFStyle()
    eventstyle.font = eventFont
    eventstyle.alignment = eventAlign
    eventstyle.borders = eventBorder

def WriteToExcel():
    global wbook, sh1, globTithiInfo, globNakInfo, globSkipTithiInfo, globSkipNakInfo, currDay
    global globTithiTime, globNakTime, globSkipTithiTime, globSkipNakTime
    global datestyle, nakstyle, tithistyle
    global firstSheet, firstSheetRow

    nakCellOffset = 1
    tithiCellOffset = 3
    outputRowNum = globRow

    sh1.write(outputRowNum, globCol, currDay, datestyle)

    outputRowNum = globRow + 1
    maxWordLen = 19

    if len(globNakInfo) + len(globNakTime) > maxWordLen:
        sh1.write(outputRowNum, globCol, globNakInfo, nakstyle)
        outputRowNum += 1
        sh1.write(outputRowNum, globCol, " until " + globNakTime, nakstyle)
    else:
        sh1.write(outputRowNum, globCol, globNakInfo + " until " + globNakTime, nakstyle)
        outputRowNum += 1

    if globSkipNakInfo:
        outputRowNum += 1
        if len(globSkipNakInfo) + len(globSkipNakTime) > maxWordLen:
            sh1.write(outputRowNum, globCol, globSkipNakInfo, nakstyle)
            outputRowNum += 1
            sh1.write(outputRowNum, globCol, " until " + globSkipNakTime, nakstyle)
        else:
            sh1.write(outputRowNum, globCol, globSkipNakInfo + " until " + globSkipNakTime, nakstyle)

    outputRowNum += 1
    if len(globTithiInfo) + len(globTithiTime) > maxWordLen:
        sh1.write(outputRowNum, globCol, globTithiInfo, tithistyle)
        outputRowNum += 1
        sh1.write(outputRowNum, globCol, " until " + globTithiTime, tithistyle)
    else:
        sh1.write(outputRowNum, globCol, globTithiInfo + " until " + globTithiTime, tithistyle)

    if globSkipTithiInfo:
        outputRowNum += 1
        if len(globSkipTithiInfo) + len(globSkipTithiTime) > maxWordLen:
            sh1.write(outputRowNum, globCol, globSkipTithiInfo, tithistyle)
            outputRowNum += 1
            sh1.write(outputRowNum, globCol, " until " + globSkipTithiTime, tithistyle)
        else:
            sh1.write(outputRowNum, globCol, globSkipTithiInfo + " until " + globSkipTithiTime, tithistyle)

    # error handling
    if (outputRowNum-globRow > 5):
        sh1.write(outputRowNum, 1, "ERROR in Month:" + currMonthName + " Row:" + str(globRow) + " Col:" + str(globCol), datestyle)
        return False;

    # fill the other rows with blank borders
    for i in range(outputRowNum+1, globRow+6):
        sh1.write(i, globCol, "", tithistyle)

    # reset global values
    globTithiInfo = ""
    globNakInfo = ""
    globSkipTithiInfo = ""
    globSkipNakInfo = ""
    globTithiTime = ""
    globNakTime = ""
    globSkipTithiTime = ""
    globSkipNakTime = ""

    return True;

def ConvertInfoToSanskrit(inputStr):
    if (inputStr == "Karthigai"):
        return "Krithika"
    if (inputStr == "Rohini"):
        return "Rohini"
    if (inputStr == "Mirugasirisham"):
        return "Mrigashirsha"
    if (inputStr == "Thiruvathirai"):
        return "Ardra"
    if (inputStr == "Punarpoosam"):
        return "Punarvasu"
    if (inputStr == "Poosam"):
        return "Pushya"
    if (inputStr == "Ayilyam"):
        return "Asresha"
    if (inputStr == "Magam"):
        return "Magha"
    if (inputStr == "Pooram"):
        return "PurvaPhalguni"
    if (inputStr == "Uthiram"):
        return "UttaraPhalguni"
    if (inputStr == "Hastham"):
        return "Hasta"
    if (inputStr == "Chithirai"):
        return "Chitra"
    if (inputStr == "Swathi"):
        return "Swati"
    if (inputStr == "Visakam"):
        return "Visakha"
    if (inputStr == "Pournami"):
        return "Purnima"
    if (inputStr == "Anusham"):
        return "Anuradha"
    if (inputStr == "Kettai"):
        return "Jyeshta"
    if (inputStr == "Moolam"):
        return "Mula"
    if (inputStr == "Pooradam"):
        return "PurvaAshada"
    if (inputStr == "Uthiradam"):
        return "UttaraAshada"
    if (inputStr == "Thiruvonam"):
        return "Shravana"
    if (inputStr == "Avittam"):
        return "Dhanishta"
    if (inputStr == "Sathayam"):
        return "Shatabhisha"
    if (inputStr == "Poorattathi"):
        return "PurvaBhadrapada"
    if (inputStr == "Uthirattathi"):
        return "UttaraBhadrapada"
    if (inputStr == "Revathi"):
        return "Revati"
    if (inputStr == "Aswini"):
        return "Ashwini"
    if (inputStr == "Bharani"):
        return "Bharani"

    if (inputStr == "Amavasai"):
        return "Amavasya"
    if (inputStr == "Pirathamai"):
        return "Pratipada"
    if (inputStr == "Thuthiyai"):
        return "Dwitiya"
    if (inputStr == "Thiruthiyai"):
        return "Tritiya"
    if (inputStr == "Sathurthi"):
        return "Chaturthi"
    if (inputStr == "Panjami"):
        return "Panchami"
    if (inputStr == "Shasti"):
        return "Shashti"
    if (inputStr == "Sapthami"):
        return "Saptami"
    if (inputStr == "Astami"):
        return "Ashtami"
    if (inputStr == "Navami"):
        return "Navami"
    if (inputStr == "Thasami"):
        return "Dashami"
    if (inputStr == "Egadashi"):
        return "Ekadashi"
    if (inputStr == "Duvadasi"):
        return "Dwadashi"
    if (inputStr == "Thirayodasi"):
        return "Trayodashi"
    if (inputStr == "Sathuradasi"):
        return "Chaturdashi"


def getNextMonth():
    global currMonthName

    if (currMonthName == ""):
        currMonthName = "Jan"
    elif (currMonthName == "Jan"):
        currMonthName = "Feb"
    elif (currMonthName == "Feb"):
        currMonthName = "Mar"
    elif (currMonthName == "Mar"):
        currMonthName = "Apr"
    elif (currMonthName == "Apr"):
        currMonthName = "May"
    elif (currMonthName == "May"):
        currMonthName = "Jun"
    elif (currMonthName == "Jun"):
        currMonthName = "Jul"
    elif (currMonthName == "Jul"):
        currMonthName = "Aug"
    elif (currMonthName == "Aug"):
        currMonthName = "Sep"
    elif (currMonthName == "Sep"):
        currMonthName = "Oct"
    elif (currMonthName == "Oct"):
        currMonthName = "Nov"
    elif (currMonthName == "Nov"):
        currMonthName = "Dec"

def getCellLocationOld(dateObj):
    global globRow, globCol
    globRow += 1
    globCol += 0

main()
