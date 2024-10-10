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
import copy

def Mbox(title, text, style):
	ctypes.windll.user32.MessageBoxW(0, text, title, style)

# SET THIS EVERY YEAR
currYear = ''
globRow = 3
globCol = "" # initialize to an invalid col-num
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
monthrowstyle = None
dayrowstyle = None
datestyle = None
greyedoutdatestyle = None
weekendgreyedoutdatestyle = None
nakstyle = None
tithistyle = None
eventstyle = None
weekenddatestyle = None
weekendnakstyle = None
weekendtithistyle = None
bottomRowStyle = None 
weekendbottomRowStyle = None
emptyWhitestyle = None
fbLinkstyle = None
MONTH_NAME_ROWNUM = 0
DAY_NAME_ROWNUM = 1
DAYS_OF_WEEK = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
WK5_ROW_IDX = 26
WK6_ROW_IDX = 32
LASTROW_WHITECELLS_COLIDX = 2
LASTROW_INFOBOX_COLIDX = 5

wbook = xlwt.Workbook()
sh1 = None #wbook.add_sheet("First")

def main():
	global globTithiInfo, globNakInfo, globSkipTithiInfo, globSkipNakInfo, currDay, currYear
	global globTithiTime, globNakTime, globSkipTithiTime, globSkipNakTime, globCol
	# Get started reading the csv
	fname = 'drikCalendarPHX-Vakyam.txt'
	##fname = 'drikCalendarPHX-TGanita.txt'

	with open(fname, 'r') as inf:
		data = inf.readlines()

		SetupFontsAndCellStyle()
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
			
			# debug particular date
			#if dateObj.day == 29 and dateObj.month == 4:
			#	print("here")
			####

			#set initial col position.
			# Columnnums for Sunday-to-Saturday is (-1 0 1 2 3 4 5) 
			# Datetime   for Sunday to Saturday is ( 6 0 1 2 3 4 5)
			if globCol == "":
				globCol = -1 if dateObj.weekday() == 6 else dateObj.weekday()	
				currYear = str(dateObj.year)
							
			if dateObj.day == 1:
				if dateObj.month != 1:
					addEmptyGreyCellsAtMonthEnd()
					populateLastRowDefaults()
					pass
				setupNewMonth(dateObj)
				addEmptyGreyCellsAtMonthBegn(dateObj, currDay if currDay else 31)
			getCellLocation(dateObj)
			currDay = dateObj.day

			cellData = ""

			tithiInfo = words[1].split(" ")
			tithiInfo = list(filter(None, tithiInfo))
			globTithiInfo = tithiInfo[0].strip()
			globTithiTime = getFormattedTime(tithiInfo)

			nakInfo = words[2].split(" ")
			nakInfo = list(filter(None, nakInfo))
			globNakInfo = nakInfo[0].strip()
			globNakTime = getFormattedTime(nakInfo)

			if (words[3].strip()):
				sktithiInfo = words[3].split(" ")
				sktithiInfo = list(filter(None, sktithiInfo))
				globSkipTithiInfo = sktithiInfo[0].strip()
				globSkipTithiTime = getFormattedTime(sktithiInfo)

			if (words[4].strip()):
				sknakInfo = words[4].split(" ")
				sknakInfo = list(filter(None, sknakInfo))
				globSkipNakInfo = sknakInfo[0].strip()
				globSkipNakTime = getFormattedTime(sknakInfo)

			#cellData = globTithiInfo + "\n" + globNakInfo + "\n" + globSkipTithiInfo + "\n" + globSkipNakInfo
			# WriteToExcelOld(cellData)
			#if WriteToExcel() == False:
			#    break;

			globTithiInfo = ConvertInfoToSanskrit(globTithiInfo)
			globNakInfo = ConvertInfoToSanskrit(globNakInfo)
			globSkipTithiInfo = ConvertInfoToSanskrit(globSkipTithiInfo)
			globSkipNakInfo = ConvertInfoToSanskrit(globSkipNakInfo)

			WriteToExcel()

	addFinishingTouchesForDec()
	inf.close()
	SetColumnWidth()
	##wbook.save("excel-2015-TGanita.xls")
	wbook.save("excel-" + currYear + "-Vakyam.xls")

def addFinishingTouchesForDec():
	global globCol
	globCol = -1 if globCol == 6 else globCol
	addEmptyGreyCellsAtMonthEnd()
	populateLastRowDefaults()

def populateLastRowDefaults():
	createWhiteCells()
	createInfoBox()

def createWhiteCells():
	global sh1, emptyWhitestyle
	sh1.merge(WK6_ROW_IDX, WK6_ROW_IDX+5, LASTROW_WHITECELLS_COLIDX,LASTROW_WHITECELLS_COLIDX+2)
	sh1.write(WK6_ROW_IDX, LASTROW_WHITECELLS_COLIDX, '', emptyWhitestyle)
	for row in range(WK6_ROW_IDX, WK6_ROW_IDX+6):
		for col in range(LASTROW_WHITECELLS_COLIDX,LASTROW_WHITECELLS_COLIDX+3):
			sh1.write(row,col, '', emptyWhitestyle)

def createInfoBox():
	global wbook, sh1, emptyWhitestyle, fbLinkstyle
	infoBoxContent = []
	infoBoxContent.append((("Notes: ", xlwt.easyfont('name Calibri, bold true, height 0x0DC')),
						("Calculations based on Vakya Panchangam",xlwt.easyfont('name Calibri, bold false, height 0x0DC'))))
	infoBoxContent.append((("Maha Ganapati Temple of Arizona", xlwt.easyfont('name Calibri, bold true, height 0x0F0')),
						("",xlwt.easyfont('name Calibri, bold false, height 0x0F0'))))
	infoBoxContent.append((("Addr: ", xlwt.easyfont('name Calibri, bold true, height 0x0C8')),
						 ("51293 W. Teel Road, Maricopa City, AZ 85139", xlwt.easyfont('name Calibri, bold false, height 0x0C8'))))
	infoBoxContent.append((("Phone: ", xlwt.easyfont('name Calibri, bold true, height 0x0C8')),
						 ("(520)568-9881", xlwt.easyfont('name Calibri, bold false, height 0x0C8'))))
	infoBoxContent.append((("Website: ", xlwt.easyfont('name Calibri, bold true, height 0x0C8')),
						 ("www.ganapati.org", xlwt.easyfont('name Calibri, bold false, height 0x0C8'))))	
	firstRowIdx = WK6_ROW_IDX
	for i,content in reversed(list(enumerate(infoBoxContent))):
		currCellStyle = copy.deepcopy(emptyWhitestyle)
		currCellStyle.borders.bottom_colour = xlwt.Style.colour_map['white']
		currCellStyle.borders.top_colour = xlwt.Style.colour_map['white']
		sh1.merge(firstRowIdx+i, firstRowIdx+i, LASTROW_INFOBOX_COLIDX,LASTROW_INFOBOX_COLIDX+1)
		sh1.write_rich_text(firstRowIdx+i, LASTROW_INFOBOX_COLIDX, infoBoxContent[i], currCellStyle)
		sh1.write(firstRowIdx+i, LASTROW_INFOBOX_COLIDX+1, "", currCellStyle)
	#fbLink = (("www.facebook.com/Mahaganapati", xlwt.easyfont('name Calibri, bold true, height 0x104')),
	#					 ("",xlwt.easyfont('name Calibri, bold false, height 0x0C8')))
	sh1.merge(firstRowIdx+5, firstRowIdx+5, LASTROW_INFOBOX_COLIDX,LASTROW_INFOBOX_COLIDX+1)
	sh1.write(firstRowIdx+5, LASTROW_INFOBOX_COLIDX, "www.facebook.com/Mahaganapati", fbLinkstyle)
	sh1.write(firstRowIdx+5, LASTROW_INFOBOX_COLIDX+1, "", fbLinkstyle)
	

def addEmptyGreyCellsAtMonthBegn(dateObj, lastDayOfPrevMonth):
	global wbook, sh1, globRow, globCol
	global greyedoutdatestyle, weekendgreyedoutdatestyle
	global nakstyle, weekendnakstyle
	
	outputRowNum = 2
	outputColNum = 0
	rowStartDate = lastDayOfPrevMonth - globCol
	createEmptyGreyCellsInRowUptoCol(outputRowNum, outputColNum, globCol, rowStartDate)
	rowStartDate = ''

def addEmptyGreyCellsAtMonthEnd():
	# add grey dates only for week5, no grey-dates for wk6 even if curr month ends in wk6
	global wbook, sh1, globRow, globCol
	global greyedoutdatestyle, weekendgreyedoutdatestyle
	global nakstyle, weekendnakstyle

	outputRowNum = globRow
	outputColNum = globCol + 1
	if outputRowNum == WK5_ROW_IDX:
		createEmptyGreyCellsInRowUptoCol(outputRowNum, outputColNum, 6, 1)
		outputColNum = 0
	createEmptyGreyCellsInRowUptoCol(WK6_ROW_IDX, outputColNum, 1, '')	
	sh1.row(WK6_ROW_IDX+5).height_mismatch = True
	sh1.row(WK6_ROW_IDX+5).height = 545 #5205
		

def createEmptyGreyCellsInRowUptoCol(outputRowNum, outputColNum, uptocolNum, rowStartDate):
	nextMonthDate = rowStartDate
	while outputColNum <= uptocolNum:
		currDateStyle = weekendgreyedoutdatestyle if outputColNum == 0 or outputColNum == 6 else greyedoutdatestyle	
		currEmptyCellStyle = weekendnakstyle if outputColNum == 0 or outputColNum == 6 else nakstyle
		sh1.write(outputRowNum, outputColNum, nextMonthDate, currDateStyle)		
		for i in range(1,5):
			sh1.write(outputRowNum+i, outputColNum, '', currEmptyCellStyle)
		currEmptyCellStyle = weekendbottomRowStyle if outputColNum == 0 or outputColNum == 6 else bottomRowStyle
		sh1.write(outputRowNum+5, outputColNum, '', currEmptyCellStyle)
		nextMonthDate = nextMonthDate+1 if nextMonthDate else ''
		outputColNum += 1
	pass

def SetColumnWidth():
	colWidth = 4699.53
	for i in range(0,12):
		sht = wbook.get_sheet(i)
		for j in range(0,7):
			sht.col(j).width = 5585 #5205

def setupNewMonth(dateObj):
	global globRow, globCol, sh1, currMonthName
	getNextMonth()
	sh1 = wbook.add_sheet(currMonthName, cell_overwrite_ok=True)
	createMonthHeader(dateObj)

def createMonthHeader(dateObj):
	global sh1, currMonthName, monthrowstyle
	sh1.merge(MONTH_NAME_ROWNUM, MONTH_NAME_ROWNUM, 0, 6)
	sh1.write(MONTH_NAME_ROWNUM, 0, currMonthName + ' ' + currYear, monthrowstyle)
	for i in range(1,7):
		sh1.write(MONTH_NAME_ROWNUM, i, '', monthrowstyle)
	sh1.row(MONTH_NAME_ROWNUM).height_mismatch = True
	sh1.row(MONTH_NAME_ROWNUM).height = 1215 #5205
	createDayNamesRow()

def createDayNamesRow():
	global sh1, dayrowstyle
	for i, dy in enumerate(DAYS_OF_WEEK):
		sh1.write(DAY_NAME_ROWNUM, i, dy, dayrowstyle)

def getCellLocation(dateObj):
	global globRow, globCol, sh1, currMonthName

	globCol += 1
	if (globCol > 6):
		globRow += 6
		globCol = 0
	if dateObj.day == 1:
		globRow = 2
	

def getFormattedTime(infoObj):
	if len(infoObj) > 3: # spl case for 'Full Night'
		return infoObj[2] + " " + infoObj[3]

	timeStr = infoObj[2]
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

def SetupFontsAndCellStyle():
	global monthrowstyle, dayrowstyle, bottomRowStyle, weekendbottomRowStyle, emptyWhitestyle
	global datestyle, weekenddatestyle, greyedoutdatestyle, weekendgreyedoutdatestyle
	global nakstyle, tithistyle, weekendnakstyle, weekendtithistyle, fbLinkstyle
	xlwt.add_palette_colour("weekday_bgColor", 0x20)
	wbook.set_colour_RGB(0x20, 254, 241, 184)
	xlwt.add_palette_colour("fblink_bgColor", 0x21)
	wbook.set_colour_RGB(0x21, 48, 84, 150)
	xlwt.add_palette_colour("weekend_bgColor", 0x22)
	wbook.set_colour_RGB(0x22, 248, 213, 104)
	xlwt.add_palette_colour("daysRow_bgColor", 0x23)
	wbook.set_colour_RGB(0x23, 131, 60, 12)
	xlwt.add_palette_colour("greydate_fontColor", 0x24)
	wbook.set_colour_RGB(0x24, 166, 166, 166)
	xlwt.add_palette_colour("fblink_fontColor", 0x25)
	wbook.set_colour_RGB(0x25, 48, 84, 150)
	
	#common pattern for weekends
	weekendBGPattern = xlwt.Pattern()
	weekendBGPattern.pattern = xlwt.Pattern.SOLID_PATTERN
	weekendBGPattern.pattern_fore_colour = xlwt.Style.colour_map['weekend_bgColor']	

	monthrowstyle = xlwt.XFStyle()
	monthrowFont = xlwt.Font()
	monthrowFont.name = 'Century Gothic'
	monthrowFont.height = 0x02D0
	monthrowBorder = xlwt.Borders()
	monthrowBorder.top = xlwt.Borders.THIN
	monthrowBorder.left = xlwt.Borders.THIN
	monthrowBorder.right = xlwt.Borders.THIN
	monthrowBorder.bottom = xlwt.Borders.THIN	
	monthrowBorder.bottom_colour = xlwt.Style.colour_map['black']
	monthrowBorder.right_colour = xlwt.Style.colour_map['black']
	monthrowAlign = xlwt.Alignment()
	monthrowAlign.horz = xlwt.Alignment.HORZ_RIGHT
	monthrowAlign.vert = xlwt.Alignment.VERT_TOP
	monthrowPattern = xlwt.Pattern()
	monthrowPattern.pattern = xlwt.Pattern.SOLID_PATTERN
	monthrowPattern.pattern_fore_colour = xlwt.Style.colour_map['weekday_bgColor']
	monthrowstyle.font = monthrowFont
	monthrowstyle.alignment = monthrowAlign
	monthrowstyle.borders = monthrowBorder
	monthrowstyle.pattern = monthrowPattern

	dayrowstyle = xlwt.XFStyle()
	dayrowFont = xlwt.Font()
	dayrowFont.name = 'Century Gothic'
	dayrowFont.height = 0x0F0
	dayrowFont.bold = True
	dayrowFont.colour_index = xlwt.Style.colour_map['white']
	dayrowBorder = xlwt.Borders()
	dayrowBorder.top = xlwt.Borders.THIN
	dayrowBorder.left = xlwt.Borders.THIN
	dayrowBorder.right = xlwt.Borders.THIN
	dayrowAlign = xlwt.Alignment()
	dayrowAlign.horz = xlwt.Alignment.HORZ_CENTER
	dayrowAlign.vert = xlwt.Alignment.VERT_TOP
	dayrowPattern = xlwt.Pattern()
	dayrowPattern.pattern = xlwt.Pattern.SOLID_PATTERN
	dayrowPattern.pattern_fore_colour = xlwt.Style.colour_map['daysRow_bgColor']
	dayrowstyle.font = dayrowFont
	dayrowstyle.alignment = dayrowAlign
	dayrowstyle.borders = dayrowBorder
	dayrowstyle.pattern = dayrowPattern

	datestyle = xlwt.XFStyle()
	dateFont = xlwt.Font()
	dateFont.name = 'Century Gothic'
	dateFont.height = 0x0118
	dateFont.bold = True
	dateBorder = xlwt.Borders()
	dateBorder.top = xlwt.Borders.THIN
	dateBorder.left = xlwt.Borders.THIN
	dateBorder.right = xlwt.Borders.THIN
	dateAlign = xlwt.Alignment()
	dateAlign.horz = xlwt.Alignment.HORZ_LEFT
	dateAlign.vert = xlwt.Alignment.VERT_TOP
	datePattern = xlwt.Pattern()
	datePattern.pattern = xlwt.Pattern.SOLID_PATTERN
	datePattern.pattern_fore_colour = xlwt.Style.colour_map['weekday_bgColor']	
	datestyle.font = dateFont
	datestyle.alignment = dateAlign
	datestyle.borders = dateBorder
	datestyle.pattern = datePattern
	
	weekenddatestyle = copy.deepcopy(datestyle)
	weekenddatestyle.pattern = weekendBGPattern
		
	greyedoutdatestyle = copy.deepcopy(datestyle)
	greyedoutdatestyle.font.colour_index = xlwt.Style.colour_map['greydate_fontColor']
	weekendgreyedoutdatestyle = copy.deepcopy(greyedoutdatestyle)
	weekendgreyedoutdatestyle.pattern = weekendBGPattern


	nakstyle = xlwt.XFStyle()
	nakFont = xlwt.Font()
	nakFont.name = 'Arial Narrow'
	nakFont.bold = True
	nakFont.height = 0x00C8
	nakAlign = xlwt.Alignment()
	nakAlign.horz = xlwt.Alignment.HORZ_CENTER
	nakAlign.vert = xlwt.Alignment.VERT_TOP
	nakBorder = xlwt.Borders()
	nakBorder.left = xlwt.Borders.THIN
	nakBorder.right = xlwt.Borders.THIN
	nakPattern = xlwt.Pattern()	
	nakPattern.pattern = xlwt.Pattern.SOLID_PATTERN
	nakPattern.pattern_fore_colour = xlwt.Style.colour_map['weekday_bgColor']	
	nakstyle.font = nakFont
	nakstyle.alignment = nakAlign
	nakstyle.borders = nakBorder
	nakstyle.pattern = nakPattern	

	weekendnakstyle = copy.deepcopy(nakstyle)
	weekendnakstyle.pattern = weekendBGPattern
		
	tithistyle = xlwt.XFStyle()
	tithiFont = xlwt.Font()
	tithiFont.name = 'Arial Narrow'
	tithiFont.bold = True
	tithiFont.height = 0x00C8
	tithiFont.colour_index = 0x3E
	tithiAlign = xlwt.Alignment()
	tithiAlign.horz = xlwt.Alignment.HORZ_CENTER
	tithiAlign.vert = xlwt.Alignment.VERT_TOP
	tithiBorder = xlwt.Borders()
	tithiBorder.left = xlwt.Borders.THIN
	tithiBorder.right = xlwt.Borders.THIN
	tithiPattern = xlwt.Pattern()	
	tithiPattern.pattern = xlwt.Pattern.SOLID_PATTERN
	tithiPattern.pattern_fore_colour = xlwt.Style.colour_map['weekday_bgColor']	
	tithistyle.font = tithiFont
	tithistyle.alignment = tithiAlign
	tithistyle.borders = tithiBorder
	tithistyle.pattern = tithiPattern

	weekendtithistyle = copy.deepcopy(tithistyle)
	weekendtithistyle.pattern = weekendBGPattern

	bottomRowStyle =  copy.deepcopy(tithistyle)
	bottomRowStyle.borders.bottom = xlwt.Borders.THIN
	weekendbottomRowStyle = copy.deepcopy(bottomRowStyle)
	weekendbottomRowStyle.pattern = weekendBGPattern
			
	emptyWhitestyle = xlwt.XFStyle()
	emptyWhiteBorder = xlwt.Borders()
	emptyWhiteFont = xlwt.Font()
	emptyWhiteFont.name = 'Calibri'
	emptyWhiteFont.bold = False
	emptyWhiteFont.height = 0x00C8
	emptyWhiteFont.colour_index = xlwt.Style.colour_map['black']
	emptyWhiteBorder.top = xlwt.Borders.THIN
	emptyWhiteBorder.left = xlwt.Borders.THIN
	emptyWhiteBorder.right = xlwt.Borders.THIN
	emptyWhiteBorder.bottom = xlwt.Borders.THIN
	#emptyWhiteBorder.bottom_colour = xlwt.Style.colour_map['black']
	emptyWhitePattern = xlwt.Pattern()	
	emptyWhitePattern.pattern = xlwt.Pattern.SOLID_PATTERN
	emptyWhitePattern.pattern_fore_colour = xlwt.Style.colour_map['white']
	emptyWhiteAlign = xlwt.Alignment()
	emptyWhiteAlign.horz = xlwt.Alignment.HORZ_LEFT
	emptyWhiteAlign.vert = xlwt.Alignment.VERT_TOP
	emptyWhitestyle.font = emptyWhiteFont
	emptyWhitestyle.borders = emptyWhiteBorder
	emptyWhitestyle.pattern = emptyWhitePattern
	emptyWhitestyle.alignment = emptyWhiteAlign
			
	fbLinkstyle = xlwt.XFStyle()
	fbLinkBorder = xlwt.Borders()
	fbLinkFont = xlwt.Font()
	fbLinkFont.name = 'Calibri'
	fbLinkFont.bold = True
	fbLinkFont.underline = True
	fbLinkFont.height = 0x0104
	fbLinkFont.colour_index = xlwt.Style.colour_map['white']
	fbLinkBorder.top = xlwt.Borders.THIN
	fbLinkBorder.left = xlwt.Borders.THIN
	fbLinkBorder.right = xlwt.Borders.THIN
	fbLinkBorder.bottom = xlwt.Borders.THIN
	fbLinkBorder.top_colour = xlwt.Style.colour_map['black']
	fbLinkPattern = xlwt.Pattern()	
	fbLinkPattern.pattern = xlwt.Pattern.SOLID_PATTERN
	fbLinkPattern.pattern_fore_colour = xlwt.Style.colour_map['fblink_fontColor']
	fbLinkAlign = xlwt.Alignment()
	fbLinkAlign.horz = xlwt.Alignment.HORZ_LEFT
	fbLinkAlign.vert = xlwt.Alignment.VERT_TOP
	fbLinkstyle.font = fbLinkFont
	fbLinkstyle.borders = fbLinkBorder
	fbLinkstyle.pattern = fbLinkPattern
	fbLinkstyle.alignment = fbLinkAlign

	eventstyle = xlwt.XFStyle()
	eventFont = xlwt.Font()
	eventFont.name = 'Arial Narrow'
	eventFont.bold = True
	eventFont.height = 0x00C8
	eventFont.colour_index = 0x3A
	eventAlign = xlwt.Alignment()
	eventAlign.horz = xlwt.Alignment.HORZ_CENTER
	eventAlign.vert = xlwt.Alignment.VERT_TOP
	eventBorder = xlwt.Borders()
	eventBorder.left = xlwt.Borders.THIN
	eventBorder.right = xlwt.Borders.THIN
	eventPattern = xlwt.Pattern()		
	eventPattern.pattern = xlwt.Pattern.SOLID_PATTERN
	eventPattern.pattern_fore_colour = xlwt.Style.colour_map['weekday_bgColor']	
	eventstyle.font = eventFont
	eventstyle.alignment = eventAlign
	eventstyle.borders = eventBorder
	eventstyle.pattern = eventPattern

	weekendeventstyle = copy.deepcopy(eventstyle)
	weekendeventstyle.pattern = weekendBGPattern

def addLowerBorderToStyle(currStyle):
	newStyle = copy.deepcopy(currStyle)
	newStyle.borders.bottom = xlwt.Borders.THIN
	return newStyle

def WriteToExcel():
	global wbook, sh1, globTithiInfo, globNakInfo, globSkipTithiInfo, globSkipNakInfo, currDay
	global globTithiTime, globNakTime, globSkipTithiTime, globSkipNakTime
	global datestyle, nakstyle, tithistyle, bottomRowStyle, weekendbottomRowStyle
	global weekenddatestyle, weekendnakstyle, weekendtithistyle

	nakCellOffset = 1
	tithiCellOffset = 3
	outputRowNum = globRow

	currDateStyle = weekenddatestyle if globCol == 0 or globCol == 6 else datestyle	
	sh1.write(outputRowNum, globCol, currDay, currDateStyle)

	outputRowNum = globRow + 1
	maxWordLen = 19

	currNakStyle = weekendnakstyle if globCol == 0 or globCol == 6 else nakstyle
	if len(globNakInfo) + len(globNakTime) > maxWordLen:
		sh1.write(outputRowNum, globCol, globNakInfo, currNakStyle)
		outputRowNum += 1
		sh1.write(outputRowNum, globCol, " until " + globNakTime, currNakStyle)
	else:
		sh1.write(outputRowNum, globCol, globNakInfo + " until " + globNakTime, currNakStyle)

	if globSkipNakInfo:
		outputRowNum += 1
		if len(globSkipNakInfo) + len(globSkipNakTime) > maxWordLen:
			sh1.write(outputRowNum, globCol, globSkipNakInfo, currNakStyle)
			outputRowNum += 1
			sh1.write(outputRowNum, globCol, " until " + globSkipNakTime, currNakStyle)
		else:
			sh1.write(outputRowNum, globCol, globSkipNakInfo + " until " + globSkipNakTime, currNakStyle)
	else:	
		outputRowNum += 1
		sh1.write(outputRowNum, globCol, "", currNakStyle)	

	currTithistyle = weekendtithistyle if globCol == 0 or globCol == 6 else tithistyle
	outputRowNum += 1
	currTithistyle = addLowerBorderToStyle(currTithistyle) if outputRowNum == globRow+5 else currTithistyle		
	if len(globTithiInfo) + len(globTithiTime) > maxWordLen:
		sh1.write(outputRowNum, globCol, globTithiInfo, currTithistyle)
		outputRowNum += 1
		currTithistyle = addLowerBorderToStyle(currTithistyle) if outputRowNum == globRow+5 else currTithistyle
		sh1.write(outputRowNum, globCol, " until " + globTithiTime, currTithistyle)
	else:
		sh1.write(outputRowNum, globCol, globTithiInfo + " until " + globTithiTime, currTithistyle)

	currTithistyle = addLowerBorderToStyle(currTithistyle) if outputRowNum+1 == globRow+5 else currTithistyle	
	if globSkipTithiInfo:
		outputRowNum += 1
		if len(globSkipTithiInfo) + len(globSkipTithiTime) > maxWordLen:
			sh1.write(outputRowNum, globCol, globSkipTithiInfo, currTithistyle)
			outputRowNum += 1
			currTithistyle = addLowerBorderToStyle(currTithistyle) if outputRowNum == globRow+5 else currTithistyle
			sh1.write(outputRowNum, globCol, " until " + globSkipTithiTime, currTithistyle)
		else:
			sh1.write(outputRowNum, globCol, globSkipTithiInfo + " until " + globSkipTithiTime, currTithistyle)
	else:
		if (outputRowNum-globRow < 5):
			outputRowNum += 1
			sh1.write(outputRowNum, globCol, "", currTithistyle)

	# error handling
	if (outputRowNum-globRow > 5):
		sh1.write(0,0, "ERROR in Month:" + currMonthName + "day:" + str(currDay) + " Row:" + str(outputRowNum) + " Col:" + str(globCol))
		return False;

	# fill the other rows with blank borders
	currBotRowstyle = weekendbottomRowStyle if globCol == 0 or globCol == 6 else bottomRowStyle	
	for i in range(outputRowNum+1, globRow+6):
		sh1.write(i, globCol, "", currBotRowstyle)
	sh1.row(globRow+5).height_mismatch = True
	sh1.row(globRow+5).height = 545 #5205

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
		return "Aslesha"
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
		return "Sashti"
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
		currMonthName = "JANUARY"
	elif (currMonthName == "JANUARY"):
		currMonthName = "FEBRUARY"
	elif (currMonthName == "FEBRUARY"):
		currMonthName = "MARCH"
	elif (currMonthName == "MARCH"):
		currMonthName = "APRIL"
	elif (currMonthName == "APRIL"):
		currMonthName = "MAY"
	elif (currMonthName == "MAY"):
		currMonthName = "JUNE"
	elif (currMonthName == "JUNE"):
		currMonthName = "JULY"
	elif (currMonthName == "JULY"):
		currMonthName = "AUGUST"
	elif (currMonthName == "AUGUST"):
		currMonthName = "SEPTEMBER"
	elif (currMonthName == "SEPTEMBER"):
		currMonthName = "OCTOBER"
	elif (currMonthName == "OCTOBER"):
		currMonthName = "NOVEMBER"
	elif (currMonthName == "NOVEMBER"):
		currMonthName = "DECEMBER"

def getCellLocationOld(dateObj):
	global globRow, globCol
	globRow += 1
	globCol += 0

main()
