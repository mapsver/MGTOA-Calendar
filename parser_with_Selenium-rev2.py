#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
# Author:      shiv
# Created:     23/10/2013
#-------------------------------------------------------------------------------

import lib2to3, os, datetime, time
from urllib.request import urlopen
from bs4 import BeautifulSoup
#from ghost import Ghost
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from collections import defaultdict

#lookupTagSet = ['Tithi', 'Nakshatra', 'Skipped Tithi', 'Skipped Nakshatra']
lookupTagSet = ['Tithi', 'Nakshathram', 'Skipped Tithi', 'Skipped Nakshathram']
#browser = webdriver.Chrome(ChromeDriverManager().install())

options = webdriver.ChromeOptions()
browser = webdriver.Chrome(options=options)

def main():

#set the calendar location to: Chandler(5289282), Phoenix(5308655)
	locationUrl="https://www.drikpanchang.com/location/panchang-city-finder.html?geoname-id=5308655"
	browser.get(locationUrl)
	time.sleep(0.5)

# Switch to Vakyam
	#firstDateUrl="http://www.drikpanchang.com/tamil/tamil-month-panchangam.html?date=16/07/2024"
	firstDateUrl= r"https://www.drikpanchang.com/tamil/tamil-month-panchangam.html?date=01/01/2025&time-format=24plushour"

	browser.get(firstDateUrl)
	browser.execute_script("dpSettingsToolbar.handlePanchangArithmeticOptionClick('suryasiddhanta', true)") # switches to Vakyam panchangam
	
	#browser.execute_script("handleTimeFormatOptionClick('24plushour')")
	#dont use this--- #browser.execute_script("switchToMoArith()")  # switches to Thiru Ganita panchangam
	time.sleep(0.5)

# Get started
	fname = 'drikCalendarPHX-Vakyam.txt'
	with open(fname, 'w') as outf:
		dateObj = datetime.datetime(2026,11,5)    # SET Start Date (yyyy, mm, dd)
		nextDate = dateObj.strftime("%d/%m/%Y")
		outf.write('Start-Time: ' + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
		outf.write("\n\nDate, Tithi, Nakshatra, Skipped Tithi, Skipped Nakshatra, \n")

		while nextDate != "01/01/2027":                 # SET End Date (dd/mm/yyyy)
			time.sleep(4.0)
			outf.write(nextDate + ', ')
			url = "http://www.drikpanchang.com/tamil/tamil-month-panchangam.html?date=" + nextDate + "&time-format=24plushour"
			returnPairs = getPairs(url)

			tithiValList = returnPairs.get(lookupTagSet[0], ' ')
			outf.write(tithiValList[0] + ', ')
			nakValList = returnPairs.get(lookupTagSet[1], ' ')
			outf.write(nakValList[0] + ', ')
			# Newer version of drikpanchang lists both Tithi and SkippedTithi under key='Tithi'
			#-- To support earlier versions, maintain skippedTithi logic also
			skipTithiValList = returnPairs.get(lookupTagSet[2], ' ')
			if skipTithiValList != ' ':
				outf.write(skipTithiValList[0] + ', ')
			elif len(tithiValList) == 2:
				outf.write(tithiValList[1] + ', ')
			else:
				outf.write(' , ')
			skipNakValList = returnPairs.get(lookupTagSet[3], ' ')
			if skipNakValList != ' ':
				outf.write(skipNakValList[0] + ', ')
			elif len(nakValList) == 2:
				outf.write(nakValList[1] + ', ')
			else:
				outf.write(' , ')

			outf.write('\n')

			dateObj = GetNextDate(dateObj)
			nextDate = dateObj.strftime("%d/%m/%Y")
		outf.write('\n\nEnd-Time: ' + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

def debugMain():

	#set the calendar location to: Chandler(5289282), Phoenix(5308655)
	locationUrl="https://www.drikpanchang.com/location/panchang-city-finder.html?geoname-id=5308655"
	browser.get(locationUrl)
	time.sleep(0.1)

	# Switch to Vakyam
	firstDateUrl="http://www.drikpanchang.com/tamil/tamil-month-panchangam.html?date=19/01/2015"
	browser.get(firstDateUrl)
	browser.execute_script("handlePanchangArithmeticOptionClick('suryasiddhanta')") # switches to Vakyam panchangam
	time.sleep(0.1)

	returnPairs = getPairs(firstDateUrl)
	firstDateUrl=""


def GetNextDate(currDate):
	return currDate + datetime.timedelta(1)

def getPairs(url):

	browser.get(url)
	#browser.save_screenshot('beforeCLick.png')
	#browser.execute_script("switchToSSArith()")
	#browser.save_screenshot('afterCLick.png')
	time.sleep(1.0)
	bs = BeautifulSoup(browser.page_source)

	returnPairs= defaultdict(list)
	ingreds = bs.find('div', {'class': 'dpPanchang'})
	allPanchEles = ingreds.find_all('span', {'class': 'dpElementKey'})
	allPanchValues = ingreds.find_all('span', {'class': 'dpElementValue'})
	for idx, el in enumerate(allPanchEles):
		elStr = el.text.strip()

		if elStr != '':
			for keyTag in lookupTagSet:
				if elStr.startswith(keyTag):
					valDivEl = allPanchValues[idx]
					keyValue = valDivEl.text.strip()
					returnPairs[keyTag].append(keyValue)
					break

	return returnPairs

#debugMain()
main()
browser.quit()

#debug:write to file
##    dataBefore = gh.content
##    with open("dataBeforeFile.txt", "w") as dataBeforeFile:
##        dataBeforeFile.write(str(dataBefore))
##    gh.evaluate("switchToSSArith();")
##    dataAfter = gh.content
##    with open("dataAfterFile.txt", "w") as dataAfterFile:
##        dataAfterFile.write(str(dataAfter))