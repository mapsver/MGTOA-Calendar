# MGTOA-Calendar

***Help doc***

- Currenly works on Python 3.3
- Confirm with Pradeep that we're using Tamil names or Sanskrit (Mula or Moolam, Thiruvathirai va Ardra)
- Confirm with Pradeep that we're using Phoenix as location
- Create a new folder for the new year calendar in the OutputArchive dir. 
	- Save a sample snapshot of the drikpanchang webpage every year to monitor any formatting changes	
- Clean the current folder of any old .txt or .xlsx output files. This is where the temp output files are written. 
- Execute parser_with_Selenium-rev2.py for extraction
	- Ensure that the 'parser_with_Selenium-rev2' script is set to Vakyam(SuryaSiddantha)	
	- REMEMBER: In the parser script, Set the start-date and end-date in the script	
	- if extraction was run multiple times.. make sure the data is consolidated into a single txt file called 'drikCalendarPHX-Vakyam.txt'..
- Execute excelWriter-rev2.py for excel output. 
	- REMEMBER: In the writer script, Change the 'globCol' value every year based on which day of the week Jan1st falls on. Note that Sunday = -1	
	- Requires all the input to be in a single txt file called 'drikCalendarPHX-Vakyam.txt'..
	- Rename the output XL file name with the correct year
- XL output Manual FORMATTING: 
	- Delete the 'first' tab

*****BUGS: Fix them manually*****
- Check for any sheets with the string 'ERROR'.
- In the final Excel output, look for cells that have 5 rows filled. In cases where the words are long it may spill over into adjacent cells which are subsequently overwritten.
- Quick lookup for long content. Open 'drikCalendarPHX-Vakyam.txt' and look for lines longer than 80chars. These are potential candidates for spillovers.
- Replace "until Full Night" with "Full Night"

*****Automation Todos*****
- Fix bug: Correct "upto Full Night" with "Full Night"
	- some how when writing "upto Full Night" .. some cells are missing the word "Night" (noticed only for Tithi)
