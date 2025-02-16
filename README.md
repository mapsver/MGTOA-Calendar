# MGTOA-Calendar

***Help doc***

- Currenly works on Python 3.10.11
- The work is done by 2 scripts: one which does the data extraction from drikpanchang and the other which constructs the excel spreadheet using the extracted data. They need to be executed manually one after the other.
- Assumptions
	- Sript uses tamil Vakya panchangam from drikpanchang.com
	- uses Sanskrit names in the output (Mula becomes Moolam, Thiruvathirai becomes Ardra)
	- uses Phoenix as location
- Before executing the scripts
	- Create a new folder for the new year calendar in the OutputArchive dir. 
	- Save a sample snapshot of the drikpanchang webpage every year to monitor any formatting changes	
- Clean the current folder of any old .txt or .xlsx output files. This is where the temp output files are written. 
- Execute parser_with_Selenium-rev2.py for extraction
	- Ensure that the 'parser_with_Selenium-rev2' script is set to Vakyam(SuryaSiddantha)	
	- REMEMBER: In the parser script, Set the start-date and end-date in the script	
	- The extraction may stall at times which requires manual re-extraction from the stalled date.. change the start/end dates and re-run.
	- if extraction was run multiple times, make sure the data is consolidated into a single txt file called 'drikCalendarPHX-Vakyam.txt'.
- Execute excelWriter-rev2.py for excel output. 
	- Requires all the input to be in a single txt file in the same dir called 'drikCalendarPHX-Vakyam.txt'.
	- no changes required to this script
	- Open .xls output file in excel and Save as .xlsx	

- Archive the extracted .txt and final .xlsx in the OutputArchive dir

*****BUGS: Fix them manually*****
- Check for any sheets with the string 'ERROR'.
- In the final Excel output, look for cells that have 5 rows filled. In cases where the words are long it may spill over into adjacent cells which are subsequently overwritten.
- Quick lookup for long content. Open 'drikCalendarPHX-Vakyam.txt' and look for lines longer than 80chars. These are potential candidates for spillovers.
- Replace "until Full Night" with "Full Night"
- Resize the mini-calendars by setting width & height to 100%

*****Automation Todos*****
- Create a venv and requirements.txt
- Fix bug: Correct "upto Full Night" with "Full Night"
	- some how when writing "upto Full Night" .. some cells are missing the word "Night" (noticed only for Tithi)

*****Useful info*****
- row-height: 1pt in excel = 20pts in code (eg: header-row-height: 60.75 = 1215)
