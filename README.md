# MGTOA-Calendar

Help doc:

- Ensure that the 'parser_with_Selenium-rev2' script is set to Vakyam(SuryaSiddantha)
- Tamil names or Sanskrit (Mula or Moolam, Thiruvathirai va Ardra)
- Take a sample snapshot of the drikpanchang webpage every year to monitor any formatting changes
- REMEMBER: In the parser script, Set the start-date and end-date in the script
- REMEMBER: In the writer script, Change the 'globCol' value every year based on which day of the week Jan1st falls on. Note that Sunday = -1
- Execute parser_with_Selenium-rev2.py for extraction
- Execute excelWriter-rev2.py for excel output
- FORMATTING: Copy the first row (month name)

*****BUGS: Fix them manually****** 
- In the final Excel output, look for cells that have 5 rows filled. In cases where the words are long it may spill over into adjacent cells which are subsequently overwritten.
- Correct "until Full Night"


Automation Todos:
- Fix bug: Correct "until Full Night" with "Full Night"
- Fix bug: When all 
- Add DaysOfTheWeek row
- Add NameOfMonth row
