# UseTimePC
AutoIT Skript that log the PC - Power On Time to an ExcelSheet

Users needs only the executable file 

This is a alpha Version. 

This skript reads the System Events 153 and 50037 | best Choice 

* 153 -- PC Start Event (Kernel Boot)
* 50037 -- PC Shut Down Event (Dhcp close event)

Alternative you cn read the Application events
*  100 -- PC Start Event  ( hmpalertsvc started )
*  101 -- PC Shut Down Event ( hmpalertsvc stoped )

The Directory where the File resist must be writeable

Tested with Excel 2016 

ToDo: 
  * Dokumentation of the Excel Sheet
  * Read break and worktime from Config sheet. 
  
  
