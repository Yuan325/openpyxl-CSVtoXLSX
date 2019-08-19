# openpyxl from .csv to .xlsx

### GOAL
- converting SQL file output from csv to xlsx
- arranging sheets in xlsx via openpyxl


### WHY??
I am thinking of sending out multiple database information from SQL server to a user-friendly and easily-understandable excel workbook with multiple sheets (basically each sheet consist of a database).

However, SQL doesn't output workbook in such format.

### WHAT DO YOU NEED?
Before running this, I call on SQL query and output it into .csv file using batch file (.cmd)

```
set YYYY=%date:~10,4%
set MM=%date:~4,2%
set DD=%date:~7,2%
set today=%YYYY%%MM%%DD%

sqlcmd -S %SERVER% -i SQLQuery2.sql -o output\%today%_SQLoutput.csv -s ","
```
The title was set to "today" date for keeping track purposes (could be taken out if not applicable).
