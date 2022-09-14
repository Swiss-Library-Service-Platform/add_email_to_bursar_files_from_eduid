##########################################
#                                        #
#  SLSP tool to add a column with email  #
#                                        #
##########################################

DESCRIPTION
-----------
This tool is used to add a new colomn to a Excel file or a CSV file (tabulated) with
the uemail address.

To get the email, this Powershell script uses the Alma APIs and the primary id.

Author: Raphaël Rey
2022
Support: SLSP support ("Third Party Integration" area)
Licence: MIT
Version 1.1

HOW TO INSTALL IT
-----------------
- Get the following files and put them in the same folder:
	- add_email_from_eduid.ps1
	- .apikey_sample
- Get an Alma Users API key with read rights
- Rename the .apikey_sample file to .apikey
- Add the api key in the .apikey file


HOW TO USE IT
-------------
- Close Excel
- Open PowerShell
- Go to the script directory
- Get a xlsx or csv file (delimitor ";"), column UserID
- Type command:
	> .\add_email_from_eduid_csv.ps1 .\path_to_the_file_to_process.xlsx # for a xlsx file
	> .\add_email_from_eduid_csv.ps1 .\path_to_the_file_to_process.csv # for a csv file
	> .\add_email_from_eduid_csv.ps1 .\path_to_folder # for a folder

RESULT
------
The script creates a new file with the email of the users as the last column. The file
will be suffixed with "_processed.xlsx" or "_processed.csv".

Note: You can also indicate a directory name. The system will process all the files contained in.
Already processed files or not csv or xlsx files are ignored.

EXCEL FILE REQUIREMENTS
-----------------------
- Has a "UserID" column
- All headers are in row 1 and data starts in row 2

CSV FILE REQUIREMENTS
---------------------
- Has a "UserID" column
- Delimitor: ";"