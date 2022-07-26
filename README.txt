##########################################
#                                        #
#  SLSP tool to add a column with email  #
#                                        #
##########################################

DESCRIPTION
-----------
This tool is used to add a new colomn to an Excel file with the email address.
To get the email, this Powershell script uses the Alma APIs and the primary id.

Author: Raphaël Rey
2022
Support: SLSP support ("Third Party Integration" area)
Licence: MIT
Version 1.0

HOW TO USE IT
-------------
- Close Excel (to avoid conflicts with the script)
- Get an Alma Users API key with read rights
- Open the script "add_email_from_eduid.ps1" for edition
- Add the key in this file below "$API_KEY" in "Variables to be edited by the user" section
- Add the absolute path to the file to modify (somthing that should normally beginn with "C:\...")
- Save and run the script with Powershell

RESULT
------
The script creates a new file with the email of the users as the last column. The file
will be suffixed with "_processed.xlsx"

Note: if the script stops for any reason without finishing the process, it is possible to
restart it at any row. You have only to indicate the row number in the "$START_ON_ROW"
variable (see "Variables to be edited by the user" section)

EXCEL FILE REQUIREMENTS
-----------------------
- Has a "UserID" column
- All headers are in row 1 and data starts in row 2