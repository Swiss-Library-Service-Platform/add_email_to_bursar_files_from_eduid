##########################################
#                                        #
#  SLSP tool to add a column with email  #
#                                        #
##########################################

######################################
# Variables to be edited by the user #
######################################
$API_KEY = "READ_USERS_ALMA_API_KEY"
$EXCEL_FILE_ABSOLUTE_PATH = "C:\Users\<path>.xlsx"
$START_ON_ROW = 2 # Default value is 2


#####################
# Process beginning #
#####################

# Open Excel in background
$ExcelObj = New-Object -comobject Excel.Application

# Create nam of the output file
$EXCEL_FILE_ABSOLUTE_PATH_DESTINATION = ($EXCEL_FILE_ABSOLUTE_PATH -split "\.")[0] + '_processed.xlsx'

# Open excel file
$ExcelWorkBook = $ExcelObj.Workbooks.Open($EXCEL_FILE_ABSOLUTE_PATH)
$ExcelWorkBook.SaveAs($EXCEL_FILE_ABSOLUTE_PATH_DESTINATION)
$ws = $ExcelWorkBook.ActiveSheet

# Get the range of the data
$rowsCount = $ws.UsedRange.Rows.Count
$columnsCount = $ws.UsedRange.Columns.Count
$userIdCol = 0

# Find the UserId column
for ($i=1;$i -le $columnsCount;$i++){
	$colHeader = $ws.cells.Item(1, $i).Text
	if ( $colHeader -eq "UserID" ) { 
		$userIdCol = $i
	}
}

# Check if there is a UserId column
if ( $userIdCol -eq 0 ) { 
	echo "No UserID in the headers of the Excel file"
	echo "Press any key to exit..."
	Read-Host
	exit
}

# Add email row header
$emailCol = $columnsCount + 1
$ws.cells.Item(1, $emailCol).value = "Email"


Write-Host @"
##########################################
#                                        #
#  SLSP tool to add a column with email  #
#                                        #
##########################################

Number of rows: ${rowsCount}
Destination file: ${EXCEL_FILE_ABSOLUTE_PATH_DESTINATION}

Start process...
"@

for ($row=$START_ON_ROW;$row -le $rowsCount;$row++){
	
	# Find primary id in the row
	$userId = $ws.cells.Item($row, $userIdCol).Text
	
	if ( $userId.Length -lt 1 )
	{
		Write-Host "No userId found on row ${row}" -ForegroundColor red
		continue
	}
	
	Write-Host
	Write-Host "Row ${row}: get data for user ${userId} ..."
	
	try {
		$response = Invoke-WebRequest -Uri "https://api-eu.hosted.exlibrisgroup.com/almaws/v1/users/${userId}?apikey=${API_KEY}&format=json" | ConvertFrom-Json
	}
	catch{
		Write-Host "Failed to fetch data for user ${userId}" -ForegroundColor red
		continue
	}
	
	$emails = $response.contact_info.email
	$nbEmails = ($emails | Measure-Object).count
	$prefEmail = ""
	for ($i=0;$i -lt $nbEmails;$i++){
		if ( $emails[$i].preferred -eq "true" ) {
			$prefEmail = $emails[$i].email_address
			$ws.cells.Item($row, $emailCol).value = $prefEmail
			$ExcelWorkBook.Save()
			echo "${userId}: preferred email found: ${prefEmail}"
			break			
		}
	}
}
$ExcelWorkBook.close()
Write-Host
echo "Process finished"
echo "Press return to exit..."
Read-Host

