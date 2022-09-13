##########################################
#                                        #
#  SLSP tool to add a column with email  #
#                                        #
##########################################

# Fetch API Key in .\apikey file
$API_KEY = (Get-Content -Path .\.apikey).replace('API_KEY=', '').trim()

# Resolve provided path
$ABSOLUTE_PATH = Resolve-Path $args[0]

# Function to get the email address with api calls
function get_email {
	$prefEmail = ""
	Write-Host "Row ${row_counter}/${rowsCount}: get data for user ""${userId}"" ..."
	
	try {
		$response = Invoke-WebRequest -Uri "https://api-eu.hosted.exlibrisgroup.com/almaws/v1/users/${userId}?apikey=${API_KEY}&format=json" | ConvertFrom-Json
	}
	catch{
		Write-Host "Failed to fetch data for user ${userId}" -ForegroundColor red
		continue
	}
	
	$emails = $response.contact_info.email
	$nbEmails = ($emails | Measure-Object).count
	for ($i=0;$i -lt $nbEmails;$i++){
		if ( $emails[$i].preferred -eq "true" ) {
			$prefEmail = $emails[$i].email_address
			
			Write-Host "${userId}: preferred email found: ${prefEmail}"
			
			break
		}
	}
	return $prefEmail

}


# Get the list of to process, if only one, make a list of one file
if (Test-Path -Path $ABSOLUTE_PATH -PathType Container) {
	$filesToProcess = @()
	foreach ($fileName in (Get-ChildItem ./test_data).Name){
		
		# Ignore file already processed
		if ($fileName -like "*_processed*"){
			Write-Host "Ignore file ${ABSOLUTE_PATH}\${fileName}" -ForegroundColor red
			continue
		}
		
		# Ignore not xlsx or csv files
		$EXTENSION =  (Get-ChildItem "${ABSOLUTE_PATH}\${fileName}").Extension
		if (!($EXTENSION -eq ".csv") -and !($EXTENSION -eq ".xlsx")){
			Write-Host "Ignore file ${ABSOLUTE_PATH}\${fileName}" -ForegroundColor red
			continue
		}
		
		
		$filesToProcess += "${ABSOLUTE_PATH}\${fileName}"
	}
} else {
	$filesToProcess = @($ABSOLUTE_PATH)
}

foreach ($FILE_ABSOLUTE_PATH in $filesToProcess) {

	if ($FILE_ABSOLUTE_PATH -eq $null) {
		exit
	}

	$EXTENSION =  (Get-ChildItem $FILE_ABSOLUTE_PATH).Extension

	if ( $EXTENSION -eq ".csv" ){
		#########################
		# Process beginning CSV #
		#########################
		
		# Open csv file
		$csv = Import-Csv -Path $FILE_ABSOLUTE_PATH -Delimiter ';'
		
		# Check UserID column
		if (($csv | Get-Member -type NoteProperty -name UserID) -eq $null) {
			Write-Host
			Write-Host "No column ""UserID""." -ForegroundColor red
			exit
		}
		
		# Count number of rows
		$rowsCount = (Import-Csv $FILE_ABSOLUTE_PATH -Delimiter ';' | Measure-Object).count

		# Create name of the output file
		$FILE_ABSOLUTE_PATH_DESTINATION = ($FILE_ABSOLUTE_PATH -split "\.")[0] + '_processed.csv'	

	} elseif ( $EXTENSION -eq ".xlsx" ){
		###########################
		# Process beginning EXCEL #
		###########################
		
		# Open Excel in background
		$ExcelObj = New-Object -comobject Excel.Application

		# Create name of the output file
		$FILE_ABSOLUTE_PATH_DESTINATION = ($EXCEL_FILE_ABSOLUTE_PATH -split "\.")[0] + '_processed.xlsx'

		# Open excel file
		$ExcelWorkBook = $ExcelObj.Workbooks.Open($FILE_ABSOLUTE_PATH)
		$ExcelWorkBook.SaveAs($FILE_ABSOLUTE_PATH_DESTINATION)
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
			Write-Host "No UserID in the headers of the Excel file"
			Write-Host "Press any key to exit..."
			Read-Host
			exit
		}

		# Add email row header
		$emailCol = $columnsCount + 1
		$ws.cells.Item(1, $emailCol).value = "Email"
	} else {
		Write-Host
		Write-Host "Bad extension file ""${EXTENSION}"", must be "".csv"" or "".xlsx""" -ForegroundColor red
		exit
	}

	Write-Host @"
##########################################
#                                        #
#  SLSP tool to add a column with email  #
#                                        #
##########################################

Number of rows: ${rowsCount}
Source file: ${ABSOLUTE_PATH}
Destination file: ${FILE_ABSOLUTE_PATH_DESTINATION}

Start process...

"@


	if ( $EXTENSION -eq ".csv" ){
		#################
		# API calls CSV #
		#################
		$row_counter = 0
		foreach ($row in $csv)
		{
			$row_counter++
			$userId = $row.UserID
			
			# Get preferred email in prefEmail variable
			$prefEmail = get_email
			
			$row | Add-Member -NotePropertyName 'Email' -NotePropertyValue $prefEmail
			
		}
		Write-Host
		Write-Host "File saved to ${FILE_ABSOLUTE_PATH_DESTINATION}"
		$csv | Export-Csv $FILE_ABSOLUTE_PATH_DESTINATION -NoTypeInformation -Delimiter ';'
		

	} else {
		##################
		# API calls XLSX #
		##################
		for ($row_counter=2;$row_counter -le $rowsCount;$row_counter++){
	
			# Find primary id in the row
			$userId = $ws.cells.Item($row_counter, $userIdCol).Text
			
			if ( $userId.Length -lt 1 )
			{
				Write-Host "No userId found on row ${row}" -ForegroundColor red
				continue
			}
			
			# Get preferred email in prefEmail variable
			$prefEmail = get_email

			$ws.cells.Item($row_counter, $emailCol).value = $prefEmail
			$ExcelWorkBook.Save()
		}
		$ExcelWorkBook.close()
	}
}


Write-Host
Write-Host "Process finished"
Write-Host "Press return to exit..."
Read-Host

