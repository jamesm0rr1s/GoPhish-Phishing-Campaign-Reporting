# Must have the following files:
# Events.csv
# Results.csv

# Set the GoPhish CSV filenames
$eventsCsvFilename = "Events.csv"
$resultsCsvFilename = "Results.csv"

# Set the directory for the GoPhish CSV files (Last character must be a slash "\")
$goPhishOutputCsvDirectory = "Example Input - GoPhish CSV Files" + "\"

# Set the directory for the phishing reports (Last character must be a slash "\")
$phishingReportsOutputFolder = "Example Output - Phishing Reports" + "\"

# Write the start time
Write-Host (Get-Date -f yyyy_MM_dd-HH:mm:ss)"- Start"

# Check if the phishing reports directory exists
if(!(test-path $phishingReportsOutputFolder)){

    # Create the report directory
    New-Item -ItemType Directory -Force -Path $phishingReportsOutputFolder | Out-Null
}

# Save a report
Function SaveReport($filePath, $data, $reportNumber){

    # Set the path for the csv file
    $csvFilePath = $filePath + ".csv"

    # Export the file to a csv
    $data | Export-Csv -noType $csvFilePath

    # Set the path for the xlsx file
    $xlsxFilePath = $filePath + ".xlsx"

    # Create an Excel object
    $Excel = New-Object -ComObject Excel.Application

    # Show the Excel window
    $Excel.Visible = $FALSE

    # Do not show the "Save/Do Not Save" window
    $Excel.DisplayAlerts = $FALSE

    # Get the current directory
    $path = iex pwd

    # Open the workbook
    $Workbook = $Excel.Workbooks.Open($path.Path + "\" + $csvFilePath)

    # Open the worksheet
    $Worksheet = $Workbook.Worksheets.Item(1)

    # Bold the header row
    $Excel.Rows.Item(1).Font.Bold = $TRUE
    
    # Check if report 1
    if($reportNumber -eq 1){

        # Set the date format
        $Worksheet.Range("B:B").NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"

        # Loop through the specific columns
        foreach($character in "A","B","C","E","F","G","H","I","K"){

            # Autofit the columns
            $Worksheet.Columns($character).AutoFit() | Out-Null
        }

        # Loop through the specific columns
        foreach($character in "A","B","C","E","F","G","H","I"){

            # Center the cell
            $Worksheet.Columns($character).HorizontalAlignment = -4108
        }
    }

    # Check if report 2
    if($reportNumber -eq 2){

        # Loop through the specific columns
        foreach($character in "A","B","C","D","E","F","G","H"){

            # Autofit the columns
            $Worksheet.Columns($character).AutoFit() | Out-Null
        }

        # Loop through the specific columns
        foreach($character in "A","B","C","D","E","F","G","H"){

            # Center the cell
            $Worksheet.Columns($character).VerticalAlignment = -4108
            $Worksheet.Columns($character).HorizontalAlignment = -4108
        }
    }

    # Save the workbook as an xlsx file
    $Workbook.SaveAs($path.Path + "\" + $xlsxFilePath, 51)

    # Quit Excel
    $Excel.Quit()

    # Remove the temporary csv file
    Remove-Item $csvFilePath
}

# Search will take ~10 minutes for 10,000 emails instead of ~80 minutes using the Where-Object within the for loop
$Source = @"
using System;
using System.Management.Automation;
namespace FastSearch{
    public static class Search{
        public static object Find(PSObject[] collection, string column, string data){
            foreach(PSObject item in collection){
                if (item.Properties[column].Value.ToString() == data) { return item; }
            }
            return null;
        }
    }
}
"@
Add-Type -ReferencedAssemblies $Assem -TypeDefinition $Source -Language CSharp

# Set the path for the events csv file
$eventsCsvPath = $goPhishOutputCsvDirectory + $eventsCsvFilename

# Check for the tokens csv file
if((Test-Path $eventsCsvPath) -eq $false) {'"' + $eventsCsvPath + '"' + " not found."; exit}

# Set the header for the tokens.csv file
$eventsCsvHeader = "Email Address", "Timestamp", "Status", "Details"

# Import the tokens.csv without the header row
$eventsCsvFile = Import-Csv -Path $eventsCsvPath -Header $eventsCsvHeader | Where-Object "Email Address" -ne "email"

# Set the path for the results csv file
$resultsCsvPath = $goPhishOutputCsvDirectory + $resultsCsvFilename

# Check for results csv file
if((Test-Path $resultsCsvPath) -eq $false) {'"' + $resultsCsvPath + '"' + "not found."; exit}

# Import the results.csv without the header row
$resultsCsvFile = Import-Csv -Path $resultsCsvPath

# Add columns from results csv file
$eventsCsvFile | Add-Member -name "ID" -value "" -MemberType NoteProperty
$eventsCsvFile | Add-Member -name "First Name" -value "" -MemberType NoteProperty
$eventsCsvFile | Add-Member -name "Last Name" -value "" -MemberType NoteProperty
$eventsCsvFile | Add-Member -name "Full Name" -value "" -MemberType NoteProperty
$eventsCsvFile | Add-Member -name "IP Address" -value "" -MemberType NoteProperty
$eventsCsvFile | Add-Member -name "User Agent" -value "" -MemberType NoteProperty
$eventsCsvFile | Add-Member -name "Data" -value "" -MemberType NoteProperty

# Set the total row count
$rows = $eventsCsvFile.Count

# Set the row count to zero
$rowCount = 0

# Loop through each row
forEach($eventsCsvRow in $eventsCsvFile){

    # Lookup the results csv row
    $resultsCsvRow = [FastSearch.Search]::Find($resultsCsvFile, "email", $eventsCsvRow."Email Address")

    # Add results data
    $eventsCsvRow."ID" = $resultsCsvRow."id"
    $eventsCsvRow."First Name" = $resultsCsvRow."first_name"
    $eventsCsvRow."Last Name" = $resultsCsvRow."last_name"
    $eventsCsvRow."Full Name" = ($resultsCsvRow."first_name" + " " + $resultsCsvRow."last_name")

    # Adjust the timestamp
    $eventsCsvRow."Timestamp" = $eventsCsvRow."Timestamp".Replace("T"," ")
    $eventsCsvRow."Timestamp" = $eventsCsvRow."Timestamp".Replace("Z","")

    # Convert the details to json
    $eventsDetailsJson = ConvertFrom-Json $eventsCsvRow.Details

    # Check if the row is for the email opened or clicked link or submitted data
    if($eventsCsvRow.Status -eq "Email Opened" -OR $eventsCsvRow.Status -eq "Clicked Link" -OR $eventsCsvRow.Status -eq "Submitted Data"){

        # Set the IP address and user agent
        $eventsCsvRow."IP Address" = $eventsDetailsJson.browser."address"
        $eventsCsvRow."User Agent" = $eventsDetailsJson.browser."user-agent"
    }

    # Check if the row is for submitted data
    if($eventsCsvRow.Status -eq "Submitted Data"){

        # Set the submitted data
        $eventsCsvRow."Data" = $eventsDetailsJson.payload | Out-String
    }

    # Increment row count, get percent and show status
    $rowCount = $rowCount + 1
    $percents = [math]::round((($rowCount/($rows+1)) * 100), 0)
    Write-Progress -Activity:"Creating Report 1" -Status:"Created $rowCount of total $rows rows ($percents%)" -PercentComplete:$percents
}

# Set the filename for the PhishMe output file
$phishingCampaignResultsPhishMeFormatFilename = $phishingReportsOutputFolder + "GoPhish Phishing Campaign Report 1 - Details"

# Save the report
SaveReport $phishingCampaignResultsPhishMeFormatFilename $eventsCsvFile 1

# Create an array of objects
$arrayOfObjects = @()

# Get the email addresses
$emails = $eventsCsvFile."Email Address" | Select-Object -unique | Where-Object {$_ -ne ""}

# Set the total row count
$rows = $emails.Count

# Set the row count to zero
$rowCount = 0

# Loop through the emails
forEach($email in $emails){

    # Skip the first 
    #if($email -eq ""){continue}

    # Search for rows matching the email
    $tempRow = $eventsCsvFile | Where-Object {($_."Email Address" -eq $email)}
    $tempRowUnique = [FastSearch.Search]::Find($eventsCsvFile, "Email Address", $email)

    # Create a temporary object for this plugin
    $tempObject = New-Object PSObject

    # Check if the email was opened
    $openedEmail = if($tempRow.Status -match "Email Opened"){"Yes"}else{"No"}
    $clickedLink = if($tempRow.Status -match "Clicked Link"){"Yes"}else{"No"}
    $submittedData = if($tempRow.Status -match "Submitted Data"){"Yes"}else{"No"}

    # Add members to the object
    $tempObject | Add-Member -Name "ID" -Value $tempRowUnique.ID -MemberType NoteProperty
    $tempObject | Add-Member -Name "First Name" -Value $tempRowUnique."First Name" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Last Name" -Value $tempRowUnique."Last Name" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Full Name" -Value ($tempRowUnique."First Name" + " " + $tempRowUnique."Last Name") -MemberType NoteProperty
    $tempObject | Add-Member -Name "Email Address" -Value $email -MemberType NoteProperty
    $tempObject | Add-Member -Name "Opened Email" -Value $openedEmail -MemberType NoteProperty
    $tempObject | Add-Member -Name "Clicked Link" -Value $clickedLink -MemberType NoteProperty
    $tempObject | Add-Member -Name "Submitted Data" -Value $submittedData -MemberType NoteProperty

    # Update the array of objects
    $arrayOfObjects += $tempObject
    
    # Increment row count, get percent and show status
    $rowCount = $rowCount + 1
    $percents = [math]::round((($rowCount/($rows+1)) * 100), 0)
    Write-Progress -Activity:"Creating Report 2" -Status:"Created $rowCount of total $rows rows ($percents%)" -PercentComplete:$percents
}

# Set the filename for the PhishMe output file
$phishingCampaignResultsPhishMeFormatFilename = $phishingReportsOutputFolder + "GoPhish Phishing Campaign Report 2 - Summary"

# Save the report
SaveReport $phishingCampaignResultsPhishMeFormatFilename $arrayOfObjects 2

# Write the end time
Write-Host (Get-Date -f yyyy_MM_dd-HH:mm:ss)"- End"