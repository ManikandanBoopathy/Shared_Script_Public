<#
.SYNOPSIS

    This script will export data from Upgrade Readines ( Azure) in CSV format and the CSV will be imported ( BulCopy) into SQL
    
 .Author

    maniboopathii@gmail.com

 .Created Date
    
    2017-10-20

 .Modified Date
    2017-10-27
    
 .version
    1.0 - Modified Import process from "Foreach" option into "BulkCoy" which helped to reduce drastic amount of time ( from 25 mins to 5 seconds for 6k+ rows)
    2.0 - Instead of "Delete from TABELNAME" , I adopted "Truncate Tabel TABELNAME". This will speed up the flush the data and it will never create an event entry for each row of deletion    
     
 .Usage
    No Special parameter required
	
 .Requirement
	1. Install "WMF 5.1" because the Install-module will not be supported in older version
	2. Install below module using the command in PowerShell
		Install-Module AzureRM.OperationalInsights -Scope AllUsers
		Install-Module "sqlserver" -Scope AllUsers

 .To Create JSON file for Azure
	1. Set-ExecutionPolicy -ExecutionPolicy Unrestricted
	2. Import-Module azureRM.profile   ( if the module already installed )
	3. Login-AzureRmAccount   ( Enter user name and password when it asks - This is one time process to save JSON file for an automation)
	4. Save-AzureRmContext -Path <path>
	
 .Reference Web links
	https://blogs.technet.microsoft.com/privatecloud/2016/04/05/using-the-oms-search-api-with-native-powershell-cmdlets/
	https://arcanecode.com/2017/04/19/what-happened-to-save-azurermprofile/
#>

# ****************************** Section can be edit and need to provide proper variable *************************

$WorkingDirectory = "D:\SCRIPTS\Upgrade_Readiness_2_SQL"
Start-Transcript ($WorkingDirectory + "\" + "Transcript_Upgrade_Readiness.log")

$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Output "Script Started: $($StopWatch.Elapsed.ToString())"

$StopWatchLap1 = $StopWatch.Elapsed.ToString('hh\:mm\:ss')

#################
# Variable for working directory, UR profilepath
#################

$TempMoveStatus = ($WorkingDirectory + "\" + "TableMove_Temp2SQL.csv")
$URProfilePath = $WorkingDirectory + "\" + "UpgradeReadinessProfile.json"

$output = @()
$i = 1
$InsertTable = $null

#################
# Variable for Sending an Email
#################

 $TO = "maniboopathii@gmail.com"
 
 $Todaydate = $(Get-Date).ToString("dd/MM/yyyy")

#################
# Upgrade Readiness Informations and UR Query
#################
Import-AzureRmContext -Path $URProfilePath
$ResourceGroupName = "ResourceGroupName"
$WorkSpaceName = "WorkSpaceName"

$URQueryUAComputer = "(Type=UAComputer OR Type=UAUpgradedComputer)"
$URQueryUAApp = "Type = UAApp"
$URQueryUADriver = "Type = UADriver"
$URQueryUASysReqIssue = "Type = UASysReqIssue"

#####################
# SQL Database Informations and SQL Query
#####################

$database = 'database'
$server = "server"

$TableURUAComputer = "tblW10UpgradeReadiness_TEMP"
$TableURUAApp = "tblW10UpgradeReadinessUAAPP_TEMP"
$TableURUADriver = "tblW10UpgradeReadinessUADriver_TEMP"
$TableUASysReqIssue = "tblW10UpgradeReadinessUASysReqIssue_TEMP"

# ********************************** Editing section end ********************************

######################
# This Function will export CSV from Upgrade Readiness
######################

Function UR_Export_to_CSV ($URQuery, $CSVFileName) {
    Write-Output "$(get-date) | Running UR Query : $URQuery and output will save as $CSVFileName"
    if ($URQuery -like "*UAApp*") {
        #write-output "Gone into special if else"
        $URSearchResults = Get-AzureRmOperationalInsightsSearchResults -WorkspaceName $WorkSpaceName -ResourceGroupName $ResourceGroupName -Query $URQuery -Top 10000000
        $URSearchResults.Value | ConvertFrom-Json | Select SourceSystem, TimeGenerated, Computer, ComputerID, AppVendor, AppName, AppVersion, AppLanguage, TotalInstalls `
            , ComputersWithIssues, MonthlyActiveComputers, PercentActiveComputers, Issue, UpgradeAssessment, Importance, UpgradeDecision, ReadyForWindows, IsRollup, AppType, AppCategory `
            , TestPlan, TestResult, id, Type, MG, __metadata | Export-Csv -Path ($WorkingDirectory + "\" + $CSVFileName + ".csv" ) -NoTypeInformation
    }
    else {

        $URSearchResults = Get-AzureRmOperationalInsightsSearchResults -WorkspaceName $WorkSpaceName -ResourceGroupName $ResourceGroupName -Query $URQuery -Top 10000000
        $URSearchResults.Value | ConvertFrom-Json | Export-Csv -Path ($WorkingDirectory + "\" + $CSVFileName + ".csv" ) -NoTypeInformation
    }    
}

##########################
# Truncate Temp table before import data
# Reason: Truncating is much faster then Delete and it will never create events for each rows
##########################

Function Truncate-Table ($DeleteTable) {
    Write-Output "$(get-date) | Truncating table $DeleteTable"
    invoke-sqlcmd -Database $database -Query "Truncate Table $DeleteTable" -serverinstance $server
}

##########################
# Import CSV to SQL Temp table
##########################

Function Import_CSV_to_SQL ($TableName) {
    function Get-Type { 
        param($type) 
 
        $types = @( 
            'System.Boolean', 
            'System.Byte[]', 
            'System.Byte', 
            'System.Char', 
            'System.Datetime', 
            'System.Decimal', 
            'System.Double', 
            'System.Guid', 
            'System.Int16', 
            'System.Int32', 
            'System.Int64', 
            'System.Single', 
            'System.UInt16', 
            'System.UInt32', 
            'System.UInt64') 
 
        if ( $types -contains $type ) { 
            Write-Output "$type" 
        } 
        else { 
            Write-Output 'System.String' 
         
        } 
    } 
    function Out-DataTable { 
        [CmdletBinding()] 
        param([Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
 
        Begin { 
            $dt = new-object Data.datatable   
            $First = $true  
        } 
        Process { 
            foreach ($object in $InputObject) { 
                $DR = $DT.NewRow()   
                foreach ($property in $object.PsObject.get_properties()) {   
                    if ($first) {   
                        $Col = new-object Data.DataColumn   
                        $Col.ColumnName = $property.Name.ToString()   
                        if ($property.value) { 
                            if ($property.value -isnot [System.DBNull]) { 
                                $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                            } 
                        } 
                        $DT.Columns.Add($Col) 
                    }   
                    if ($property.Gettype().IsArray) { 
                        $DR.Item($property.Name) = $property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                    }   
                    else { 
                        $DR.Item($property.Name) = $property.value 
                    } 
                }   
                $DT.Rows.Add($DR)   
                $First = $false 
            } 
        }  
      
        End { 
            Write-Output @(, ($dt)) 
        } 
    }

    ##################
    # Used bulk copy to speed up copy process
    ##################

    try {
        Write-Output "$(get-date) | Importing CSV: $( $WorkingDirectory + "\" + $TableName + ".csv" ) into Table - $TableName"
        $CSVData = Import-CSV ( $WorkingDirectory + "\" + $TableName + ".csv" ) | Out-DataTable
        $SQLConnection = new-object System.Data.SqlClient.SqlConnection("Server=$server;Database=$database;Integrated Security=SSPI");
        $SQLConnection.Open()
        $BulkCopy = new-object ("System.Data.SqlClient.SqlBulkCopy") $SQLConnection
        $BulkCopy.DestinationTableName = $TableName
        $BulkCopy.WriteToServer($CSVData)
        $SQLConnection.Close()
    }
    catch { 
        $global:ImportError = "Error" 
     }
}

#######################
# Move the file to Archive folder and Rename CSV file that will import into SQL table
#######################

Function CSV_Rename ($FileName) {
    $NewFileName = $(Get-Date -F yyyy-MM-dd_HH-mm-ss)
    Write-Host "$(get-date) | Rename file from $($FileName + ".csv") to $($WorkingDirectory + "\Archive\" + $FileName + $NewFileName + ".csv") "
    Move-Item ($WorkingDirectory + "\" + $FileName + ".csv") ($WorkingDirectory + "\Archive\" + $FileName + ".csv")
    get-childitem ($WorkingDirectory + "\Archive\" + $FileName + ".csv") -ErrorAction SilentlyContinue | Rename-Item  -NewName ($WorkingDirectory + "\Archive\" + $FileName + $NewFileName + ".csv") -ErrorAction SilentlyContinue
}

Function Set-CellColor {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory, Position = 0)]
        [string]$Property,
        [Parameter(Mandatory, Position = 1)]
        [string]$Color,
        [Parameter(Mandatory, ValueFromPipeline)]
        [Object[]]$InputObject,
        [Parameter(Mandatory)]
        [string]$Filter,
        [switch]$Row
    )
    
    Begin {
        Write-Verbose "$(Get-Date): Function Set-CellColor begins"
        If ($Filter) {
            If ($Filter.ToUpper().IndexOf($Property.ToUpper()) -ge 0) {
                $Filter = $Filter.ToUpper().Replace($Property.ToUpper(), "`$Value")
                Try {
                    [scriptblock]$Filter = [scriptblock]::Create($Filter)
                }
                Catch {
                    Write-Warning "$(Get-Date): ""$Filter"" caused an error, stopping script!"
                    Write-Warning $Error[0]
                    Exit
                }
            }
            Else {
                Write-Warning "Could not locate $Property in the Filter, which is required.  Filter: $Filter"
                Exit
            }
        }
    }
    
    Process {
        $InputObject = $InputObject -split "`r`n"
        ForEach ($Line in $InputObject) {
            If ($Line.IndexOf("<tr><th") -ge 0) {
                Write-Verbose "$(Get-Date): Processing headers..."
                $Search = $Line | Select-String -Pattern '<th ?[a-z\-:;"=]*>(.*?)<\/th>' -AllMatches
                $Index = 0
                ForEach ($Match in $Search.Matches) {
                    If ($Match.Groups[1].Value -eq $Property) {
                        Break
                    }
                    $Index ++
                }
                If ($Index -eq $Search.Matches.Count) {
                    Write-Warning "$(Get-Date): Unable to locate property: $Property in table header"
                    Exit
                }
                Write-Verbose "$(Get-Date): $Property column found at index: $Index"
            }
            If ($Line -match "<tr( style=""background-color:.+?"")?><td") {
                $Search = $Line | Select-String -Pattern '<td ?[a-z\-:;"=]*>(.*?)<\/td>' -AllMatches
                $Value = $Search.Matches[$Index].Groups[1].Value -as [double]
                If (-not $Value) {
                    $Value = $Search.Matches[$Index].Groups[1].Value
                }
                If (Invoke-Command $Filter) {
                    If ($Row) {
                        Write-Verbose "$(Get-Date): Criteria met!  Changing row to $Color..."
                        If ($Line -match "<tr style=""background-color:(.+?)"">") {
                            $Line = $Line -replace "<tr style=""background-color:$($Matches[1])", "<tr style=""background-color:$Color"
                        }
                        Else {
                            $Line = $Line.Replace("<tr>", "<tr style=""background-color:$Color"">")
                        }
                    }
                    Else {
                        Write-Verbose "$(Get-Date): Criteria met!  Changing cell to $Color..."
                        $Line = $Line.Replace($Search.Matches[$Index].Value, "<td style=""background-color:$Color"">$Value</td>")
                    }
                }
            }
            Write-Output $Line
        }
    }    
    End {
        Write-Verbose "$(Get-Date): Function Set-CellColor completed"
    }
}
   
Function SendMail {
    #########################################
    #This function will read CSV file and send mail to multiple recpients in HTML format
    #########################################

    $htmlformat = '<title>Upgrade Readiness 2 SQL</title>'
    $htmlformat += '<style>'
    $htmlformat += 'td {text-align:center;}'
    $htmlformat += 'TABLE{border-width: 3px;border-style: solid;border-color: black;border-collapse: collapse;}'
    $htmlformat += 'TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:#D6D6F5}'
    $htmlformat += 'TD{border-width: 1px;padding: 8px;border-style: solid;border-color: black;}'
    $htmlformat += '</style>'
    $timezone = (Get-WMIObject -class Win32_TimeZone).caption

    ############################################
    #Converting CSV into HTML format
    ############################################
    $EmailBody = Import-Csv -Path $TempMoveStatus | ConvertTo-Html -PostContent "<h6><font color='#808080'>This automation export the data from Upgrade Readiness and import into SQL database<br>If any error during import and this will marked as not imported</br></font></h6>" `
        -Head $htmlformat -Body "<font face='Calibri'>Hi Admins,<br>Upgrade Readiness 2 SQL DB complete report on $Todaydate</br><font face='Calibri'><br>Script start time - <strong>$($StopWatchLap1)</strong><font face='Calibri'>H<br>Script elapsed time - <strong>$($StopWatchLap2)</strong></br><br>This report generated from $env:COMPUTERNAME<br>$timezone</font></br><br>" `
        |Set-CellColor -Property Temp_TableName -Color red -Filter "Temp_TableName -like '*CSV*'" -Row `
        | Out-String
    Send-MailMessage -From Automation@Dummy.com -Subject "Upgrade Readiness 2 SQL" -To $TO -Body $EmailBody -BodyAsHtml `
        -SmtpServer SMTPServer.com -Attachments $TempMoveStatus , ($WorkingDirectory + "\" + "Transcript_Upgrade_Readiness.log")
						 
    Write-Output "$(Get-Date) | Sending mail to Admins"		
}

##################
# Check the CSV is available, if it is available and it will be moved to archive
##################

if ((Test-Path ($WorkingDirectory + "\" + "*.csv") ) -eq $true) {
    Write-Output "$(get-date) | Old CSV files are found and proceeding to rename & archive"
    CSV_Rename $TableURUAComputer
    CSV_Rename $TableURUAApp
    CSV_Rename $TableURUADriver
    CSV_Rename $TableUASysReqIssue
}

###################
# Exporting CSV file from UR
###################

Write-Output "$(get-date) | Running UR Query and exporting into CSV"
UR_Export_to_CSV $URQueryUAComputer $TableURUAComputer
UR_Export_to_CSV $URQueryUAApp $TableURUAApp
UR_Export_to_CSV $URQueryUADriver $TableURUADriver
UR_Export_to_CSV $URQueryUASysReqIssue $TableUASysReqIssue
Write-Output "Total elapsed Time to export CSV: $($StopWatch.Elapsed.ToString())"

####################
# Deleting temp table value before import it
####################

Write-Output "$(get-date) | Truncating existing data from the table"
Truncate-Table $TableURUAComputer
Truncate-Table $TableURUAApp
Truncate-Table $TableURUADriver
Truncate-Table $TableUASysReqIssue

####################
# Importing data from CSV to SQL temp table
####################

#$StopWatchImportDB = [System.Diagnostics.Stopwatch]::StartNew()
Write-Output "Stop watch start to import CSV 2 SQL: $($StopWatch.Elapsed.ToString())"
Write-Output "$(get-date) | Importing CSV into SQL"

Import_CSV_to_SQL $TableURUAComputer
Import_CSV_to_SQL $TableURUAApp
Import_CSV_to_SQL $TableURUADriver
Import_CSV_to_SQL $TableUASysReqIssue

Write-Output "Total elapsed Time to import CSV 2 SQL: $($StopWatch.Elapsed.ToString())"

#####################
# Comparing CSV count with SQL temp table count 
#####################

$Temp_Table = ($TableURUAComputer, $TableURUAApp, $TableURUADriver, $TableUASysReqIssue)

Foreach ($Temp_Table in $Temp_Table) {
    $SNO = $i++
    if (((Import-Csv ($WorkingDirectory + "\" + $Temp_Table + ".csv")).count) -eq ((Invoke-Sqlcmd -query "SELECT count(*) FROM $Temp_Table" -ServerInstance $server -Database $database) | select column1 -ExpandProperty column1)) {
        
        $CSVCount = (Import-Csv ($WorkingDirectory + "\" + $Temp_Table + ".csv")).count
        $SQLTableCount = (Invoke-Sqlcmd -query "SELECT count(*) FROM $Temp_Table" -ServerInstance $server -Database $database) | select column1 -ExpandProperty column1
        Write-Output "$(get-date) | Count of $($Temp_Table + ".csv") - $CSVCount and count of $($Temp_Table) - $SQLTableCount"
        #Write-Host "Matched"
        $NewTable = $Temp_Table.Split("_")[0]
        
        ##############
        # Truncating production table data before INSERT from temp to Prod
        ##############
        try {
            Write-Output "$(get-date) | Truncate data from $NewTable"
            Invoke-Sqlcmd -query "Truncate Table $NewTable" -ServerInstance $server -Database $database
        }
        catch {
            $DeleteTable = "Error"
        }
        try {
            Write-Output "$(get-date) | INSERT INTO $NewTable from $Temp_Table"
            Invoke-Sqlcmd -query "INSERT INTO $NewTable select * from $Temp_Table" -ServerInstance $server -Database $database
        }
        catch {
            $InsertTable = "Error"
        }

        ##############
        # Checking if any error occurred during INSERT from temp table to Prod table
        ##############

        if (($InsertTable -match "Error") -or ($InsertTable -like "Error")) {
            $output += New-Object PSObject -Property @{ "S.No" = $SNO
                Temp_TableName = $Temp_Table
                Prod_TableName = $NewTable
                Count_of_CSV = $CSVCount
                Count_of_SQL_Table = $SQLTableCount
                Status = "Failed to move table"
            } | select S.No, Count_of_CSV, Count_of_SQL_Table, Status, Temp_TableName, Prod_TableName
        }
        else {
            $output += New-Object PSObject -Property @{ "S.No" = $SNO
                Temp_TableName = $Temp_Table
                Prod_TableName = $NewTable
                Count_of_CSV = $CSVCount
                Count_of_SQL_Table = $SQLTableCount
                Status = "Table moved successfully"
            } | select S.No, Count_of_CSV, Count_of_SQL_Table, Status, Temp_TableName, Prod_TableName
        }
    }
    else {
        #Write-Host "Value not matched"
        $CSVCount = (Import-Csv ($WorkingDirectory + "\" + $Temp_Table + ".csv")).count
        $SQLTableCount = (Invoke-Sqlcmd -query "SELECT count(*) FROM $Temp_Table" -ServerInstance $server -Database $database) | select column1 -ExpandProperty column1
        Write-Output "$(get-date) | Count of $($Temp_Table + ".csv") - $CSVCount and count of $($Temp_Table) - $SQLTableCount"

        $output += New-Object PSObject -Property @{ "S.No" = $SNO
            Temp_TableName = "CSV count & Temp table count not matched"
            Prod_TableName = "CSV count & Temp table count not matched"
            Count_of_CSV = $CSVCount
            Count_of_SQL_Table = $SQLTableCount
            Status = "Didn't attempt to move"
        } | select S.No, Count_of_CSV, Count_of_SQL_Table, Status, Temp_TableName, Prod_TableName
    }
    
}
$output | Export-Csv $TempMoveStatus -ErrorAction SilentlyContinue -NoTypeInformation

Write-Output "Total elapsed Time to complete script: $($StopWatch.Elapsed.ToString())"
$StopWatchLap2 = $StopWatch.Elapsed.ToString('hh\:mm\:ss')
$StopWatch.Stop()

Stop-Transcript
SendMail