<#
.SYNOPSIS

    This script will be placed on each MDT server.
    This will import tasch scheuler and it will trigger it on remote server (Task name - "MDT Maintenance" and Script name - "MDT_Replace-WIM.bat")
    Note: We tried to run the MDT commands via Invoke-command & PSEXEC but it failed due to limitations in MDT ( integrated with AD )
    
 .Author

    maniboopathii@gmail.com

 .Created Date
    
    2017-07-25

 .Modified Date

    2017-07-28

 .version
     
     1.0 - Script created

 .Usage
    !!!!!!!!!!!!!! In Batch file you have mention the "Target" value based on requirement!!!!!!!!!!!!!

    Specific - This will not export the list from DFS and it will not upate "LimitedListComputers.txt". This will run against the list that you have updated
    All - This will export the list from DFS and it will update the list of MDT members and this will run against the list of servers but it will exclude regional servers.
     
#>

param
(
    [parameter(Mandatory=$true)] `
    [string] $Target
)

###############
#Setting variables for log and CSV
###############
$TranscriptLog = "Your Folder\Transcript_MDT_Replace-WIM.log"
$DFSXML= "Your Folder\MDT_Replace-WIM\MDTTargets.xml"
$Complist= "Your Folder\MDT_Replace-WIM\LimitedListComputers.txt"

$CSVFolder = 'Your Folder\MDT_ReplaceWIM\CSVes'
$CSVFolderBackup = 'Your Folder\MDT_ReplaceWIM\CSVes\Backup'
$TrasncriptFolder = 'Your Folder\MDT_ReplaceWIM\Transcripts'
$TrasncriptFolderBackup = 'Your Folder\MDT_ReplaceWIM\Transcripts\Backup'
$OutputFile = 'Your Folder\MDT_ReplaceWIM\' + $(get-date).ToString("yyyyMMdd") + '_Final_MDT_Replace-WIM.csv'

Move-Item ($CSVFolder+"\*.csv") $CSVFolderBackup -Force
Move-Item ($TrasncriptFolder+"\*.txt") $TrasncriptFolderBackup -Force

Start-Transcript $TranscriptLog

###############
# Checking sechedule task name "MDT Maintenance". 
# if it is there and task will be triggered, 
# if not it will import them and task will be triggered
###############

Function Import-ScheduledTask
    {
       foreach ($ServerList in $ServerList)
            {
                $CheckScheduledTask = Invoke-Command -ComputerName $ServerList -ScriptBlock {$(Get-ScheduledTask).TaskName | where {$_ -like "MDT Maintenance"}}
                If($CheckScheduledTask -like "MDT Maintenance")
                    {       
                                 
                        Write-Output "$(Get-date) | Initiating Scheduled Task on $ServerList"
                        Invoke-Command -ComputerName $ServerList -ScriptBlock {Start-ScheduledTask -TaskName "MDT Maintenance"}
                    }
                else
                    {
                        Write-Output "$(Get-date) | Importing Scheduled Task on $ServerList"
                        Invoke-Command -ComputerName $ServerList -ScriptBlock {Register-ScheduledTask -Xml (Get-Content -Path "M:\MDTProduction\MDTInstallationFiles\MDT_Replace-WIM\MDT Maintenance.xml" | Out-String) -TaskName "MDT Maintenance"}
                        Write-Output "$(Get-date) | Initiating Scheduled Task on $ServerList"
                        Invoke-Command -ComputerName $ServerList -ScriptBlock {Start-ScheduledTask -TaskName "MDT Maintenance"}
                    }
            }        
    }

##############
# Reading Target value and it will run accordingly
##############

if($Target -like "All")
    {
        dfsutil /root:\\ad.Organization.com\ /export:$DFSXML

        #Creating computer list from the XML file

        Write-Output "$(Get-Date) | Collecting MDT Targets computername from XML as per the '-Computername' and '-excludecomputername' parameter"
        [xml] $xmlPath= Get-Content "$DFSXML"
        $link = $xmlPath.GetElementsByTagName("Link")
        $link.target | Where-Object {$_.folder -eq "MDTproduction$" -and $_.server -like "*ad.Organization.com" -and $_.server -notlike "RegionalServer1" `
                                                                                                               -and $_.server -notlike "RegionalServer2" `
                                                                                                               -and $_.server -notlike "RegionalServer3"} | Select-Object -ExpandProperty Server `
                                                                                                                                                                             | Out-File -FilePath $Complist -Encoding utf8
        $ServerList = Get-Content $Complist
                                                                                    
        Import-ScheduledTask                                                                                                    
    }
elseif ($Target -like "Specific")
    {
        Write-Output "$(Get-Date) | This script will run against on specific servers"
        $ServerList = Get-Content $Complist
        Import-ScheduledTask
        #########
        # Pause the script if the Target value is "Specific" 
        # to allow the servers to do their job
        #########

        sleep -Seconds 120
    }

###############
# Reading all CSVes and it will parse
###############

$CSV= @()

Get-ChildItem -Path $CSVFolder -Filter *.csv | ForEach-Object { 
    $CSV += @(Import-Csv -Path $_.FullName)
}

$CSV | Export-Csv -Path $OutputFile -NoTypeInformation -Force

    
Stop-Transcript



