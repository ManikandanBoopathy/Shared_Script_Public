<#
.SYNOPSIS

    This script will be placed on each MDT server.
    This will be executed by task scheduler which will be triggered by MDT Administrator with the script "MDT_Trigger-ScheduledTask.bat"
    Note: We tried to run the MDT commands via Invoke-command & PSEXEC but it failed due to limitations in MDT ( integrated with AD )
    
 .Author

    maniboopathii@gmail.com

 .Created Date
    
    2017-07-25

 .Modified Date

    2017-07-28

 .version
     
     1.0 - Script created
     
#>

$output = @()

###############
#Setting variables for log and CSV
###############

$TranscriptLog = "M:\Temp\" + "$env:COMPUTERNAME" + "_Transcript_MDT_Replace-WIM.txt"
$Finalreport = "M:\Temp\" + "$env:COMPUTERNAME" + "_Replace-WIM.csv"

###############
# Checking the Temp folder is exist or not. if not and it will create Temp folder in M:\
###############

Start-Transcript $TranscriptLog

if((Test-Path M:\Temp -ErrorAction SilentlyContinue ) -eq $true)
    {
        #Dummy
        Write-Output "$env:COMPUTERNAME | $(Get-Date) | Temp folder already exist"
    }
else
    {
        New-Item -Path M:\ -Name Temp -ItemType Directory
        Write-Output "$env:COMPUTERNAME | $(Get-Date) | Temp folder folder created"
    }
    
#################
#Reading old Boot.wim Name and Filename  
#################

$OldBootx64Name = get-wdsbootimage -architecture x64 | where {$_.Name -like "BootWIMName*"} | select -ExpandProperty Name

$OldBootx64Filename = get-wdsbootimage -architecture x64 | where {$_.filename -like "LiteTouchPE_x64.wim"} | Select -ExpandProperty Filename

$OldBootx86Name = get-wdsbootimage -architecture x86 | where {$_.Name -like "BootWIMName*"} | select -ExpandProperty Name

$OldBootx86Filename = get-wdsbootimage -architecture x86 | where {$_.filename -like "LiteTouchPE_x86.wim"} | Select -ExpandProperty Filename

################
# writing the old Boot.wim file information to transcript log
################

Write-Output "$env:COMPUTERNAME | $(Get-Date) | Attempting to get all boot.wim file namese" 
Write-Output "*****************Old Boot.WIM files************"
$OldBootx64Name
$OldBootx86Name
Write-Output "************************************************"

$date=$(get-date).ToString("yyyyMMdd")

################
#Setting new boot file name and path to import
################

$Newx64BootName = "NewBootWIMName x64 $date"
$Newx86BootName = "NewBootWIMName x86 $date"
$Newx64BootPath = "YourFolder\LiteTouchPE_x64.wim"
$Newx86BootPath = "YourFolder\LiteTouchPE_x86.wim"

################
# Checking to remove and boot.wim files
###############

if(($OldBootx64Name -like "BootWIMName*") -and ($OldBootx64Filename -like "LiteTouchPE_x64.wim"))
    {
        ##############
        # Removing boot files
        ##############

        try
            { 
                Write-Output "$env:COMPUTERNAME | $(Get-Date) | Attempting to remove old boot file - $OldBootx64Name"
                Remove-WdsBootImage -Architecture X64 -ImageName $OldBootx64Name -FileName $OldBootx64Filename
            }
        Catch
            {
                Write-Output "$env:COMPUTERNAME | $(Get-Date) | Failed to remove old boot file - $OldBootx64Name"
                $Errorx64Remove = "Failed"
            }
        if($Errorx64Remove -eq "Failed")
            {
                #Dummy
            }
        else
            {
                ###########
                # Attempting to add new boot.wim file
                ###########

                Write-Output "$env:COMPUTERNAME | $(Get-Date) | Removed old boot file - $OldBootx64Name successfully"
                try
                    {
                        Write-Output "$env:COMPUTERNAME | $(Get-Date) | Attempting to add new boot file - $Newx64BootName"
                        Import-WdsBootImage -Path $Newx64BootPath -NewImageName $Newx64BootName
                    }
                catch
                    {
                        $Errorx64Add = "Failed"
                    }
                if($Errorx64Add -eq "Failed")
                    {
                        #Dummy
                    }
                else
                    {
                        Write-Output "$env:COMPUTERNAME | $(Get-Date) | New boot file - $Newx64BootName added successfully"
                        $Newbootx64Validated = get-wdsbootimage -architecture x64 | where {$_.Name -like "BootWIMName*"} | select -ExpandProperty Name
                    }
            }
    }
else
    {
            ###############
            # No attempt to remove old boot.wim files
            # and Adding new boot.wim file.
            ###############

           try
              {
                  Write-Output "$env:COMPUTERNAME | $(Get-Date) | There was no old boot.wim so adding new boot.wim"
                  Import-WdsBootImage -Path $Newx64BootPath -NewImageName $Newx64BootName
              }
          catch
              {
                 $Errorx64Add = "Failed"
              }
          if($Errorx64Add -eq "Failed")
              {
                        #Dummy
              }
          else
              {
                  Write-Output "$env:COMPUTERNAME | $(Get-Date) | New boot file - $Newx64BootName added successfully"
                  $Newbootx64Validated = get-wdsbootimage -architecture x64 | where {$_.Name -like "BootWIMName*"} | select -ExpandProperty Name
              }
    }


if(($OldBootx86Name -like "BootWIMName*") -and ($OldBootx86Filename -like "LiteTouchPE_x86.wim"))
    {
        ##############
        # Removing boot files
        ##############

        try
            {
                Write-Output "$env:COMPUTERNAME | $(Get-Date) | Attempting to remove old boot file - $OldBootx86Name"
                Remove-WdsBootImage -Architecture x86 -ImageName $OldBootx86Name -FileName $OldBootx86Filename
            }
        Catch
            {
                Write-Output "$env:COMPUTERNAME | $(Get-Date) | Failed to remove old boot file - $OldBootx86Name"
                $Errorx86Remove = "Failed"
            }
        if($Errorx86Remove -eq "Failed")
            {
                #Dummy
            }
        else
            {
                ###########
                # Attempting to add new boot.wim file
                ###########

                Write-Output "$env:COMPUTERNAME | $(Get-Date) | Removed old boot file - $OldBootx86Name successfully"
                try
                    {
                        Write-Output "$env:COMPUTERNAME | $(Get-Date) | Attempting to add new boot file - $Newx86BootName"
                        Import-WdsBootImage -Path $Newx86BootPath -NewImageName $Newx86BootName
                    }
                catch
                    {
                        $Errorx86Add = "Failed"
                    }
                if($Errorx86Add -eq "Failed")
                    {
                        #Dummy
                    }
                else
                    {
                        Write-Output "$env:COMPUTERNAME | $(Get-Date) | New boot file - $Newx86BootName added successfully"
                        $Newbootx86Validated = get-wdsbootimage -architecture x86 | where {$_.Name -like "BootWIMName*"} | select -ExpandProperty Name
                    }
            }
    }
else
    {
         ###############
         # No attempt to remove old boot.wim files and Adding new boot.wim file.
         ###############
        try
            {
                Write-Output "$env:COMPUTERNAME | $(Get-Date) | There was no old boot.wim so adding new boot.wim"
                Import-WdsBootImage -Path $Newx86BootPath -NewImageName $Newx86BootName
                
            }
        catch
            {
                $Errorx86Add = "Failed"
            }
        if($Errorx86Add -eq "Failed")
                    {
                        #Dummy
                    }
                else
                    {
                        Write-Output "$env:COMPUTERNAME | $(Get-Date) | New boot file - $Newx86BootName added successfully"
                        $Newbootx86Validated = get-wdsbootimage -architecture x86 | where {$_.Name -like "BootWIMName*"} | select -ExpandProperty Name
                    }
    }

##########
# Writing output into CSV
##########

 $output+=New-Object PSObject -Property @{"Server Name" = $env:COMPUTERNAME
                                                                 "Old Boot x86 Name" = $oldbootx86Name
                                                                 "New Boot x86 Name" = $Newbootx86Validated
                                                                 "Old Boot x64 Name" = $oldbootx64Name
                                                                 "New Boot x64 Name" = $Newbootx64Validated                   
                                                                     } | select "Server Name","Old Boot x86 Name","New Boot x86 Name","Old Boot x64 Name","New Boot x64 Name"

$output| Export-Csv $Finalreport -ErrorAction SilentlyContinue -NoTypeInformation

###########
# Copying the both CSV and 
# Trascnript log to central location
###########

Copy-Item $Finalreport '\\CentralServer\CentralLocation\CSVes'
Copy-Item $TranscriptLog '\\CentralServer\CentralLocation\Transcripts'
Stop-Transcript