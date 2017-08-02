1. This whole automation to design to replace WIM in remote servers which are configured with WDS with AD integrated.
2. All files have been placed in remote server
3. Update "LimitedListComputers.txt" if you are target is single client or let it export from DFS group and update this list by the script
4. "MDT Maintenance.xml" is being configured to run with "Sytem" account as the AD account or member (Ad account) of local Administrator.
5. "MDT_Replace-WIM.bat" will be triggered by Scheduled Task "MDT Maintenance"
6. You have to trigger any of the script based on your need "MDT_Trigger-ScheduledTask - All.bat" or "MDT_Trigger-ScheduledTask - Selected Servers.bat".
7. The trigger script (PS1) will be executed by above script (Point 6) and this will do the magic 


Magic????
1. That will check scheduled task "MDT Maintenance" if not it will import the xml.
2. and the scheduled task will be started with system account
3. That will allow sometime to execute the script "MDT_Replace-WIM.ps1" on remote server and it will copy CSV file and transcript log to central location
4. Finally "MDT_Trigger-ScheduledTask.ps1" will parse all CSV's which are available and it will be merged into single CSV (Cumulative report)
