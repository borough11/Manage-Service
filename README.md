# Manage-Service
Send 'Start', 'Stop', 'Restart', 'Pause' or 'Resume' actions to a service(s) on the local or remote computer(s)

## Why do we need this?
I had a software vendor supply us with a method of restarting the services for their application each night via a batch file that just:

1. Performed a TASKKILL on 27 processes across 6 remote machines
2. Called an external SLEEP.EXE file for a dumb wait period of 240s
3. Then finally called another batch file that ran an SC START over the 27 services

This would hopefully result in restarted services. It worked … *sometimes*. When it didn’t work, we were unsure why because there was no logging.

This isn't really the greatest solution; so I set about replacing this with a PowerShell script to manage the services, not just for this particular piece of software but to make it generic so it can be used for any service(s)/computer(s). It allows you to send 'Start', 'Stop', 'Restart', 'Pause' or 'Resume' actions to a service on the local or remote computer. I included complete logging, error handling and an optional forcekill option if the service doesn't stop nicely after x seconds.

Initially the best script I came across on the net was Khoa Nguyen's "Stop, Start, Restart Windows Service" so I based my script on this: https://www.syspanda.com/index.php/2017/10/04/stop-start-restart-windows-services-powershell-script/ by Khoa Nguyen on October 4, 2017

## Usage - *From command line...*
dot source the Manage-Service.ps1 script to expose access to the Manage-Service function

`PS C:\> . .\Manage-Service.ps1`

Example 1:

Will stop the Telephony service, wait until complete (default timeout 5 seconds), then Start the Telephony service, on the local computer.

`PS C:\> Manage-Service -ServiceName Telephony -Action Restart`
 
Example 2:

Will Stop the Telephony service, wait until complete (default timeout 5 seconds), on the remote computer SRV-APPSERVER

`PS C:\> Manage-Service -ServiceName Telephony -Action Stop -ComputerName SRV-APPSERVER`

Example 3:

Will Stop the Telephony service, wait until complete (manual timeout 30 seconds), if not stopped after 30 seconds then the process id is obtained and the processed killed, then an attempt to start the Telephony service on the remote computer SRV-APPSERVER.

`PS C:\> Manage-Service -ServiceName Telephony -Action Restart -Timeout 30 -ForceKill -ComputerName SRV-APPSERVER`



## Usage - *From a helper script...*

Create a PowerShell script specific to your requirements, perhaps building a hash table of computers you want to query including an array of Service names on each computer, similar to:
```
$AppServersAndServices = @{}

$AppServersAndServices."srv-app01" = @()
$AppServersAndServices."srv-app02" = @()

$AppServersAndServices."srv-app01" += "GenericServiceOne"
$AppServersAndServices."srv-app01" += "GenericServiceTwo"
$AppServersAndServices."srv-app01" += "SpecificServiceToApp01"

$AppServersAndServices."srv-app02" += "GenericServiceOne"
$AppServersAndServices."srv-app02" += "GenericServiceTwo"
```

Then if wanting to `Restart` each service, loop through this hash table and pass each computerName and ServiceName to the function `Manage-Service` with the `-Action` parameter set to `Restart`:
```
#requires -Version 3.0

# dot source the Manage-Service script to expose access to its functions
. "$PSScriptroot\Manage-Service.ps1"

# get current session user context
$currentSessionUserContext = ([Security.Principal.WindowsIdentity]::GetCurrent()).Name

Write-Host "`r`nStart..." -ForegroundColor Magenta
$AppServersAndServices
$beginDate = Get-Date

$finalStatuses = @()
$AppServersAndServices.GetEnumerator() | ForEach-Object {
    $ComputerName = $_.Name
    $_.Value | ForEach-Object {
        $ServiceName = $_
        Write-Host "`r`nPassing this to Manage-Service function, ServiceName: $ServiceName ComputerName: $ComputerName`..." -ForegroundColor Cyan
        Manage-Service -ServiceName $ServiceName -Action Restart -Timeout 30 -ForceKill -ComputerName $ComputerName
        $finalStatusCheck = $null
        $finalStatusCheck = Test-Service -ServiceName $ServiceName -ComputerName $ComputerName | Select MachineName,DisplayName,Name,Status
        If ($finalStatusCheck) {
            Write-Host "Maintain-Service actioned OK." -ForegroundColor Green
            $finalStatuses += $finalStatusCheck
        } Else {
            Write-Host "Manage-Service action FAILED! (for ServiceName ""$ServiceName"" on ComputerName ""$ComputerName"")" -ForegroundColor Yellow
            $customPSStatusObj = New-Object -TypeName PSobject
            $customPSStatusObj | Add-Member -MemberType NoteProperty -Name MachineName -Value $ComputerName
            $customPSStatusObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $ServiceName
            $customPSStatusObj | Add-Member -MemberType NoteProperty -Name Name -Value $ServiceName
            $customPSStatusObj | Add-Member -MemberType NoteProperty -Name Status -Value "UNKNOWN"
            $finalStatuses += $customPSStatusObj
        }
    }
}
```

Finally, perhaps add this information to a rolling daily log file specific to your computers and services:

```
Write-Host "`r`nFinal Statuses..." -ForegroundColor Magenta
$finalStatuses | Format-Table
Write-Host "servicesToRestart=$($finalStatuses.count)"
Write-Host "servicesRestartedOK=$(($finalStatuses | Where {$_.Status -eq "Running"}).Count)"
Write-Host "servicesNotRunning=$(($finalStatuses | Where {$_.Status -ne "Running"}).Count)"

# add final statuses to daily log file
$logFolder = "$PSScriptroot"
$logName = "$($MyInvocation.MyCommand)_log"
$logFile = Join-Path -Path "$logFolder" -ChildPath ("$logName-{0:ddd}.txt" -f (Get-Date))
# first remove previous daily log file (named by day) if it hasn't been written to today
If (Test-Path -Path $logFile) {
    If ((Get-Item $logFile).LastWriteTime -lt (Get-Date).AddDays(-1)) {
        Write-Host " ...today's log file is from last week, delete it before writing today's new log info to $logFile"
        Remove-Item -Path $logFile -Force
    } Else {
        Write-Host " ...today's log file was written to today, so keep adding to $logFile"
    }
}
$logContent = @"
`r`nBegin: $beginDate
Run on computer: $(([System.Net.Dns]::GetHostByName($env:COMPUTERNAME)).HostName)
Run as user: $currentSessionUserContext
Restarting (stopping then starting) services on below computers...
$(ForEach ($finalStatus in ($finalStatuses | Group-Object -Property MachineName | Sort-Object -Property Name)) {
"`r`n`t- $(($finalStatus.Name).ToUpper())`r`n"
ForEach ($svc in $finalStatus.Group) {
If ($svc.Status -ne "Running"){
"`t  ! [$($svc.Status)] - $($svc.DisplayName) ($($svc.Name))"
} Else {
"`t  * [$($svc.Status)] - $($svc.DisplayName) ($($svc.Name))"
}
"`r`n"
}})
Total services attempted to restart: $($finalStatuses.count)
Services restarted OK (in Running state): $(($finalStatuses | Where {$_.Status -eq "Running"}).Count)
Services failed to restart (access? name?): $(($finalStatuses | Where {$_.Status -ne "Running"}).Count)
End: $(Get-Date)
"@
Add-Content -Path $logFile -Value $logContent
```
