function Manage-Service {
    <#
    .SYNOPSIS
      Start/Stop services on local or remote computers

    .DESCRIPTION
      Send 'Start', 'Stop', 'Restart', 'Pause' or 'Resume' actions to a service on the local or a
      remote computer

    .PARAMETER ServiceName (Mandatory)
      DisplayName OR Name of the service

    .PARAMETER Action (Mandatory)
      What action to perform on the service
      Only allow 'Start', 'Stop', 'Restart', 'Pause' and 'Resume'

    .PARAMETER ComputerName
      Set to remote computername, otherwise it's set to the local computer

    .PARAMETER Timeout
      Number of seconds to wait for a service action to complete (default 5 seconds if no value
      passed in)

    .PARAMETER ForceKill
      If a 'stop' action doesn't complete within the timeout timespan, then attempt to kill the
      process instead

    .EXAMPLE
        PS C:\> . .\Manage-Service.ps1
        PS C:\> Manage-Service -ServiceName Telephony -Action Restart

        Description:
        dot source the Manage-Service.ps1 script to expose access to the Manage-Service function
        Will stop the Telephony service, wait until complete (default timeout 5 seconds), then Start
        the Telephony service, on the local computer

    .EXAMPLE
        PS C:\> . .\Manage-Service.ps1
        PS C:\> Manage-Service -ServiceName Telephony -Action Stop -ComputerName SRV-APPSERVER

        Description:
        dot source the Manage-Service.ps1 script to expose access to the Manage-Service function
        Will Stop the Telephony service, wait until complete (default timeout 5 seconds), on the
        remote computer SRV-APPSERVER

    .EXAMPLE
        PS C:\> . .\Manage-Service.ps1
        PS C:\> Manage-Service -ServiceName Telephony -Action Restart -Timeout 8 -ForceKill -ComputerName SRV-APPSERVER

        Description:
        dot source the Manage-Service.ps1 script to expose access to the Manage-Service function
        Will Stop the Telephony service, wait until complete (manual timeout 8 seconds), if not
        stopped after 8 seconds then the process id is obtained and the processed killed, then start
        the Telephony service, on the remote computer SRV-APPSERVER

    .OUTPUTS
        To session and to transcript file.
        * transcript files older than 30 days are purged     

    .NOTES
        Author: Steve Geall
        Date: December 2018
        Version: 1.1

        History: v1.1 - 19/12/2018 - Removed unnecessary write-host and unnecessary commented lines

        Based on Stop, Start, Restart Windows Services – PowerShell Script
        by Khoa Nguyen on October 4, 2017
        https://www.syspanda.com/index.php/2017/10/04/stop-start-restart-windows-services-powershell-script/
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Start','Stop','Restart','Pause','Resume')]
        [string]$Action,

        [string]$ComputerName=$env:COMPUTERNAME,
        [switch]$ForceKill,
        [int]$Timeout=5
    )

    #region Prepare
    # start transcript (stop first in case it's already running in this session)
    try {
        Stop-Transcript
    } catch {
    }
    Start-Transcript "$PSScriptroot\transcripts\Transcript_$ComputerName`_$ServiceName`_$(Get-Random).log" -Force

    # variables
    $TStimeout = New-TimeSpan -Seconds $Timeout
    $ComputerName = ([System.Net.Dns]::GetHostByName($ComputerName)).HostName
    $functionName = $PSCmdlet.MyInvocation.InvocationName
    Write-Output "FUNCTION BEGIN: $functionName..."
    $parameterList = (Get-Command -Name $functionName).Parameters
    ForEach ($parameter in $parameterList) {
        Get-Variable -Name $parameter.Values.Name -ErrorAction SilentlyContinue | ft
    }
    Write-Output `r`n
    #endregion Prepare

    #region ServiceActionFunctions
    # stop
    function StopService {
        param (
            [object[]]$Service,
            [string]$ComputerName,
            $timeout
        )
        Try {
            Write-Output "Stopping service $($Service.DisplayName)..."
            Stop-Service -InputObject $Service -Force
            $Service.WaitForStatus('Stopped',$timeout)
        } Catch {
            Write-Output "ERROR: $($PSItem.Exception.Message)"
        }
        If ($Service.Status -ne 'Stopped') {
            If ($ForceKill) {
                Write-Output "Service hasn't stopped after the timeout period of $timeout, -ForceKill switch passed in, so let's just kill it..."
                $processId = Get-WmiObject win32_service -ComputerName $ComputerName | Where { $_.Name -like "$($Service.Name)" } | Select -ExpandProperty ProcessId
                If ($processId -gt 0) {
                    Write-Output "Stopping process ($($Service.Name)) with Id: $processId"
                    # Stop-Process doesn't have a -ComputerName parameter option so use Get-Process and pipe it through
                    Get-Process -Id $processId -ComputerName $ComputerName | Stop-Process -Force -Verbose
                    $stopCount = 0
                    Do {
                        Sleep -Seconds 1
                        Write-Output "waiting for process with Id: $processId to stop..."
                        $processExists = Get-Process -Id $processId -ErrorAction SilentlyContinue
                        $stopCount++
                    } Until (!$processExists -Or $stopCount -gt 60)
                } Else {
                    Write-Output "processId is: $processId, nothing to kill."
                }
            } Else {
                Write-Output "Service hasn't stopped after the timeout period of $timeout, but -ForceKill switch not passed in, so just leave service in this state."
            }
        }
        $Service = Test-Service -ServiceName $Service.DisplayName -ComputerName $ComputerName
        Write-Output "Current status: Service $($Service.DisplayName) is [$($Service.Status)]."
    }

    # start
    function StartService {
        param (
            [object[]]$Service,
            [string]$ComputerName,
            $timeout
        )
        Try {
            Write-Output "Stating service $($Service.DisplayName)..."
            Start-Service -InputObject $Service -ErrorAction Stop
            $Service.WaitForStatus('Running',$timeout)
        } Catch {
            Write-Output "error: $($PSItem.Exception.Message)"
        }
        $Service = Test-Service -ServiceName $Service.DisplayName -ComputerName $ComputerName
        Write-Output "Current status: Service $($Service.DisplayName) is [$($Service.Status)]."
    }

    # pause
    function SuspendService {
        param (
            [object[]]$Service,
            [string]$ComputerName,
            $timeout
        )
        Try {
            Write-Output "Suspending service $($Service.DisplayName)..."
            Suspend-Service -InputObject $Service -ErrorAction Stop
            $Service.WaitForStatus('Paused',$timeout)
        } Catch {
            Write-Output "error: $($PSItem.Exception.Message)"
        }
        $Service = Test-Service -ServiceName $Service.DisplayName -ComputerName $ComputerName
        Write-Output "Current status: Service $($Service.DisplayName) is [$($Service.Status)]."
    }

    # resume
    function ResumeService {
        param (
            [object[]]$Service,
            [string]$ComputerName,
            $timeout
        )
        Try {
            Write-Output "Resuming service $($Service.DisplayName)..."
            Resume-Service -InputObject $Service -ErrorAction Stop
            $Service.WaitForStatus('Running',$timeout)
        } Catch {
            Write-Output "error: $($PSItem.Exception.Message)"
        }
        $Service = Test-Service -ServiceName $Service.DisplayName -ComputerName $ComputerName
        Write-Output "Current status: Service $($Service.DisplayName) is [$($Service.Status)]."
    }
    #endregion ServiceActionFunctions

    #region MainScript
    # check if service exists
    $Service = Test-Service -ServiceName $ServiceName -ComputerName $ComputerName

    If ($Service) {
        Switch ($Action) {
            # condition if user wants to stop a service
            'Stop' {
                If ($Service.Status -eq 'Stopped') {
                    Write-Output "Service: $ServiceName is already stopped! [$($Service.Status)]"
                } ElseIf ($Service.Status -eq 'Paused') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to stop (resume then stop)..."
                    ResumeService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                    StopService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } ElseIf ($Service.Status -eq 'Running') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to stop..."
                    StopService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } Else {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], nothing to do? Problem? Email admins?"
                }
            }
            # condition if user wants to start a service
            'Start' {
                If ($Service.Status -eq 'Running') {
                    Write-Output "Service: $ServiceName is already running! [$($Service.Status)]"
                } ElseIf ($Service.Status -eq 'Paused') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to resume..."
                    ResumeService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } ElseIf ($Service.Status -eq 'Stopped') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to start..."
                    StartService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } Else {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], nothing to do? Problem? Email admins?"
                }
            }
            # condition if user wants to restart a service
            'Restart' {
                If ($Service.Status -eq 'Running') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to restart (stop service, then start it)..."
                    StopService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                    StartService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } ElseIf ($Service.Status -eq 'Paused') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to restart (resume service, stop it, then start it)..."
                    ResumeService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                    StopService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                    StartService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } ElseIf ($Service.Status -eq 'Stopped') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to start..."
                    StartService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } Else {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], nothing to do? Problem? Email admins?"
                }
            }
            # condition if user wants to pause a service
            'Pause' {
                If ($Service.Status -eq 'Paused' -Or $Service.Status -eq 'Stopped') {
                    Write-Output "Service: $ServiceName is already paused (or stopped)! [$($Service.Status)]"
                } ElseIf ($Service.Status -eq 'Running') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to pause..."
                    SuspendService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } Else {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], nothing to do? Problem? Email admins?"
                }
            }
            # condition if user wants to resume a service
            'Resume' {
                If ($Service.Status -eq 'Running') {
                    Write-Output "Service: $ServiceName is already running! [$($Service.Status)]"
                } ElseIf ($Service.Status -eq 'Paused') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to resume..."
                    ResumeService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } ElseIf ($Service.Status -eq 'Stopped') {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], preparing to start..."
                    StartService -Service $Service -ComputerName $ComputerName -timeout $TStimeout
                } Else {
                    Write-Output "Service: $ServiceName has a status of [$($Service.Status)], nothing to do? Problem? Email admins?"
                }
            }
            # condition if action is anything other than stop, start, restart
            default {
                Write-Output "Action parameter is missing or invalid!"
            }
        }

    } Else {
        # condition if provided ServiceName is invalid
        Write-Output "Service: $ServiceName not found"
    }

    # output final status of service
    If ($Service) {
        $finalStatus = Test-Service -ServiceName $ServiceName -ComputerName $ComputerName | Select MachineName,DisplayName,Name,Status
        Write-Output $finalStatus | Format-Table
    } Else {
        Return "Service: $ServiceName not found"
    }
    #endregion MainScript

    #region Cleanup
    # purge transcripts
    $dayLimit = (Get-Date).AddDays(-30)
    $transcriptFiles = "$PSScriptroot\transcripts\*transcript*.log"
    # delete files older than the dayLimit
    $transcriptFilesGet = Get-ChildItem -Path $transcriptFiles -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $dayLimit }
    if ($transcriptFilesGet) {
        Write-Output "`r`nPurging transcript files: $transcriptFilesGet"
        Get-ChildItem -Path $transcriptFiles -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $dayLimit } | Remove-Item -Force
    }

    # stop transcript
    try {
        Stop-Transcript
    } catch {
    }
    #endregion Cleanup

} #function Manage-Service


function Test-Service {
    <#
    .SYNOPSIS
    Solely just for checking a service and returning its current status/state

    .DESCRIPTION
    Sets a var to null, then attempts to assign service status to it and returns that var
    Allows external helper scripts to use this function (when this file is dot sourced)

    .PARAMETER ServiceName (Mandatory)
    The DisplayName OR Name of the service

    .PARAMETER ComputerName
    The name of the remote computer, otherwise it's set to the local computer
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,
        [string]$ComputerName=$env:COMPUTERNAME
    )
    $ServiceCheck = $null
    $ServiceCheck = Get-Service $ServiceName -ComputerName $ComputerName -ErrorAction SilentlyContinue
    Return $ServiceCheck
} #function Test-Service