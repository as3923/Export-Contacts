<#
===========================================================================
 AUTHOR  : Andrew Shen  
 DATE    : 2018-08-03
 VERSION : 1.1
===========================================================================

.SYNOPSIS
    Exports contacts as a PST file for a mailbox
    
.DESCRIPTION
    Exports contacts as a PST file for a mailbox. It is also possible to
    pipe a list of mailboxes to the function. Defaults for the parameters 
    can be set within the function, or they can be changed for each 
    instance. Currently this has only been tested on Microsoft Exchange 2013.
    
.PARAMETER Mailbox
    Specify the Mailbox to export Contacts from

.PARAMETER Server
    Specify the Exchange server to connect to. It is recommended to set 
    a default in the function below.

.PARAMETER ExportPath
    Specify the UNC path of the folder where PSTs will be exported to. It is 
    recommended to set a default in the function below.

.PARAMETER SimultaneousJobs
    Specify the number of simultaneous Export Requests. The default is 
    20 export requests.

.PARAMETER WaitTime
    Specify the number of seconds to wait before checking for 
    Completed/Failed Export Requests to queue up new requests. The default 
    is 15 seconds.

.PARAMETER ShowRuntime
    Specify whether the to display the total runtime in Hours at the end 
    of the script.

.EXAMPLE
    Get-Mailbox -Identity "ashen" | Export-ExchangeContacts

    # Exports the Contacts for "ashen" to a PST in the folder the function is run from
    
.EXAMPLE
    Get-Mailbox -ResultSize 40 | Export-ExchangeContacts

    # Exports Contacts for the first 40 Mailboxes returned by Get-Mailbox
#>

function Export-ExchangeContacts {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline,
            Mandatory=$true,
            HelpMessage="Enter mailboxes")]
        [Object] $Mailbox,
        [Parameter(HelpMessage="Enter FQDN of the Exchange Server to connect to")]
        [String] $Server = "defaultExchangeServer",
        [Parameter(HelpMessage="Enter location to save PSTs")]
        [String] $ExportPath = (Convert-Path .),
        [Parameter(HelpMessage="Enter number of exports to run simultaneously")]
        [Int] $SimultaneousJobs = 20,
        [Parameter(HelpMessage="Enter number of seconds to wait before checking for completed requests")]
        [Int] $WaitTime = 15,
        [Parameter(HelpMessage="Display total script run time after completion")]
        [Switch] $ShowRuntime
    )
    BEGIN {
        ### Validate parameters ###
        if (!(Test-Path $ExportPath -PathType Container -ErrorAction SilentlyContinue)){
            Throw "$ExportPath is not a valid folder"
        } else {
            Write-Verbose "PSTs will be exported to $ExportPath..."
        }

        ### Connect to Exchange ###
        $StopWatch = [system.diagnostics.stopwatch]::StartNew()
        Write-Verbose "Connecting to $Server..."
        $Session = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://$Server/powershell -Authentication Kerberos
        Import-PSSession -Session $Session -AllowClobber

        ### Set the inputs and outputs for exported data ###
        if ($ExportPath[-1] -ne "\"){$ExportPath += "\"}
        $timestamp = (Get-Date).tostring('yyyyMMddHHmmss')
        $batch = "Contacts_Export_$timestamp"

        Write-Verbose "$SimultaneousJobs Exports will be run concurrently..."
        Write-Verbose "Contacts for $($Mailbox.count) mailboxes will be exported..."
    }
    PROCESS {
        Try {
            ### Loop through each mailbox to begin export ###
            foreach ($m in $Mailbox){
                $PSTfile = $ExportPath + $m.Alias + ".pst"
	            $exportName = $m.Alias + "_" + $batch
	            New-MailboxExportRequest -Mailbox $m.Alias -IncludeFolders "#Contacts#" -BatchName $batch -Name $exportName -FilePath $PSTfile -ExcludeDumpster | Out-Null
                Write-Verbose "$exportName started..."

                ### Check number of Requests running and sleep if there are 20 or more ###
                do {
                    $runCount = @(Get-MailboxExportRequest | where {$_.BatchName -eq $batch -and $_.status -ne "Completed" -and $_.status -ne "Failed"}).count
                    Write-Verbose "$runCount Exports are InProgress..."
                    if ($runCount -ge $simultaneousJobs) {
                        Write-Verbose "Waiting $WaitTime seconds..."
                        Start-Sleep -Seconds $WaitTime
                    }
                } while ($runCount -ge $simultaneousJobs)
            }    
        } Catch {
            Write-Error $_
        }
    }
    END {
        ### Wait for all Requests to be Completed/Failed ###
        do {
            $runCount = @(Get-MailboxExportRequest | where {$_.BatchName -eq $batch -and $_.status -ne "Completed" -and $_.status -ne "Failed"}).count
            Write-Verbose "$runCount Exports are InProgress..."
            if ($runCount -gt 0) {
                Write-Verbose "Waiting 60 seconds for all Exports to Complete..."
                Start-Sleep -Seconds 60
            }
        } while ($runCount -gt 0)

        ### Export and remove completed results ###
        $LogPath = $ExportPath + "Logs\"
        if (!(Test-Path $LogPath -PathType Container -ErrorAction SilentlyContinue)){
            Write-Verbose "$LogPath does not exist"
            Write-Verbose "Creating $LogPath folder"
            New-Item -Path $LogPath -ItemType Directory
        }
        $resultCsv = $LogPath + "$batch.csv"
        $currentExportBatch = Get-MailboxExportRequest | where {$_.BatchName -eq $batch -and ($_.status -eq "Completed" -or $_.status -eq "Failed")}
        Write-Verbose "Saving export results to $resultCsv..."
        $currentExportBatch | Export-Csv -Path $resultCsv -NoClobber -NoTypeInformation
        Write-Verbose "Removing Completed/Failed Exports from the export batch $batch..."
        $currentExportBatch | Remove-MailboxExportRequest -Confirm:$false
        Write-Verbose "Contacts Export COMPLETE"

        ### Disconnect PSSession ###
        Remove-PSSession -Session $Session
        Write-Verbose "Disconnected PSSession from $Server..."

        ### Display run time/duration ###
        $StopWatch.Stop()
        if ($showRuntime) {Write-Output "Total script run time: $($StopWatch.Elapsed.TotalHours) Hours"}
    }
}

### Run function ###

#$userMailboxes = Get-Mailbox -ResultSize 40
#Export-PWContacts -Mailbox $userMailboxes -Verbose -ShowRunTime