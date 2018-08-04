<#
===========================================================================
 AUTHOR  : Andrew Shen  
 DATE    : 2018-08-03
 VERSION : 1.1
===========================================================================

.DESCRIPTION
    Exports contacts as a PST file for a list of mailboxes

#>

function Export-Contacts {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline,
            Mandatory=$true,
            HelpMessage="Enter list of mailboxes")]
        [String] $Mailboxes,
        [Parameter(HelpMessage="Enter the Exchange Server to connect to")]
        [String] $Server = "defaultExchangeServer",
        [Parameter(HelpMessage="Enter location to save PSTs")]
        [String] $ExportPath = "\\SharedFolder\BackupOfEveryonesContacts\",
        [Parameter(HelpMessage="Enter number of exports to run simultaneously")]
        [Int] $SimultaneousJobs = 20,
        [Parameter(HelpMessage="Display total script run time after completion")]
        [Bool] $ShowRunTime = $false
    )
    BEGIN {
        ### Validate parameters ###
        if (Test-Path $ExportPath -PathType Container -ErrorAction SilentlyContinue){
            Throw "$ExportPath is not a valid folder"
        }

        ### Connect to Exchange ###
        $StopWatch = [system.diagnostics.stopwatch]::StartNew()
        Write-Verbose "Connecting to $Server..."
        $Session = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://$Server/powershell -Authentication Kerberos
        Import-PSSession -Session $Session -AllowClobber

        ### Set the inputs and outputs for exported data ###
        if ($ExportPath[-1] -ne "\"){$ExportPath += "\"}
        $timestamp = (Get-Date).tostring('yyyyMMddHHmmss')
        $batch = "Export_$timestamp"

        Write-Verbose "PSTs will be exported to $ExportPath..."
        Write-Verbose "$SimultaneousJobs Exports will be run concurrently..."
        Write-Verbose "Contacts for $($mailboxes.count) mailboxes will be exported..."
        Write-Verbose ""
    }
    PROCESS {
        Try {
            ### Loop through each mailbox to begin export ###
            foreach ($m in $Mailboxes){
                $PSTfile = $ExportPath + $m.Alias + ".pst"
	            $exportName = $m.Alias + " Contacts Export_" + $timestamp
	            New-MailboxExportRequest -Mailbox $m.Alias -IncludeFolders "#Contacts#" -BatchName $batch -Name $exportName -FilePath $PSTfile -ExcludeDumpster | Out-Null
                Write-Verbose "$exportName started..."

                ### Check number of Requests running and sleep if there are 20 or more ###
                do {
                    $runCount = (Get-MailboxExportRequest | where {$_.BatchName -eq $batch -and $_.status -ne "Completed" -and $_.status -ne "Failed"}).count
                    Write-Verbose "$runCount Exports are InProgress..."
                    if ($runCount -ge $simultaneousJobs) {
                        Write-Verbose "Waiting 15 seconds..."
                        Start-Sleep -Seconds 15
                    }
                    Write-Verbose ""        
                } while ($runCount -ge $simultaneousJobs)
            }    
        } Catch {
            Write-Error $_
        }
    }
    END {
        ### Wait for all Requests to be Completed/Failed ###
        do {
            $runCount = (Get-MailboxExportRequest | where {$_.BatchName -eq $batch -and $_.status -ne "Completed" -and $_.status -ne "Failed"}).count
            Write-Verbose "$runCount Exports are InProgress..."
            if ($runCount -gt 0) {
                Write-Verbose "Waiting 60 seconds for all Exports to Complete..."
                Start-Sleep -Seconds 60
            }
        } while ($runCount -gt 0)

        ### Export and remove completed results ###
        $resultCsv = $ExportPath + "Logs\$batch.csv"
        $currentExportBatch = Get-MailboxExportRequest | where {$_.BatchName -eq $batch -and ($_.status -eq "Completed" -or $_.status -eq "Failed")}
        Write-Verbose "Saving export results to $resultCsv..."
        $currentExportBatch | Export-Csv -Path $resultCsv -NoClobber -NoTypeInformation
        Write-Verbose "Removing Completed/Failed Exports from the export batch $batch..."
        $currentExportBatch | Remove-MailboxExportRequest -Confirm:$false
        Write-Verbose "Contacts Export COMPLETE"

        ### Disconnect PSSession ###
        Exit-PSSession
        Write-Verbose "Disconnected PSSession from $Server..."

        ### Display run time/duration ###
        $StopWatch.Stop()
        if ($showRunTime -eq $true) {Write-Output "Total script run time: $($StopWatch.Elapsed.TotalDays) Days"}
    }
}

### Run function ###

#$userMailboxes = Get-Mailbox -ResultSize 40
#Export-PWContacts -Mailboxes $userMailboxes -Verbose -ShowRunTime