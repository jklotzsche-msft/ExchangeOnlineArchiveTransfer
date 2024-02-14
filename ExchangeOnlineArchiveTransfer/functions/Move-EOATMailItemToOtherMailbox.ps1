function Move-EOATMailItemToOtherMailbox {
    <#
    .SYNOPSIS
    Move mail items to another mailbox
    
    .DESCRIPTION
    Move mail items to another mailbox
    This function moves mail items to another mailbox using the Exchange Web Services (EWS) API.
    
    .PARAMETER MailItems
    PSCustomObject array. Default value is $null.
    This parameter defines the mail items, which should be moved.
    You can use the function 'Get-EOATMailItem' to get a list of mail items.
    
    .PARAMETER TargetMailbox
    String value. Default value is $null.
    This parameter defines the target mailbox, to which the mail items should be moved.
    
    .PARAMETER TargetFolder
    String value. Default value is $null.
    This parameter defines the target folder, to which the mail items should be moved.
    
    .PARAMETER MaxBatchSize
    Integer value. Default value is 96636764160 (90 GB).
    This parameter defines the maximum batch size in Byte, which is used to retrieve the mail items.
    If more than 90 GB of mail items are send to this function, the mails will be split into batches minimum 90 GB.
    A single batch could be larger than 90GB, because we add items to the batch until AT LEAST 90GB are reached.

    .PARAMETER WaitTime
    Integer value. Default value is 300 (5 minutes).
    This parameter defines the time in seconds, which the function will wait, before continuing with the next batch.

    .PARAMETER CheckTargetFolderEmpty
    Switch value. Default value is $true.
    This parameter defines, if the function should check, if the target folder is empty, before continuing with the next batch.
    If the target folder is not empty, the function will wait $WaitTime seconds before continuing with the next batch.
    As soon as the target folder is empty, the function will ask the script user, if the function should continue with the next batch.

    .PARAMETER LogEnabled
    Switch value. Default value is $false.
    This parameter defines, if the function should log the mail item copy process to a CSV file.

    .PARAMETER LogFilePath
    String value. Default value is "$env:temp\Copy-EOATMailItemToOtherMailbox-$(Get-Date -Format yyyyMMddhhmmss).csv".
    This parameter defines the path to the log file, which is used to log the mail item copy process.
    The log file have the following header: SourceMailbox, SourceFolderId, TargetMailbox, TargetFolder, TargetFolderId, SourceMailItemId, Sender, Subject, Received, SizeInMB, CurrentWindowsUser

    .PARAMETER LogDelimiter
    String value. Default value is ';'.
    This parameter defines the delimiter, which is used to separate the values in the log file.

    .PARAMETER Service
    ExchangeService object. Default value is the script variable $script:EwsService.
    This parameter defines the ExchangeService object, which is used to connect to Exchange Online EWS.
    This parameter will be set automatically, if you use the function 'Connect-EOATExchangeWebService' to connect to Exchange Online EWS.
    
    .PARAMETER Confirm
    Switch value. Default value is $true.
    This parameter defines, if the function should ask the script user to continue with the next batch, if the target folder is empty.

    .PARAMETER WhatIf
    Switch value. Default value is $false.
    This parameter defines, if the function should simulate the execution of the function.
    This parameter is currently not implemented

    .EXAMPLE
    Move-EOATMailItemToOtherMailbox -MailItems $MailItems -TargetMailbox 'temporary_mailbox@contoso.com' -TargetFolder 'Inbox'
    
    This example moves the mail items defined in the variable $MailItems to the mailbox 'temporary_mailbox@contoso.com'. The mail items will be moved to the folder 'Inbox'.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]
        $MailItems,

        [Parameter(Mandatory = $true)]
        [String]
        $TargetMailbox,

        [ArgumentCompleter({
                [Microsoft.Exchange.Webservices.Data.WellKnownFolderName] | Get-Member -Static -MemberType Properties | Select-Object -ExpandProperty Name
            })]
        [Parameter(Mandatory = $true)]
        [String]
        $TargetFolder,

        [Int64]
        $MaxBatchSize = 90GB,

        [Int32]
        $WaitTime = 300,

        [bool]
        $CheckTargetFolderEmpty = $true,

        [switch]
        $LogEnabled,

        [string]
        $LogFilePath = "$env:temp\Move-EOATMailItemToOtherMailbox-$(Get-Date -Format yyyyMMddhhmmss).csv",

        [string]
        $LogDelimiter = ';',

        [Parameter(DontShow)]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]
        $Service = $script:EwsService
    )

    Begin {
        # trap statement
        $ErrorActionPreference = 'Stop'
        trap {
            Write-Error -Message $_.Exception.Message
            Write-Error -Message $_.Exception.StackTrace
            return
        }

        # Check if target folder exists
        try {
            $targetFolderId = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, [Microsoft.Exchange.WebServices.Data.FolderId]::new($TargetFolder, $TargetMailbox))
        }
        catch {
            Write-Error -Message "The target folder '$TargetFolder' does not exist in mailbox '$TargetMailbox' or you do not have FullAccess permission on the target mailbox."
            return
        }

        # Check, if we have to split the result array
        if (($MailItems.size | Measure-Object -Sum).Sum -gt $MaxBatchSize) {
            $mailItemBatches = Split-EOATMailItem -MailItems $MailItems -MaxBatchSize $MaxBatchSize
        }
        else {
            $mailItemBatches = [System.Collections.Generic.List[PSCustomObject]]::new()
            $null = $mailItemBatches.Add($MailItems)
        }

        # If LogEnabled, check if the log file already exists. If so, stop the script and ask the user to delete the log file.
        if($LogEnabled) {
            if (Test-Path -Path $LogFilePath) {
                Write-Error -Message "The file '$LogFilePath' already exists. Please delete the log file and try again."
                return
            }
        }
    }

    # Move mail items
    Process {

        foreach ($mailItemBatch in $mailItemBatches) {
            Write-Verbose -Message "Moving $($mailItemBatch.Count) mail items to folder '$TargetFolder' in mailbox '$TargetMailbox'"
            foreach ($mailItem in $mailItemBatch) {
                # Write progress
                $progressProps = @{
                    Activity        = "Batch $($mailItemBatches.IndexOf($mailItemBatch) + 1) of $($mailItemBatches.count)"
                    Status          = "Moving mail item '$($mailItem.Subject)' to folder '$TargetFolder' in mailbox '$TargetMailbox'"
                    PercentComplete = (($mailItemBatch.IndexOf($mailItem) + 1) / $mailItemBatch.Count * 100)
                }
                Write-Progress @progressProps
                
                $retryCount = 0
                $moveSuccess = $false
                do {
                    # If the method fails three times with the error "The server cannot service this request right now. Try again later.", the function will return an error
                    if($retryCount -eq 3) {
                        Write-Error -Message "The method failed with error 'Try again later' three times. The function will return an error."
                        return
                    }

                    try {
                        # Move mail item
                        $null = $mailItem.Move($targetFolderId.Id)
                    }
                    catch {
                        if ($_.Exception.Message -like "*The server cannot service this request right now. Try again later.*") { # If the error message contains "The server cannot service this request right now. Try again later.", we will wait and retry the method
                            $backOffMilliseconds = $_.Exception.InnerException.BackOffMilliseconds + 100
                            Write-Warning -Message "The method failed with error '$($_.Exception.Message)'. Waiting $backOffMilliseconds milliseconds before retrying the method."
                            Start-Sleep -Milliseconds $backOffMilliseconds
                            $retryCount++
                            continue
                        }
                        else { # If the error message does not contain "The server cannot service this request right now. Try again later.", we will write the error message and stack trace to the console and return
                            throw $_
                        }
                    }

                    # If the method was successful, we will set $moveSuccess to $true
                    $moveSuccess = $true
                } while (-not $moveSuccess)

                if ($LogEnabled) {
                    # Export log info to CSV file
                    @{
                        SourceMailbox      = $script:SourceMailbox
                        # SourceFolder     = '' # SourceFolder Name is not available on this object. Will be ignored for now for performance reasons. SourceFolderId can be used to get the SourceFolder Name using Get-EOATMailFolder.
                        SourceFolderId     = $mailItem.ParentFolderId.UniqueId
                        TargetMailbox      = $TargetMailbox
                        TargetFolder       = $TargetFolder
                        TargetFolderId     = $targetFolderId.Id
                        SourceMailItemId   = $mailItem.Id.UniqueId
                        Sender             = $mailItem.From.Address
                        Subject            = $mailItem.Subject
                        Received           = $mailItem.DateTimeReceived.ToString('yyyy-MM-dd-hh-mm-ss')
                        SizeInMB           = "{0:N2}" -f ($mailItem.Size / 1024 / 1024)
                        CurrentWindowsUser = "$env:USERDOMAIN\$env:USERNAME"
                    } | Export-Csv -Path $LogFilePath -Append -NoTypeInformation -Encoding utf8 -Delimiter $LogDelimiter
                }
            }

            # if we had only one batch or if we are at the last batch, we will not ask the user to continue with the next batch
            if (($mailItemBatches.count -eq 1) -or ($mailItemBatches.IndexOf($mailItemBatch) -eq ($mailItemBatches.count - 1))) {
                Write-Verbose -Message "No more batches to process."
                return
            }
                
            # Check the target folder count every x seconds before continuing with the next batch, to prevent the mailbox from exceeding the quota
            # Ask the script user if we should continue, as soon as the target folder is empty
            if($CheckTargetFolderEmpty) {
                $targetFolderItemcount = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, [Microsoft.Exchange.WebServices.Data.FolderId]::new($TargetFolder, $TargetMailbox)) | Select-Object -ExpandProperty TotalCount
                while ($targetFolderItemcount -ne 0) {
                    Write-Warning -Message "Target folder '$TargetFolder' in mailbox '$TargetMailbox' is not empty. Waiting $WaitTime seconds before continuing with the next batch."
                    Start-Sleep -Seconds $WaitTime
                    $targetFolderItemcount = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, [Microsoft.Exchange.WebServices.Data.FolderId]::new($TargetFolder, $TargetMailbox)) | Select-Object -ExpandProperty TotalCount
                }
            }

            # If confirm is set to false, we will continue with the next batch immediately
            if ($ConfirmPreference -eq "None") {
                continue
            }
        
            # Ask the script user if we should continue, if the target folder is empty
            $continue = $false
            while ($continue -eq $false) {
                $continuationPrompt = ""
                if($CheckTargetFolderEmpty) {
                    $continuationPrompt += "Target folder '$TargetFolder' in mailbox '$TargetMailbox' is empty.`n"
                }
                $continuationPrompt += 'Do you want to continue with the next batch? (Y/N)'
                $continue = Read-Host -Prompt $continuationPrompt
                if ($continue -eq 'Y') {
                    $continue = $true
                }
                elseif ($continue -eq 'N') {
                    $continue = $true
                    Write-Verbose -Message "Script execution aborted by user."
                    return
                }
                else {
                    $continue = $false
                }
            }
        }
    }
    End {
        if($LogEnabled) {
            Write-Verbose -Message "Log file was created: $LogFilePath"
        }
    }
}