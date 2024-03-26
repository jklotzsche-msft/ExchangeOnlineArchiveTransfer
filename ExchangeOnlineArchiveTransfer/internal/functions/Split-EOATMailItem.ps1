function Split-EOATMailItem {
    <#
    .SYNOPSIS
    Split the mail items into batches
    
    .DESCRIPTION
    This function will split the mail items into batches.
    The default batch size is 90 GB, which is also the recommended maximum batch size.
    
    .PARAMETER MailItems
    PSCustomObject array. Default value is $null.
    This parameter defines the mail items, which should be split into batches.
    You can use the function 'Get-EOATMailItem' to get a list of mail items.
    
    .PARAMETER MaxBatchSize
    Integer value. Default value is 96636764160 (90 GB).
    This parameter defines the maximum batch size in Byte, which is used to retrieve the mail items.
    If more than 90 GB of mail items are send to this function, the mails will be split into batches of 90 GB.
    
    .EXAMPLE
    Split-EOATMailItem -MailItems $MailItems -MaxBatchSize 90GB

    This example splits the mail items into batches of 90 GB.

    .EXAMPLE
    Split-EOATMailItem -MailItems $MailItems -MaxBatchSize 20GB

    This example splits the mail items into batches of 20 GB.
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[PSCustomObject]])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]
        $MailItems,

        [int]
        $MaxBatchSize = 90GB
    )

    begin {
        $batches = [System.Collections.Generic.List[PSCustomObject]]::new()
        $currentBatch = [System.Collections.Generic.List[PSCustomObject]]::new()
        $currentbatchSize = 0
    }
    process {
        foreach ($mailItem in $MailItems) {
            # add mail item to current batch
            $null = $currentBatch.Add($mailItem)
            # add mail item size to current batch size
            $currentbatchSize += $mailItem.Size

            # check if current batch size is greater than max batch size
            if ($currentbatchSize -ge $MaxBatchSize) {
                # add current batch to batches
                $null = $batches.Add($currentBatch)
                # reset current batch size
                $currentbatchSize = 0
                # reset current batch
                $currentBatch = [System.Collections.Generic.List[PSCustomObject]]::new()
            }
        }
    }
    end {
        # add the last batch to batches (if it is not empty), as it is not added in the foreach loop above, because the last batch is not greater than the max batch size
        if ($currentBatch.Count -gt 0) {
            $null = $batches.Add($currentBatch)
        }
        
        # return all batches
        $batches
    }
}