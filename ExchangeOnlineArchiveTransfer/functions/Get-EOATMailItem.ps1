function Get-EOATMailItem {
    <#
    .SYNOPSIS
    Get a list of mail items from a mail folder.
    
    .DESCRIPTION
    Get a list of mail items from a mail folder.
    This function returns a list of mail items from a mail folder using the Exchange Web Services (EWS) API.
    The items are automatically returned from new to old by the Exchange Web Service.
    You can use the returned list of mail items to select the items, which you want to use for further steps.
    
    .PARAMETER MailFolders
    PSCustomObject array. Default value is $null.
    This parameter defines the mail folders, from which the function gets the mail items.
    You can use the function 'Get-EOATMailFolder' to get a list of mail folders.
    
    .PARAMETER ResultSizePerFolder
    Integer value. Default value is null.
    This parameter defines the number of mail items, which are returned by the function per provided folder.
    So, if you provide three mail folders, the function returns 3000000 mail items.
    If you provide three mail folders and set the ResultSizePerFolder to 100, the function returns 300 mail items.
    If you do not provide a value for this parameter, the function returns all mail items per folder.
    
    .PARAMETER StartDate
    DateTime value. Default value is $null.
    This parameter defines the start date of the mail items, which are returned by the function.
    The start date is defined as the "received date" of the mail item.
    The function will search for mails in the provided folder(s) based on the provided ResultSizePerFolder and start date.

    .PARAMETER EndDate
    DateTime value. Default value is $null.
    This parameter defines the end date of the mail items, which are returned by the function.
    The end date is defined as the "received date" of the mail item.
    The function will search for mails in the provided folder(s) based on the provided ResultSizePerFolder and end date.

    .PARAMETER Service
    ExchangeService object. Default value is the script variable $script:EwsService.
    This parameter defines the ExchangeService object, which is used to connect to Exchange Online EWS.
    This parameter will be set automatically, if you use the function 'Connect-EOATExchangeWebService' to connect to Exchange Online EWS.
    
    .EXAMPLE
    Get-EOATMailItem -MailFolders $MailFolders

    This example returns a list of mail items from the mail folders defined in the variable $MailFolders.

    .EXAMPLE
    Get-EOATMailItem -MailFolders $MailFolders

    This example returns a list of mail items from the mail folders defined in the variable $MailFolders. The number of returned mail items is 1000000.

    .EXAMPLE
    Get-EOATMailFolder -SearchBase ArchiveMsgFolderRoot -ShowGui | Get-EOATMailItem -ResultSize 150

    This example returns a list of mail items from the mail folders selected in the GUI window. The number of returned mail items is 150 per provided folder.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([System.Collections.Generic.List[Object]])]
    param (

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Default')]
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'DateRange')]
        [PSCustomObject[]]
        $MailFolders,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'DateRange')]
        [Int32]
        $ResultSizePerFolder,

        [Parameter(Mandatory = $true, ParameterSetName = 'DateRange')]
        [datetime]
        $StartDate,

        [Parameter(Mandatory = $true, ParameterSetName = 'DateRange')]
        [datetime]
        $EndDate,

        [Parameter(ParameterSetName = 'Default', DontShow)]
        [Parameter(ParameterSetName = 'DateRange', DontShow)]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]
        $Service = $script:EwsService
    )

    Begin {

        # Check if StartDate and EndDate are set and if StartDate is before EndDate
        if ($StartDate -and $EndDate -and $StartDate -gt $EndDate) {
            Write-Error -Message "The StartDate must be before the EndDate."
            return
        }

        # Check if input object for Mailfolders was provided. If not, ignore for now
        if ($null -eq $MailFolders) {
            return
        }

        # Check if input object contains property 'Id'
        if (($MailFolders | Get-Member -MemberType NoteProperty).Name -notcontains "Id") {
            Write-Error -Message "The input object must contain a property named 'Id'. Please use the function 'Get-EOATMailFolder' to retrieve the folders."
            return
        }
    }

    Process {
        # Prepare return value
        $returnValue = [System.Collections.Generic.List[Object]]::new()

        # List mail items in selected folders
        foreach ($mailFolder in $MailFolders) {
            # prepare variable to count items in this folder
            $returnedItemsInThisFolder = 0

            # Define SubFolderId
            $SubFolderId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId($mailFolder.id)
            
            #Define ItemView to retrieve items in pages
            $pageSize = 1000
            $offSet = 0
            $ivItemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView(($pageSize + 1), $offSet)
            $ivItemView.PropertySet = New-Object -TypeName Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived)
            $ivItemView.OrderBy.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived, 'Descending')

            # Prepare Searchfilter, if StartDate and EndDate were provided
            if ($StartDate -and $EndDate) {
                $searchQuery = "received:>=$($StartDate.ToString("MM/dd/yyyy")) AND received:<=$($EndDate.ToString("MM/dd/yyyy"))"
            }

            # Get items from folder and add them to list
            $continueLoop = $true
            do {
                if ($StartDate -and $EndDate) {
                    $foundItems = $Service.FindItems($SubFolderId, $searchQuery, $ivItemView)
                }
                else {
                    $foundItems = $Service.FindItems($SubFolderId, $ivItemView)
                }

                # Check if more items are available
                $continueLoop = $foundItems.MoreAvailable
                if ($foundItems.MoreAvailable) {
                    $ivItemView.Offset += $pageSize
                }

                foreach ($foundItem in $foundItems) {
                    # Add item to list
                    $returnValue.Add($foundItem)
                    $returnedItemsInThisFolder++
                    
                    if($returnedItemsInThisFolder -eq $ResultSizePerFolder) {
                        Write-Verbose -Message "ResultSize of $ResultSizePerFolder items reached. Ending search."
                        
                        # end outer do-while loop
                        $continueLoop = $false
                        
                        # end inner foreach loop
                        break
                    }
                }

                Write-Verbose -Message "Items Count: $($foundItems.Items.Count), Offset: $($ivItemView.Offset)"

            } while ($continueLoop)
        }

        # return the selected items as list
        $returnValue

    }
}