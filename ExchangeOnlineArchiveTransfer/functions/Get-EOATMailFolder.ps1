function Get-EOATMailFolder {
    <#
    .SYNOPSIS
    Get a list of mail folders from a mailbox.
    
    .DESCRIPTION
    Get a list of mail folders from a mailbox.
    This function returns a list of mail folders from a mailbox using the Exchange Web Services (EWS) API.
    You can use the returned list of mail folders to select the folders, which you want to use for further steps.
    
    .PARAMETER SearchBase
    String value of the WellKnownFolderName enum. Default value is 'ArchiveMsgFolderRoot'.
    This parameter defines the root folder, from which the function starts to search for mail folders.
    
    .PARAMETER FolderViewCount
    Integer value. Default value is 100.
    This parameter defines the number of mail folders, which are returned by the function.
    
    .PARAMETER ShowGui
    Switch parameter. Default value is $false.
    If this switch is set, the function returns a list of mail folders in a GUI window. You can select the folders, which you want to use for further steps.
    
    .PARAMETER Service
    ExchangeService object. Default value is the script variable $script:EwsService.
    This parameter defines the ExchangeService object, which is used to connect to Exchange Online EWS.
    This parameter will be set automatically, if you use the function 'Connect-EOATExchangeWebService' to connect to Exchange Online EWS.
    
    .EXAMPLE
    Get-EOATMailFolder -SearchBase 'ArchiveMsgFolderRoot' -ShowGui

    This example returns a list of mail folders from the root folder 'ArchiveMsgFolderRoot' in a GUI window. You can select the folders, which you want to use for further steps.

    .EXAMPLE
    Get-EOATMailFolder -SearchBase 'ArchiveMsgFolderRoot' -FolderViewCount 150

    This example returns a list of mail folders from the root folder 'ArchiveMsgFolderRoot'. The number of returned mail folders is 150.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[Object]])]
    param (
        [ArgumentCompleter({
                [Microsoft.Exchange.Webservices.Data.WellKnownFolderName] | Get-Member -Static -MemberType Properties | Select-Object -ExpandProperty Name
            })]
        [String]
        $SearchBase = 'ArchiveMsgFolderRoot',

        [int]
        $FolderViewCount = 100,

        [switch]
        $ShowGui,

        [Parameter(DontShow)]
        [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]
        $Service = $script:EwsService
    )

    Process {
        $folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView($FolderViewCount)
        $folderview.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.Webservices.Data.BasePropertySet]::FirstClassProperties)
        $folderview.PropertySet.Add([Microsoft.Exchange.Webservices.Data.FolderSchema]::DisplayName)
        $folderview.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep
        $foldersResult = $Service.FindFolders([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::$SearchBase, $folderview)
 
        #List folders result
        $foundFolders = $foldersResult | Select-Object -Property DisplayName, TotalCount, UnreadCount, FolderClass, Id | Sort-Object -Property DisplayName

        # if ShowGui is set, show the folders in a GUI window...
        if ($ShowGui) {
            $foundFolders = $foundFolders | Out-GridView -Title 'Select the folders, which you want to use for further steps' -PassThru
        }
        # ...and check if user selected a folder
        if ($null -eq $foundFolders) {
            Write-Verbose -Message "No folder was selected. Please select at least one folder."
            return
        }

        # Check if only one folder was selected. If so, return it as an array
        if ($foundFolders.count -eq 1) {
            $foundFolders = @($foundFolders)
        }

        # return the selected folders as array
        $foundFolders
    }
}