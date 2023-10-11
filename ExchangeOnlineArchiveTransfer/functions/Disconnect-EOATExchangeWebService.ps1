function Disconnect-EOATExchangeWebService {
    <#
    .SYNOPSIS
    Disconnect from Exchange Online EWS. This function clears the script variables.
    
    .DESCRIPTION
    Disconnect from Exchange Online EWS. This function clears the script variables.

    .EXAMPLE
    Disconnect-EOATExchangeWebService

    This example disconnects from Exchange Online EWS, by clearing the script variables.

    #>
    [CmdletBinding()]
    param()

    # Clear script variables
    $script:EwsService = $null
    $script:SourceMailbox = $null

    # Verbose output
    Write-Verbose -Message "Successfully disconnected from Exchange Online EWS"
}