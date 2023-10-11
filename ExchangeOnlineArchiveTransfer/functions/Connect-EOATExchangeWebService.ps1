function Connect-EOATExchangeWebService {
    <#
    .SYNOPSIS
    Connect to Exchange Online EWS
    
    .DESCRIPTION
    Connect to Exchange Online EWS
    This function returns an EWS access token
    
    .PARAMETER ApplicationId
    Guid of the Azure AD application
    
    .PARAMETER TenantId
    Guid of the Azure AD tenant

    .PARAMETER MailboxName
    Name of the mailbox to connect to
    
    .PARAMETER Scopes
    Scope of the Azure AD application. Default value is 'https://outlook.office365.com/EWS.AccessAsUser.All'
    
    .EXAMPLE
    Connect-EOATExchangeWebService -ApplicationId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000' -MailboxName 'user@contoso.com'

    This example connects to Exchange Online EWS with the Azure AD application '00000000-0000-0000-0000-000000000000' and the Azure AD tenant '00000000-0000-0000-0000-000000000000' and the mailbox 'user@contoso.com'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Guid]
        $ApplicationId,

        [Parameter(Mandatory = $true)]
        [Guid]
        $TenantId,

        [Parameter(Mandatory = $true)]
        [String]
        $MailboxName,

        [String]
        $Scopes = 'https://outlook.office365.com/EWS.AccessAsUser.All'
    )

    # Prepare connection properties
    $connectProps = @{
        ClientId = $ApplicationId.Guid
        TenantId = $TenantId.Guid
        Scopes   = $Scopes
    }

    # Verbose output
    Write-Verbose -Message "Connecting to Exchange Online EWS with ApplicationId '$ApplicationId' and TenantId '$TenantId'"

    # Get token
    Write-Warning -Message "Please use the credentials of the source mailbox to connect to Exchange Online EWS. This mailbox must have FullAccess permission on the target mailbox as well, to be able to access the target mailbox."
    $token = Connect-DeviceCode @connectProps

    # create EWS service object
    $script:EwsService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
 
    #Use Modern Authentication
    $EwsService.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$token.access_token
 
    #Check EWS connection
    $EwsService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
    $EwsService.AutodiscoverUrl($MailboxName, { $true })
    #EWS connection is Success if no error returned.

    # set script variable for source mailbox
    $script:SourceMailbox = $MailboxName

    # Verbose output
    Write-Verbose -Message "Successfully connected to Exchange Online EWS"
}