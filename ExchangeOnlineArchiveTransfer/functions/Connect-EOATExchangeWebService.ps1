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

    .PARAMETER PassThru
    Pass the EWS service object and token to the pipeline
    
    .EXAMPLE
    Connect-EOATExchangeWebService -ApplicationId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000' -MailboxName 'user@contoso.com'

    This example connects to Exchange Online EWS with the Azure AD application '00000000-0000-0000-0000-000000000000' and the Azure AD tenant '00000000-0000-0000-0000-000000000000' and the mailbox 'user@contoso.com'

    .EXAMPLE

    Connect-EOATExchangeWebService -ApplicationId '00000000-0000-0000-0000-000000000000' -TenantId '00000000-0000-0000-0000-000000000000' -MailboxName 'user@contoso.com' -PassThru

    $ewsObj, $tokenObj = This example connects to Exchange Online EWS with the Azure AD application '00000000-0000-0000-0000-000000000000' and the Azure AD tenant '00000000-0000-0000-0000-000000000000' and the mailbox 'user@contoso.com' and passes the EWS service object and token to the pipeline


    #>
    [CmdletBinding()]
    [OutputType([Microsoft.Exchange.WebServices.Data.ExchangeService])]
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
        $Scopes = 'https://outlook.office365.com/EWS.AccessAsUser.All',

        [Switch]
        $PassThru
    )

    # Verbose output
    Write-Verbose -Message "Connecting to Exchange Online EWS with ApplicationId '$ApplicationId' and TenantId '$TenantId'"

    # Get token
    Write-Warning -Message "Please use the credentials of the source mailbox to connect to Exchange Online EWS. This mailbox must have FullAccess permission on the target mailbox as well, to be able to access the target mailbox."
    $ewsEntraService = @{
        Name          = 'ExoEwsService'
        ServiceUrl    = 'https://graph.microsoft.com/v1.0'
        Resource      = 'https://outlook.office365.com'
        DefaultScopes = @()
        HelpUrl       = ''
        Header        = @{}
        NoRefresh     = $false
    }
    $null = Register-EntraService @ewsEntraService
    $script:EwsToken = Connect-EntraService -ClientID $ApplicationId -TenantID $TenantId -Service Graph -Scopes $Scopes -DeviceCode -PassThru

    # create EWS service object
    $script:EwsService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
 
    #Use Modern Authentication
    $EwsService.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EwsToken.AccessToken
 
    #Check EWS connection
    $EwsService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
    Write-Verbose "Starting Autodiscover for $MailboxName"
    $EwsService.AutodiscoverUrl($MailboxName, { $true }) # use overload with callback
    #EWS connection is Success if no error returned.

    # set script variable for source mailbox
    $script:SourceMailbox = $MailboxName

    # Verbose output
    Write-Verbose -Message "Successfully connected to Exchange Online EWS"

    # Pass EWS service object and token to the pipeline
    if ($PassThru) {
        $EwsService
        $EwsToken
    }
}