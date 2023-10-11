# ExchangeOnlineArchiveTransfer

Copy or move items of a Exchange Online mailbox (primary mailbox AND archive mailbox possible) to any folder of any target mailbox in your tenant.

## What this module can do

This module can be used to read the contents of a mailbox and create a list of folders or items (without the actual email body).
This module can also be used to copy and move the contents of a mailbox to a target mailbox.

## Getting started

### Requirements

There is only a single module required to use ExchangeOnlineArchiveTransfer: Azure.Function.Tools
This module is used for the authentication process and will be installed automatically, if it is not already installed.

Install the ExchangeOnlineArchiveTransfer module from PowerShell Gallery using the following command:

```powershell
Install-Module -Name ExchangeOnlineArchiveTransfer
```

### Authentication

This module uses the Azure AD v2.0 endpoint to authenticate to Exchange Online.
Therefore, you must create a new Azure AD app registration and grant it the required permissions.

- Create App Registration
  - Set Name to a friendly name for your app.
  - Enable "public client flows" for Device Code Flow.
  - Add API permission: "Office 365 Exchange Online" -> "EWS.AccessAsUser.All" (Delegate permission)

- Configure corresponding Enterprise Application to your needs, for example:
  - Properties --> Enable "Assignment required?"
    - Users and Groups --> Add users or groups, which will be allowed to use this app registration

More information abot the authentication process at [How to authenticate an EWS application by using OAuth](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

## How it works

This module uses the EWS Managed API to connect to a mailbox and read the contents.
The EWS Managed API is a .NET library that can be used to connect to Exchange Online.
The EWS Managed API is used in combination with the Azure AD v2.0 endpoint to authenticate to Exchange Online.
The EWS Managed API is already downloaded and part of this module. If you want to get your own copy of the library go to [the link in the GitHub repo](https://github.com/OfficeDev/ews-managed-api/tree/master)

Using this library, the module can connect to a mailbox and read the contents, list folders and items and copy or move items to a target mailbox.

## How to use

Step 1: Connect to Exchange Web Service

Connect to Exchange Web Service using the following command:

```powershell
Connect-EOATExchangeWebService -ApplicationId "00000000-0000-0000-0000-000000000000" -TenantId "00000000-0000-0000-0000-000000000000" -MailboxName "source@domain.com"
```

Step 2: List all folders and select the ones, you want to use as source

```powershell
$folders = Get-EOATMailFolder -SearchBase ArchiveMsgFolderRoot -ShowGui
```

Step 3: Get a list of mail items from a mail folder from a specific date. Do not list the StartDate and EndDate parameters to get all items.

```powershell
$items = Get-EOATMailItem -StartDate "01/01/2023" -EndDate "08/01/2023" -MailFolders $folders
```

Step 4a: Copy the items to a target mailbox and log the copied items to a CSV file. Check out the comment-based help of the function for more information, like copying the items in batches of specific size.

```powershell
Copy-EOATMailItemToOtherMailbox -MailItems $items -TargetMailbox 'destination@domain.com' -TargetFolder Inbox -LogEnabled
```

Step 4a: Copy the items to a target mailbox and log the copied items to a CSV file. Check out the comment-based help of the function for more information, like moving the items in batches of specific size.

```powershell
Move-EOATMailItemToOtherMailbox -MailItems $items -TargetMailbox 'destination@domain.com' -TargetFolder Inbox -LogEnabled
```

## This individual functions can do more

Checkout the comment-based help of each function for more information.

```powershell
Get-Help -Name Connect-EOATExchangeWebService -Full
Get-Help -Name Disconnect-EOATExchangeWebService -Full
Get-Help -Name Get-EOATMailFolder -Full
Get-Help -Name Get-EOATMailItem -Full
Get-Help -Name Move-EOATMailItemToOtherMailbox -Full
Get-Help -Name Copy-EOATMailItemToOtherMailbox -Full
```
