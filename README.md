<!-- Requirements -->
## Requirements
This HelloID Service Automation Delegated Form uses the [Exchange Online PowerShell V2 module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)

<!-- Description -->
## Description
This HelloID Service Automation Delegated Form provides an Exchange Online report containing the mailboxes to which the specified user has access.
Retrieving this information from Exchange takes so much time that we will not show it directly in the overview.
We will send an email with the overview as a CSV file attachment.
The following options are available:
 1. Overview of mailboxes that match this report
 2. Send the overview as CSV attachment in a mail to the requester's mail

<!-- TABLE OF CONTENTS -->
## Table of Contents
* [Description](#description)
* [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  * [Getting started](#getting-started)
* [Post-setup configuration](#post-setup-configuration)
* [Manual resources](#manual-resources)


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

 _Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_


### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.


## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)
<table>
  <tr><td><strong>Variable name</strong></td><td><strong>Example value</strong></td><td><strong>Description</strong></td></tr>
  <tr><td>ExchangeOnlineAdminUsername</td><td>user@domain.com</td><td>Exchange admin account</td></tr>
  <tr><td>ExchangeOnlineAdminPassword</td><td>********</td><td>Exchange admin password</td></tr>
</table>

## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source 'mailbox-generate-users-with-permission-userprincipalname'
This Powershell data source runs an Exchange Online and AD query to select the user accounts that match this report.

# HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/
