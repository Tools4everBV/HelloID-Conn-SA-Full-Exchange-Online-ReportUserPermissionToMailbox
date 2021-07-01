# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("Users","HID_administrators") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Exchange Online","Mailbox Reporting") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> ExchangeOnlineAdminUsername
$tmpName = @'
ExchangeOnlineAdminUsername
'@ 
$tmpValue = @'
svc_helloid@zorgnetonline.nl
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #2 >> ExchangeOnlineAdminPassword
$tmpName = @'
ExchangeOnlineAdminPassword
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});


#make sure write-information logging is visual
$InformationPreference = "continue"
# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic}
    Write-Information "Using prefilled API credentials"
} else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key}
    Write-Information "Using manual API credentials"
}
# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
} else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}
# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  
 
function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )
    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body
    
            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid
            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
        Write-Error "Variable '$Name', message: $_"
    }
}
function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}
    
        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task
            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = [Object[]]($Variables | ConvertFrom-Json);
            }
            $body = ConvertTo-Json -InputObject $body
    
            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid
            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }
    $returnObject.Value = $taskGuid
}
function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }
    
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
      
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = [Object[]]($DatasourceModel | ConvertFrom-Json);
                automationTaskGUID = $AutomationTaskGuid;
                value              = [Object[]]($DatasourceStaticValue | ConvertFrom-Json);
                script             = $DatasourcePsScript;
                input              = [Object[]]($DatasourceInput | ConvertFrom-Json);
            }
            $body = ConvertTo-Json -InputObject $body
      
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
              
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
      Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }
    $returnObject.Value = $datasourceGuid
}
function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = [Object[]]($FormSchema | ConvertFrom-Json)
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }
    $returnObject.Value = $formGuid
}
function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][String][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                accessGroups    = [Object[]]($AccessGroups | ConvertFrom-Json);
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
            }    
            $body = ConvertTo-Json -InputObject $body
    
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true
            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }
    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #><# End: HelloID Data sources #>

<# Begin: Dynamic Form "Mailbox - List mailbox that user has permissions to" #>
$tmpSchema = @"
[{"templateOptions":{},"type":"markdown","summaryVisibility":"Show","body":"Retrieving this information from Exchange takes so much time that we will not show it directly in the overview.\nYou will receive an email  with the overview as a CSV file attachment","requiresTemplateOptions":false,"requiresKey":false},{"key":"userPrincipalName","templateOptions":{"label":"UserPrincipalName of User","required":true},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Mailbox - List mailbox that user has permissions to
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
foreach($group in $delegatedFormAccessGroupNames) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $delegatedFormAccessGroupGuid = $response.groupGuid
        $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
        
        Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
    } catch {
        Write-Error "HelloID (access)group '$group', message: $_"
    }
}
$delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Compress)
$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
Mailbox - List mailbox that user has permissions to
'@
Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-list" -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

<# Begin: Delegated Form Task #>
if($delegatedFormRef.created -eq $true) { 
	$tmpScript = @'
$tempFileLocation = "C:\Temp\HelloID\MailboxPermissions" 

# Send mail parameters
$mailSmtpServer = "smtprelay.enyoi.local"
$mailSmtpPort = 25
$mailUseSsl = $false
#$mailSmtpUsername =  ""
#$mailSmtpPassword = ""
$mailEncoding = "UTF8"
$mailFrom = "HelloID@enyoi.nl"
$mailTo = $RequesterMail
$mailCC = ""
$mailBCC = ""


# Connect to Office 365
try{
    Hid-Write-Status -Event Information -Message "Connecting to Office 365.."

    $module = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($ExchangeOnlineAdminUsername, $securePassword)

    $exchangeSession = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -TrackPerformance:$false -ErrorAction Stop 

    Hid-Write-Status -Event Success -Message "Successfully connected to Office 365"
}catch{
    throw "Could not connect to Exchange Online, error: $_"
}

# Get Exchange mailbox permissions
try {
    Hid-Write-Status -Event Information -Message "Searching for user: $UserPrincipalName"
    $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$($UserPrincipalName)'" -Properties MemberOf
    
    # Can't be used because of a bug in PS 5.1
    #$adGroups = Get-ADPrincipalGroupMembership -Identity $adUser
    $adGroups = New-Object System.Collections.ArrayList
    foreach($group in $adUser.MemberOf) {
        $null = $adGroups.Add((Get-ADGroup $group)) # direct output to NULL or else we'll get an int
    }

    $adGroupsWithMailboxPermissions = $adGroups | Where-Object { $_.Name -Like "Mbx_*" }

    # Get All mailboxes
    Hid-Write-Status -Event Information -Message "Gathering all mailboxes.."
    $mailboxes = Get-EXOMailbox -PropertySets Minimum,Delivery -ResultSize Unlimited -ErrorAction Stop
    $mailBoxesGrouped = $mailboxes | Group-Object -Property Identity -AsHashTable
    [System.Collections.ArrayList]$allMailboxesWithPermission = @()


    # List all users with Full Access permissions
    Hid-Write-Status -Event Information -Message "Gathering Full Access Permissions.."
    [System.Collections.ArrayList]$mailboxesFullAccess = @()    
    $fullAccessPermissions = $mailboxes | Get-EXOMailboxPermission | Where-Object { ($_.AccessRights -like "*fullaccess*") -and -not ($_.Deny -eq $true) -and -not ($_.User -match "NT AUTHORITY") } -ErrorAction Stop
    foreach($fullAccessPermission in $fullAccessPermissions){
        if($fullAccessPermission.User -like $adUser.UserPrincipalName){
            $mailbox = $mailBoxesGrouped."$($fullAccessPermission.Identity)"

            if($mailbox){
                $mailboxFullAccess = New-Object PsObject

                $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name Permission -Value "Full Access" -Force
                $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName -Force
                $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName -Force
                $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias -Force
                $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $mailbox.PrimarySMTPAddress -Force
                $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name InheritedFromGroup -Value $false -Force
                $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name Group -Value $null -Force

                $null = $mailboxesFullAccess.Add($mailboxFullAccess)
            }
        }

        if($fullAccessPermission.User -in $adGroupsWithMailboxPermissions.Name){
            foreach($adGroup in $adGroupsWithMailboxPermissions){
                if($fullAccessPermission.User -like $adGroup.Name){
                    $mailbox = $mailBoxesGrouped."$($fullAccessPermission.Identity)"

                    if($mailbox){
                        $mailboxFullAccess = New-Object PsObject

                        $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name Permission -Value "Full Access" -Force
                        $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName -Force
                        $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName -Force
                        $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias -Force
                        $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $mailbox.PrimarySMTPAddress -Force
                        $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name InheritedFromGroup -Value $false -Force
                        $mailboxFullAccess | Add-Member -MemberType NoteProperty -Name Group -Value $null -Force

                        $null = $mailboxesFullAccess.Add($mailboxFullAccess)
                    }
                }
            }
        }
    }

    Hid-Write-Status -Event Information -Message "Mailboxes which user has Full Access permissions to: $($mailboxesFullAccess.Name.Count)"
    
    if($mailboxesFullAccess.Name.Count -gt 0){
        foreach($entry in $mailboxesFullAccess){
            $null = $allMailboxesWithPermission.Add($entry)
        }
    }



    # List all mailboxes to which a user has Send As permissions
    Hid-Write-Status -Event Information -Message "Gathering Send As Permissions.."
    [System.Collections.ArrayList]$mailboxesSendAs = @()
    $SendAsPermissions = Get-EXORecipientPermission -ResultSize Unlimited -AccessRights SendAs
    foreach($SendAsPermission in $SendAsPermissions){
        if($SendAsPermission.trustee -like $adUser.UserPrincipalName){
            $mailbox = $mailBoxesGrouped."$($SendAsPermission.Identity)"

            if($mailbox){
                $mailBoxSendAs = New-Object PsObject

                $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name Permission -Value "Send As" -Force
                $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName -Force
                $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName -Force
                $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias -Force
                $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $mailbox.PrimarySMTPAddress -Force
                $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name InheritedFromGroup -Value $false -Force
                $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name Group -Value $null -Force

                $null = $mailboxesSendAs.Add($mailBoxSendAs)
            }
        }

        if($SendAsPermission.trustee -in $adGroupsWithMailboxPermissions.Name){
            foreach($adGroup in $adGroupsWithMailboxPermissions){
                if($SendAsPermission.trustee -like $adGroup.Name){
                    $mailbox = $mailBoxesGrouped."$($SendAsPermission.Identity)"

                    if($mailbox){
                        $mailBoxSendAs = New-Object PsObject

                        $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name Permission -Value "Send As" -Force
                        $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName -Force
                        $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName -Force
                        $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias -Force
                        $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $mailbox.PrimarySMTPAddress -Force
                        $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name InheritedFromGroup -Value $false -Force
                        $mailBoxSendAs | Add-Member -MemberType NoteProperty -Name Group -Value $null -Force

                        $null = $mailboxesSendAs.Add($mailBoxSendAs)
                    }
                }
            }
        }
    }

    Hid-Write-Status -Event Information -Message "Mailboxes which user has Send As permissions to: $($mailboxesSendAs.Name.Count)"
    
    if($mailboxesSendAs.Name.Count -gt 0){
        foreach($entry in $mailboxesSendAs){
            $null = $allMailboxesWithPermission.Add($entry)
        }
    }


    # List all mailboxes to which a particular security principal has Send on behalf of permissions
    Hid-Write-Status -Event Information -Message "Gathering Send On Behalf Permissions.."
    [System.Collections.ArrayList]$mailboxesSendOnBehalf = @()

    foreach($mailbox in $mailboxes){
        if(![String]::IsNullOrEmpty($mailbox.GrantSendOnBehalfTo)){
            if($mailbox.GrantSendOnBehalfTo -match "$($adUser.Name)"){
                $mailBoxSendOnBehalf = New-Object PsObject

                $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name Permission -Value "Send On Behalf" -Force
                $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName -Force
                $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName -Force
                $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias -Force
                $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $mailbox.PrimarySMTPAddress -Force
                $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name InheritedFromGroup -Value $false -Force
                $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name Group -Value $null -Force

                $null = $mailboxesSendOnBehalf.Add($mailBoxSendOnBehalf)
            }

            if($mailbox.GrantSendOnBehalfTo -in $adGroupsWithMailboxPermissions.Name){
                foreach($adGroup in $adGroupsWithMailboxPermissions){
                    if($mailbox.GrantSendOnBehalfTo -match "$($adGroup.Name)"){
                        $mailBoxSendOnBehalf = New-Object PsObject

                        $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name Permission -Value "Send On Behalf" -Force
                        $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName -Force
                        $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName -Force
                        $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias -Force
                        $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $mailbox.PrimarySMTPAddress -Force
                        $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name InheritedFromGroup -Value $false -Force
                        $mailBoxSendOnBehalf | Add-Member -MemberType NoteProperty -Name Group -Value $null -Force

                        $null = $mailboxesSendOnBehalf.Add($mailBoxSendOnBehalf)
                    }
                }
            }
        }
    }

    Hid-Write-Status -Event Information -Message "Mailboxes which user has Send On Behalf permissions to: $($mailboxesSendOnBehalf.Name.Count)"
    
    if($mailboxesSendOnBehalf.Name.Count -gt 0){
        foreach($entry in $mailboxesSendOnBehalf){
            $null = $allMailboxesWithPermission.Add($entry)
        }
    }

    # Create temp csv file
    $currentDate = (Get-Date).ToString("yyyy_MM_dd_HHmmss")

    if(!(Test-Path -Path $tempFileLocation)){
        $newfile = New-Item $tempFileLocation -Force -Confirm:$false
    }

    $fileName = "$tempFileLocation\$UserPrincipalName $currentDate.csv"
    $allMailboxesWithPermission | Sort-Object Permission,DisplayName,UserPrincipalName | Export-Csv -Path $fileName -Delimiter ';' -Encoding UTF8 -NoTypeInformation -Force -Confirm:$false

    try{
        # Send mail parameters
        $mailSubject = "Toegang tot mailboxen voor $UserPrincipalName"
        $mailBodyAsHtml = $true
        $mailBody = "
            <p>Beste,
            </p>
            <p>In de bijlage vindt u een CSV bestand met hierin een overzicht van de mailboxen waar $UserPrincipalName toegang toe heeft.
            </p>
            <p>Met vriendelijke groeten,<br>
            </p>
            <p>HelloID<br>
            </p>
        "

        # Send mail
        if($mailSmtpUsername -and $mailSmtpPassword){
            $mailSecurePassword = $mailSmtpPassword | ConvertTo-SecureString -asPlainText -Force
            $mailCredentials = New-Object System.Management.Automation.PSCredential($mailSmtpUsername,$mailSecurePassword)
        }

        $allParams = @{
            SmtpServer = $mailSmtpServer
            Encoding = $mailEncoding
            Port = $mailSmtpPort
            UseSsl = $mailUseSsl
            Credential = $mailCredentials
            From = $mailFrom
            To = $mailTo
            CC = $mailCC
            BCC = $mailBCC
            Subject = $mailSubject
            BodyAsHtml = $mailBodyAsHtml
            Body = $mailBody
            Attachments = $fileName
        }

        $filledParams = @{}
        foreach($key in $allParams.keys){
            if(![string]::IsNullOrEmpty($allParams.$key)){
                $filledParams += @{ "$key" = $($allParams.$key) }
            }
        }
    
        Send-MailMessage @filledParams -ErrorAction Stop
    
        Hid-Write-Status -Event Success -Message "Successfully sent mail to [$mailTo]"   
    }catch{
        Hid-Write-Status -Event Warning -Message $aduser
        Hid-Write-Status -Event Error -Message "Error sending mail to [$mailTo]: $_"
    }


    # Clean up temp csv file
    Remove-Item -Path $fileName -Force -Confirm:$false
} catch {
    HID-Write-Status -Message "Error gathering mailbox permissions for [$UserPrincipalName]. Error: $_" -Event Error
    HID-Write-Summary -Message "Error gathering mailbox permissions for [$UserPrincipalName]" -Event Failed
} finally {
    Hid-Write-Status -Event Information -Message "Disconnecting from Office 365.."
    $exchangeSessionEnd = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false -ErrorAction Stop
    Hid-Write-Status -Event Success -Message "Successfully disconnected from Office 365"
}
'@; 

	$tmpVariables = @'
[{"name":"RequesterMail","value":"{{requester.contactEmail}}","secret":false,"typeConstraint":"string"},{"name":"UserPrincipalName","value":"{{form.userPrincipalName}}","secret":false,"typeConstraint":"string"}]
'@ 

	$delegatedFormTaskGuid = [PSCustomObject]@{} 
$delegatedFormTaskName = @'
List mailbox that user has permissions to and Send as mail attachment
'@
	Invoke-HelloIDAutomationTask -TaskName $delegatedFormTaskName -UseTemplate "False" -AutomationContainer "8" -Variables $tmpVariables -PowershellScript $tmpScript -ObjectGuid $delegatedFormRef.guid -ForceCreateTask $true -returnObject ([Ref]$delegatedFormTaskGuid) 
} else {
	Write-Warning "Delegated form '$delegatedFormName' already exists. Nothing to do with the Delegated Form task..." 
}
<# End: Delegated Form Task #>
