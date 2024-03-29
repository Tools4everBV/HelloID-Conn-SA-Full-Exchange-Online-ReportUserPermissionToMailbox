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
    Write-Information -Message "Connecting to Office 365.."

    $module = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername,$securePassword)

    $exchangeSession = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -TrackPerformance:$false -ErrorAction Stop 

    Write-Information -Message "Successfully connected to Office 365"

    $Log = @{
        Action            = "MoveAccount" # optional. ENUM (undefined = default) 
        System            = "ExchangeOnline" # optional (free format text) 
        Message           = "Successfully connected to Office 365" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = ""# optional (free format text) 
        TargetIdentifier  = "" # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log

}catch{
    throw "Could not connect to Exchange Online, error: $_"
}

# Get Exchange mailbox permissions
try {
    Write-Information -Message "Searching for user: $UserPrincipalName"
    $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$($UserPrincipalName)'" -Properties MemberOf
    
    # Can't be used because of a bug in PS 5.1
    #$adGroups = Get-ADPrincipalGroupMembership -Identity $adUser
    $adGroups = [System.Collections.ArrayList]::new()
    foreach($group in $adUser.MemberOf) {
        $null = $adGroups.Add((Get-ADGroup $group)) # direct output to NULL or else we'll get an int
    }

    $adGroupsWithMailboxPermissions = $adGroups | Where-Object { $_.Name -Like "Mbx_*" }

    # Get All mailboxes
    Write-Information -Message "Gathering all mailboxes.."
    $mailboxes = Get-EXOMailbox -PropertySets Minimum,Delivery -ResultSize Unlimited -ErrorAction Stop
    $mailBoxesGrouped = $mailboxes | Group-Object -Property Identity -AsHashTable
    [System.Collections.ArrayList]$allMailboxesWithPermission = @()


    # List all users with Full Access permissions
    Write-Information -Message "Gathering Full Access Permissions.."
    [System.Collections.ArrayList]$mailboxesFullAccess = @()    
    $fullAccessPermissions = $mailboxes | Get-EXOMailboxPermission | Where-Object { ($_.AccessRights -like "*fullaccess*") -and -not ($_.Deny -eq $true) -and -not ($_.User -match "NT AUTHORITY") } -ErrorAction Stop
    foreach($fullAccessPermission in $fullAccessPermissions){
        if($fullAccessPermission.User -like $adUser.UserPrincipalName){
            $mailbox = $mailBoxesGrouped."$($fullAccessPermission.Identity)"

            if($mailbox){
                $mailboxFullAccess = [PsObject]::new()

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
                        $mailboxFullAccess = [PsObject]::new()

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

    Write-Information -Message "Mailboxes which user has Full Access permissions to: $($mailboxesFullAccess.Count)"
    
    if($mailboxesFullAccess.Count -gt 0){
        foreach($entry in $mailboxesFullAccess){
            $null = $allMailboxesWithPermission.Add($entry)
        }
    }



    # List all mailboxes to which a user has Send As permissions
    Write-Information -Message "Gathering Send As Permissions.."
    [System.Collections.ArrayList]$mailboxesSendAs = @()
    $SendAsPermissions = Get-EXORecipientPermission -ResultSize Unlimited -AccessRights SendAs
    foreach($SendAsPermission in $SendAsPermissions){
        if($SendAsPermission.trustee -like $adUser.UserPrincipalName){
            $mailbox = $mailBoxesGrouped."$($SendAsPermission.Identity)"

            if($mailbox){
                $mailBoxSendAs = [PsObject]::new()

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
                        $mailBoxSendAs = [PsObject]::new()

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

    Write-Information -Message "Mailboxes which user has Send As permissions to: $($mailboxesSendAs.Count)"
    
    if($mailboxesSendAs.Count -gt 0){
        foreach($entry in $mailboxesSendAs){
            $null = $allMailboxesWithPermission.Add($entry)
        }
    }


    # List all mailboxes to which a particular security principal has Send on behalf of permissions
    Write-Information -Message "Gathering Send On Behalf Permissions.."
    [System.Collections.ArrayList]$mailboxesSendOnBehalf = @()

    foreach($mailbox in $mailboxes){
        if(![String]::IsNullOrEmpty($mailbox.GrantSendOnBehalfTo)){
            if($mailbox.GrantSendOnBehalfTo -match "$($adUser.Name)"){
                $mailBoxSendOnBehalf = [PsObject]::new()

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
                        $mailBoxSendOnBehalf = [PsObject]::new()

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

    Write-Information -Message "Mailboxes which user has Send On Behalf permissions to: $($mailboxesSendOnBehalf.Count)"
    
    if($mailboxesSendOnBehalf.Count -gt 0){
        foreach($entry in $mailboxesSendOnBehalf){
            $null = $allMailboxesWithPermission.Add($entry)
        }
    }

    # Create temp csv file
    $currentDate = (Get-Date).ToString("yyyy_MM_dd_HHmmss")

    if(!(Test-Path -Path $tempFileLocation -PathType Container)){
        $newPath = New-Item $tempFileLocation -ItemType Directory -Force -Confirm:$false
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
            $mailCredentials = [System.Management.Automation.PSCredential]::new($mailSmtpUsername,$mailSecurePassword)
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
    
        Write-Information -Message "Successfully sent mail to [$mailTo]"   
    }catch{
        Write-Warning -Message $aduser
        Write-Error -Message "Error sending mail to [$mailTo]: $_"
    }


    # Clean up temp csv file
    Remove-Item -Path $fileName -Force -Confirm:$false
} catch {
    Write-Error -Message "Error gathering mailbox permissions for [$UserPrincipalName]. Error: $_"

    $Log = @{
        Action            = "MoveAccount" # optional. ENUM (undefined = default) 
        System            = "ExchangeOnline" # optional (free format text) 
        Message           = "Error gathering mailbox permissions for [$UserPrincipalName]" # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = ""# optional (free format text) 
        TargetIdentifier  = "" # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log

} finally {
    Write-Information -Message "Disconnecting from Office 365.."
    $exchangeSessionEnd = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false -ErrorAction Stop
    Write-Information -Message "Successfully disconnected from Office 365" 
}