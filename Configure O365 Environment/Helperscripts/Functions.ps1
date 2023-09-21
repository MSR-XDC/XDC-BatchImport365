cd 'C:\Users\Administrator\Desktop\XDC-PowershellTools\BatchImport users\Configure O365 Environment'

.'.\Helperscripts\Classes.ps1'


function VerifyConnected() {
    $Details = Get-MgContext
    if($Details -ne $null){
    $Scopes = $Details | Select-Object -ExpandProperty Scopes
    $Scopes = $Scopes -join ","
    $OrgName = (Get-MgOrganization).DisplayName

$ConnectionDetails = @"
Microsoft Graph current session details:
---------------------------------------
Tenant Id = $($Details.TenantId)
Client Id = $($Details.ClientId)
Org name = $OrgName
App Name = $($Details.AppName)
Account = $($Details.Account)
Scopes = $Scopes
"@
    write-host $ConnectionDetails -ForegroundColor Cyan
    Indicate-OK ; Write-host "Successfully connected to Graph!" -ForegroundColor Cyan
    }
    else{
        Indicate-Error ; Write-host "Failed connecting to Graph..."
    }

}



function Reset-AzureUserPassword($userid, $password){
    $method = "28c10230-6103-485e-b985-444c60001490"
    try{
        $error.Clear()
        Reset-MgUserAuthenticationMethodPassword -userId $userid -AuthenticationMethodId $method -NewPassword $password
        write-host "Reset password: $userid , to $password"
        write-host  $LASTEXITCODE -ForegroundColor Magenta
        if($error[0].Exception.Message -match 'The specified user could not be found.'){
            write-host 'Error detected' -ForegroundColor Yellow
            $error.Clear()
            return($false)
        }
        else{
         $error.Clear()
         return($true)
        }
    }
    catch{
        $error.Clear()
        return $false
    }

    
}



function Generate-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [int] $length
    )
 
    $charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.ToCharArray()
 
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
  
    $rng.GetBytes($bytes)
  
    $result = New-Object char[]($length)
  
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i]%$charSet.Length]
    }
    $result += "!"
    return -join $result
}



Function Create-AzureUser([NewUser]$currentUser){
    
        # Create Azure AD user
       try{
       $response = New-MgUser -DisplayName $currentUser.DisplayName`
                   -MailNickname $currentUser.MailNickName`
                   -UserPrincipalName $currentUser.UserPrincipalName`
                   -PasswordProfile $currentUser.PasswordProfile`
                   -Company $currentUser.company`
                   -AccountEnabled
          
          write-host "Created user: " $currentUser.DisplayName -ForegroundColor Green
          return $true
          }
          catch{
          Write-Host "User: $($currentUser.MailNickName) already exists!" -ForegroundColor Yellow
          return $false
          }
}





Function Append-MailboxAliases([string]$userid, [string[]]$mailboxAliases){
    try{
        try{
            write-host "Looking up mailbox for: $($userid)"
            $mailbox= Get-Recipient $userid | Get-Mailbox
            if($mailbox -ne $null){write-host "Found mailbox!"}
        }
        catch{
            write-host "Could not look up mailbox belonging to: $($userid)"
            return $false
        }

        foreach($mail in $mailboxAliases){
        try{
            Set-Mailbox -Identity $mailbox -EmailAddresses @{Add="smtp:$mail"} -WarningVariable $Warning
            write-host "Successfully added $mail to $($userid)'s mailbox"
            }
            catch{
            write-host "Could not add $mail to $($userid)'s mailbox. Does your account permissions work? Is it a AzADSync'ed account?"

            }
        }        
        return $true
    }
    catch{
        return $false
    }
}



Function Cleanup-AlternateEmails([string[]]$emails){
    $CleanedMails = @()
    foreach($address in $emails)
    {
        $CleanedMails += $address.TrimStart().TrimEnd()
    }
return $CleanedMails
}





Function OutputResults($ResultSet, $Config, $outputpath){
   
    write-host "Outputting"
    write-host "$resultset"
    write-host "$Config"
    write-host "$outputpath"
    Get-Type $config

    if($Config.createUsers){
    Write-host "Users successfully created:"
    $ResultSet.Collection | % {
        if($_.CreateUserResult){
            Write-host "$($_.UserPrincipalName) - Password: $($_.PasswordProfile["Password"] )"
        }
    
    }
    
    $ResultSet.Collection | % {
        Write-Host "Users that did NOT get created" -ForegroundColor Yellow
        if(!$_.CreateUserResult){
            Write-host "$($_.UserPrincipalName) - Password: $($_.PasswordProfile["Password"] )"
        }
    
        }
    }

    if($Config.AssignLicenses){
    Write-host "Users successfully assigned licenses:"
    $ResultSet.Collection | % {
        if($_.AssignLicensesResult){
            Write-host "$($_.UserPrincipalName)"
        }
    
    }
    $ResultSet.Collection | % {
        Write-Host "Users that did NOT get licenses assigned" -ForegroundColor Yellow
        if(!$_.AssignLicensesResult){
            Write-host "$($_.UserPrincipalName)"
        }

    
    }
}
 
 
    if($Config.AssignProxyAddresses){
    Write-host "Users successfully SMTP Aliases:"
    $ResultSet.Collection | % {
        if($_.AssignProxyAddressesResult){
            Write-host "$($_.UserPrincipalName)"
        }
    
    }

    $ResultSet.Collection | % {
        Write-Host "Users that did NOT get SMTP Aliases assigned" -ForegroundColor Yellow
        if(!$_.AssignProxyAddressesResult){
            Write-host "$($_.UserPrincipalName)"
        }

    
    }
 }   
    if($Config.ResetPasswords){
    Write-host "Users with successfully reset passwords:"
    $ResultSet.Collection | % {
        if($_.resetPasswordResult){
            Write-host "$($_.UserPrincipalName) .. Password: $($_.PasswordProfile["Password"] )"
        }
    
    }
    $ResultSet.Collection | % {
        Write-Host "--------------Users that did NOT get their passwords reset-------------" -ForegroundColor Yellow
        if(!$_.resetPasswordResult){
            Write-host "$($_.UserPrincipalName)"
        }

    
    }
 }

    
}


Function Prepare-Modules($Config){

$modules = 'Microsoft.Graph',
'Microsoft.Graph.Identity.DirectoryManagement',
'Microsoft.Graph.Users',
'ExchangeOnlineManagement',
'Microsoft.Graph.Authentication',
'Microsoft.Graph.Users.Actions',
'Microsoft.Graph.Identity.SignIns'

# Find those that are already installed.
    $installed = @((Get-Module $modules -ListAvailable).Name | Select-Object -Unique)

    # Infer which ones *aren't* installed.
    $notInstalled = Compare-Object $modules $installed -PassThru

    if ($notInstalled) { # At least one module is missing.

      # Prompt for installing the missing ones.
      $promptText = @"
      The following modules aren't currently installed:
  
          $notInstalled
  
      Would you like to install them now?
"@
      $choice = $host.UI.PromptForChoice('Missing modules', $promptText, ('&Yes', '&No'), 0)
  
      if ($choice -ne 0) { Write-Warning 'Will not install required modules....'}
      else {
      Write-host "Installing required modules..."
      Install-Module -Scope CurrentUser $notInstalled
      }
      # Install the missing modules now.
      
    }

 # Prompt for updating the remaining ones.
      $promptText = @"
      Would you like to update the modules that was already installed?
"@
      $choice = $host.UI.PromptForChoice('Missing modules', $promptText, ('&Yes', '&No'), 0)
      if ($choice -ne 0) { Write-Warning 'Will not update modules....'}
      else{write-host 'Updating modules...'
        $Modules | % {Update-Module -Name $_ -Force}
      }
      



    if($Config.AssignLicenses -or $Config.CreateUsers -or $Config.ResetPasswords){
        #Import-Module Microsoft.Graph
        Import-Module Microsoft.Graph.Users
    }

    if($Config.ResetPasswords){
        #Import-Module Microsoft.Graph
        #Import-Module Microsoft.Graph.Users
        Import-Module Microsoft.Graph.Authentication
        Import-Module Microsoft.Graph.Identity.SignIns
    }

    if($Config.AssignProxyAddresses){
        Import-Module ExchangeOnlineManagement
    }
}
$Config.createUsers, $Config.AssignLicenses, $config.AssignProxyAddresses, $config.ResetPasswords 

Function OutputResults($ResultSet, $createUsers, $AssignLicenses, $AssignProxyAddresses, $ResetPasswords, $outputpath){
   
    write-host "Outputting"
    write-host "createUsers: $createUsers"
    write-host "AssignLicenses: $AssignLicenses"
    write-host "AssignProxyAddresses: $AssignProxyAddresses "
    write-host "ResetPasswords: $ResetPasswords"
    write-host "$outputpath"
    

    if($createUsers){
    Write-host "Users successfully created:"
    $ResultSet.Collection | % {
        if($_.CreateUserResult){
            Write-host "$($_.UserPrincipalName) - Password: $($_.PasswordProfile["Password"] )"
        }
    
    }
    
    Write-Host "Users that did NOT get created" -ForegroundColor Yellow
    $ResultSet.Collection | % {
        
        if(!$_.CreateUserResult){
            Write-host "$($_.UserPrincipalName) - Password: $($_.PasswordProfile["Password"] )"
        }
    
        }
    }

    if($AssignLicenses){
    Write-host "Users successfully assigned licenses:"
    $ResultSet.Collection | % {
        if($_.AssignLicensesResult){
            Write-host "$($_.UserPrincipalName)"
        }
    
    }
    Write-Host "Users that did NOT get licenses assigned" -ForegroundColor Yellow
    $ResultSet.Collection | % {
        
        if(!$_.AssignLicensesResult){
            Write-host "$($_.UserPrincipalName)"
        }

    
    }
}
 
 
    if($AssignProxyAddresses){
    Write-host "Users successfully SMTP Aliases:"
    $ResultSet.Collection | % {
        if($_.AssignProxyAddressesResult){
            Write-host "$($_.UserPrincipalName)"
        }
    
    }
    Write-Host "Users that did NOT get SMTP Aliases assigned" -ForegroundColor Yellow
    $ResultSet.Collection | % {
        
        if(!$_.AssignProxyAddressesResult){
            Write-host "$($_.UserPrincipalName)"
        }

    
    }
 }   
    if($ResetPasswords){
    Write-host "Users with successfully reset passwords:"
    $ResultSet.Collection | % {
        if($_.resetPasswordResult){
            Write-host "$($_.UserPrincipalName) .. Password: $($_.PasswordProfile["Password"] )"
        }
    
    }
    Write-Host "Users that did NOT get their passwords reset" -ForegroundColor Yellow
    $ResultSet.Collection | % {
        
        if(!$_.resetPasswordResult){
            Write-host "$($_.UserPrincipalName)"
        }

    
    }
 }

    
}


Function Create-Contacts([string]$name, [string]$address){
    try{
        New-MailContact -Name $name -ExternalEmailAddress $address
    }
    catch{
        Write-host "Failed to add: $name - with address: $address"
    }
}


Function Create-DistributionGroup([string]$GroupName){
    try{
        New-DistributionGroup -Name $GroupName
        #Set-DistributionGroup -RequireSenderAuthenticationEnabled $False
        }
    catch {
        Write-host "Failed to create distribution group with name: $GroupName"
    }
}

Function Populate-DistributionGroup([string]$name, [string]$group){
    try{
        Add-DistributionGroupMember -Identity "Staff" -Member "JohnEvans@contoso.com"
    }
    catch{
        Write-host "Failed to add distribution group member: $Name - to group: $GroupName"
    }
}

Function ConfigureMailbox-DanishLanguageAndTimeFormat([string]$User){
    Set-MailboxRegionalConfiguration -Identity "$user"`
                                     -Language "DA-dk"`
                                     -DateFormat "dd-MM-yyyy"`
                                     -LocalizeDefaultFolderName
}

Function ConfigureMailbox-DanishWeekStartConfiguration([string]$User){
    Set-MailboxCalendarConfiguration -Identity "$User"`
                                     -FirstWeekOfYear "FirstFourDayWeek"`
                                     -ShowWeekNumbers $true `
                                     -TimeIncrement "FifteenMinutes"
}




Function Indicate-OK {
    Write-Host " [" -NoNewline
    Write-Host " OK " -ForegroundColor Green -NoNewline
    Write-host "] ... " -NoNewline
}
Function Indicate-Error {
    Write-Host " [" -NoNewline
    Write-Host " ERROR " -ForegroundColor Red -NoNewline
    Write-host "] ..." -NoNewline
}