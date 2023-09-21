cd 'C:\Users\Administrator\Desktop\XDC-PowershellTools\BatchImport users\Configure O365 Environment'


## Opsummering ##
# Scriptet loader først CSV Filerne, og arrangerer dataene i objekter, der kan loopes igennem.
# Hvert objekt, bliver alt efter hvad der er slået til, processeret med metoder, der udfører forskellige handlinger, med data fra objekterne.
#
## Lidt om CSV formatet. Vi bruger ';' som delimiter/separator. Dvs. hvis du bruger kommaer, fejler det. Årsagen ligger i, at vi tager imod komma-separeret input i CSV filerne.
#  Encoding bør være UTF-8.
# 
# Som default, er debug aktiveret. Dette medfører et væld af ekstra information. Debug kan dog slåes fra. Mest lavet til udvikleren.
#
# >>>>>> Når UseTestData er aktiveret, slås der automatisk over på test-data CSV-filerne.<<<<<<<
#        Disse ligger i mappen "Test data", i løsningens rod-mappe.
#
#
## Layout ##
# Til at starte med, dot-sources klasser og funktioner. Vi gør det på denne måde, for at holde dette "main-script" så "clean" som muligt.
# 
# Så kommer initiering/indstilling af config-objektet, scriptet navigerer efter.
# Herefter initialisering af relevanter objekter.
# Så kommer indlæsning af CSV filer, samt organisering af objekterne i lister.
# Og til sidst "Worker loops". Det er her at funktionerne skydes af, med data fra objekterne.
# Hvert worker loop har en kommentar, der beskriver hvad loopet udfører hvis det er enabled.

## Load helpers
.'.\Helperscripts\Functions.ps1'
.'.\Helperscripts\Classes.ps1'



## Configure
#$Config.CsvPath              = '..\365brugere_copy.csv'
$Config = [ConfigOptions]::new()
$Config.Debug                           = $true
$Config.UseTestData                     = $true
$Config.UserCsvPath                     = '..\Kunde Data\Users\Users.csv'
$Config.ContactCsvPath                  = '..\Kunde Data\Contacts\Contacts.csv'
$Config.DistributionGroupCsvPath        = '..\Kunde Data\Distribution groups\DistributionGroups2.csv'
$Config.OutputfilePath                  = '.\currentRun.csv'

 ## Functions - Enable/Disable
$Config.AssignLicenses                  = $false # Not created
$Config.AssignProxyAddresses            = $true  # Ready
$Config.CreateUsers                     = $True  # Ready
$Config.CreateContacts                  = $true  # WIP
$Config.CreateDistributionGroups        = $true  # WIP
$Config.ResetPasswords                  = $false # Ready
$Config.SetMailboxDanishTimeAndLanguage = $false # Tilret så vi laver en Get-Mailbox *, og piper til set-mailbox
$Config.SetMailboxStartWeekConfig       = $false
$Config.HandleModules                   = $false #Ready

## Load test files if DEBUG is activated.
if($Config.UseTestData -eq $true){
    $Config.UserCsvPath                     = '..\Test Data\Users\Users.csv'
    $Config.ContactCsvPath                  = '..\Test Data\Contacts\Contacts.csv'
    $Config.DistributionGroupCsvPath        = '..\Test Data\Distribution groups\DistributionGroups.csv'
}

## Clear screen
Clear-Host


## Test if USER CSV files exists and is parseable
Write-host "<<<<<<<<<<<<<<<<<<<<<<< Testing file paths #######################" -ForegroundColor Yellow
if (Test-Path $Config.UserCsvPath) {
if($Config.Debug -eq $true){    Indicate-OK; Write-Host "User CSV File exists" -ForegroundColor Cyan }
    # Test if the CSV file is parseable
    if($test = Import-csv -Path $Config.UserCsvPath -Delimiter ';')
        { if($Config.Debug -eq $true){ Indicate-OK ; Write-Host "User CSV appears to be readable." -ForegroundColor Cyan} }
    Else{ if($Config.Debug -eq $true){ Indicate-ERROR ; Write-Host 'User CSV File does not appear parseable. Remember we use ";" as delimiter'} }
    } 
else {
        { if($Config.Debug -eq $true){ Indicate-ERROR ; Write-Host "User CSV File doesn't exist"} }
}


## Test if Contacts CSV files exists and is parseable
if (Test-Path $Config.ContactCsvPath) {
    if($Config.Debug -eq $true){Indicate-OK; Write-Host "Contacts CSV File exists" -ForegroundColor Cyan}
    # Test if the CSV file is parseable
    if($test = Import-csv -Path $Config.ContactCsvPath -Delimiter ';')
        { if($Config.Debug -eq $true){ Indicate-OK    ; Write-Host "Contacts CSV File appears to be readable." -ForegroundColor Cyan} }
    else{ if($Config.Debug -eq $true){ Indicate-ERROR ; Write-Host 'Contacts CSV File does not appear parseable. Remember we use ";" as delimiter'}}
    } 
else {
        { if($Config.Debug -eq $true){ Indicate-ERROR ; Write-Host "Contacts CSV File doesn't exist"} }
}


## Test if Distribution group CSV files exists and is parseable
if (Test-Path $Config.DistributionGroupCsvPath) {
    if($Config.Debug -eq $true){ Indicate-OK; Write-Host "Distribution group CSV File exists" -ForegroundColor Cyan }
    # Test if the CSV file is parseable
    if($test = Import-csv -Path $Config.UserCsvPath -Delimiter ';')
        { if($Config.Debug -eq $true){ Indicate-OK    ; Write-Host "Distribution group CSV File appears to be readable." -ForegroundColor Cyan} }
    else{ if($Config.Debug -eq $true){ Indicate-ERROR ; Write-Host 'Distribution group CSV File does not appear parseable. Remember we use ";" as delimiter'} }
} 
else {
         if($Config.Debug -eq $true){ Indicate-ERROR ; Write-Host "Distribution group CSV File doesn't exist" }
}

if($Config.Debug -eq $true){ Write-host "################## Done Testing file paths >>>>>>>>>>>>>>>>>>>>>>> $([System.Environment]::NewLine)" -ForegroundColor Yellow }


## Config headsup
Write-host "####################### Tasks that will be completed ##############$([System.Environment]::NewLine)" -ForegroundColor Yellow

Write-host " The following tasks will be completed:$([System.Environment]::NewLine)"
$config.Print_Config() ; Write-host "$([System.Environment]::NewLine)"

Write-host "####################### Finished fetching tasks ####################### $([System.Environment]::NewLine)" -ForegroundColor Yellow


## Initialize variables
$usersArray = @()             # To store the CSV data
$UserList = [UserList]::new() # To store the working objects
$ContactList = [ContactList]::new()
$DistributionGroupList = [DistributionGroupList]::new()

## Load user-data from CSV
Write-host "<<<<<<<<<<<<<<<<<<<<<<< Loading CSV Data ####################### $([System.Environment]::NewLine)" -ForegroundColor Yellow
If($Config.AssignLicenses -or $Config.AssignProxyAddresses -or $Config.CreateUsers -or $Config.ResetPasswords -or $Config.SetMailboxStartWeekConfig -or $Config.SetMailboxDanishTimeAndLanguage ){
    try{
        
        $userData = Import-Csv -Path $Config.UserCsvPath -Delimiter ';' -Encoding UTF8
        if($Config.Debug -eq $true){Indicate-OK ; Write-host "Loaded User CSV data..." -ForegroundColor Cyan}
    }
    catch{
        Indicate-Error; write-host "Could not parse User CSV"
    }
}

## Load Contacts-data from CSV

if($Config.CreateContacts){
    try{
        $ContactsData = Import-Csv -Path $Config.ContactCsvPath -Delimiter ';' -Encoding UTF8
        if($Config.Debug -eq $true){Indicate-OK ; Write-host "Loaded Contacts CSV data..." -ForegroundColor Cyan}
    }
    catch{
        Indicate-Error; write-host "Could not parse Contacts CSV"
    }
    
}

## Load Distribution group-data from CSV
if($Config.CreateDistributionGroups){
    try{
        $DistributionGroupData = Import-Csv -Path $Config.DistributionGroupCsvPath -Delimiter ';' -Encoding UTF8
        if($Config.Debug -eq $true){Indicate-OK ; Write-host "Loaded Distribution group CSV data..." -ForegroundColor Cyan}
        }
    catch{
        Indicate-Error; write-host "Could not parse Distribution Groups CSV"
    }
}
if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)################## Done Loading CSV Data >>>>>>>>>>>>>>>>>>>>>>> $([System.Environment]::NewLine)" -ForegroundColor Yellow }


## Load required modules
if($Config.HandleModules){
    if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)<<<<<<<<<<<<<< Installing / Updating modules ####################### $([System.Environment]::NewLine)" -ForegroundColor Yellow }
    Prepare-Modules($Config)
    if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)########### Done Installing / Updating modules >>>>>>>>>>>>>>>>>>>>>>> $([System.Environment]::NewLine)" -ForegroundColor Yellow }
}






## Load and Organize user data
If($Config.AssignLicenses -or $Config.AssignProxyAddresses -or $Config.CreateUsers -or $Config.ResetPasswords -or $Config.SetMailboxStartWeekConfig -or $Config.SetMailboxDanishTimeAndLanguage ){
    if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)<<<<<<<<<<<<<<<<< Processing: User Data - Shaping Array ##################### $([System.Environment]::NewLine)" -ForegroundColor Yellow }
    $InitialPosition = $host.UI.RawUI.CursorPosition
    if($Config.Debug -eq $true){Write-host "Now processing: " -NoNewline}
    # Record the current position of the cursor 
    $OriginalPosition = $host.UI.RawUI.CursorPosition

    $userData | ForEach-Object -Begin {
    } -Process {
          
      # Set the position of the cursor back to where it was, before the loop started
      [Console]::SetCursorPosition($originalPosition.X,$originalPosition.Y)
       if($Config.Debug -eq $true){Write-host "$($_.'Logon / E-mail')                            "}

        $PasswordProfile = @{
            Password                             = Generate-RandomPassword 10
            ForceChangePasswordNextSignIn        = $true
            ForceChangePasswordNextSignInWithMfa = $true
        }

        $user = [NewUser]::new()
        $user.DisplayName = $_.'Fulde navn'
        $user.FirstName = $_.Fornavn
        $user.LastName = $_.Efternavn
        $user.MailNickName = ($_.'Logon / E-mail'.Split('@')[0])
        $user.UserPrincipalName = $_.'Logon / E-mail'
        $user.PasswordProfile = $PasswordProfile
        if($user.ProxyAddresses -ne $null){
            $user.ProxyAddresses = Cleanup-AlternateEmails($_.'Alternate email address'.Split(','))
        }

        if($_.'OFFICE PAKKE' -contains 'JA ')         { $user.OfficePakke = $true}
                                                 else { $user.OfficePakke     = $false  }

        if($_.'ADGANG ADMINISTRATION' -contains 'JA') { $user.AdgangAdministration = $true }  
                                                 else { $user.AdgangAdministration = $false  }
    
        if($_.'ADGANG ALLE' -contains  'JA')          { $user.AdgangAlle = $true  }           
                                                 else { $user.AdgangAlle      = $false  }

        if($_.'ADGANG DIREKTION' -contains 'JA')      { $user.AdgangDirektion = $true }       
                                                 else { $user.AdgangDirektion = $false  }

        if($_.'ADGANG HR' -contains 'JA')             { $user.AdgangHR = $true }              
                                                 else { $user.AdgangHR        = $false  }

        if($_.'ADGANG LAGER' -contains 'JA')          { $user.AdgangLager = $true }           
                                                 else { $user.AdgangLager     = $false  }

        if($_.'ADGANG SALG' -contains 'JA')           { $user.AdgangSalg = $true }            
                                                 else { $user.AdgangSalg      = $false  }
    
        if($_.'ADGANG VÆRKSTED' -eq 'JA')             { $user.AdgangVaerksted = $true }       
                                                 else { $user.AdgangVaerksted = $false  }
        
        if($_.'Logon / E-mail' -ne $null)             { $null = $UserList.Collection.Add($user)}
                                                  else{ write-host "Skipping user: $($_.'Fulde navn')" }
                                                  
    
    } -End{
    [Console]::SetCursorPosition($InitialPosition.X,$InitialPosition.Y)
    Indicate-OK ; Write-Host "$($userlist.Collection.Count) Users parsed and ready to go!" -ForegroundColor Cyan
    } ## UserData is now parsed, and we can loop through them
  
    if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)##################### Finished Shaping User Data Array >>>>>>>>>>>>>>>>  $([System.Environment]::NewLine)" -ForegroundColor Yellow }
}

## Load and organize Contacts
if($Config.CreateContacts){
    if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)<<<<<<<<<<<<<<<<< Processing: Contacts Data - Shaping Array ##################### $([System.Environment]::NewLine)" -ForegroundColor Yellow
        $InitialPosition = $host.UI.RawUI.CursorPosition
        Write-host "Now processing: " -NoNewline
        # Record the current position of the cursor 
        $WorkingAreaPosition = $host.UI.RawUI.CursorPosition
    }
        $ContactsData | % {
            if($Config.Debug -eq $true){
                [Console]::SetCursorPosition($WorkingAreaPosition.X,$WorkingAreaPosition.Y)
                Write-host "$($_.SmtpAddress)                  "
            }
            $Contact = [Contact]::new()
            $Contact.DisplayName = $_.DisplayName
            $Contact.SmtpAddress = $_.SmtpAddress
            $ContactList.Collection.Add($Contact) | Out-null
            
        }
    [Console]::SetCursorPosition($InitialPosition.X,$InitialPosition.Y)
    Indicate-OK ; Write-host "$($ContactList.Collection.Count) contacts parsed and ready to go!" -ForegroundColor Cyan
    if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)##################### Finished Shaping Contacts Data Array >>>>>>>>>>>>>>>>  $([System.Environment]::NewLine) " -ForegroundColor Yellow }
}





## Load and organize Distribution groups
if($Config.CreateDistributionGroups){
 if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)<<<<<<<<<<<<<<<<< Processing: Distribution Group Data - Shaping Array ##################### $([System.Environment]::NewLine)" -ForegroundColor Yellow
        $InitialPosition = $host.UI.RawUI.CursorPosition
        Write-host "Now processing: " -NoNewline
        # Record the current position of the cursor 
        $WorkingAreaPosition = $host.UI.RawUI.CursorPosition
    }
    $UniqueGroups = $DistributionGroupData | Select ParentGroup -Unique 

    $UniqueGroups | % {
        $DistributionGroup = [DistributionGroup]::new()
        $DistributionGroup.GroupName = $_.ParentGroup
        $DistributionGroupList.Collection.Add($DistributionGroup) | Out-Null
    }


$DistributionGroupList.Collection | % {
    if($Config.Debug -eq $true){
        [Console]::SetCursorPosition($WorkingAreaPosition.X,$WorkingAreaPosition.Y)
        Write-host "$($_.GroupName)                            "
        
            }
    $GroupName = $_.GroupName
    $GroupMembers = $DistributionGroupData | ? { $_.ParentGroup -eq "$GroupName"}
 
    Foreach($GroupMember in $GroupMembers){
        $DistributionGroupMember = [DistributionGroupMember]::new()
        $DistributionGroupMember.MemberName = $GroupMember.Name
        $DistributionGroupMember.SmtpAddress = $GroupMember.Smtp
        $_.Collection.Add($DistributionGroupMember) | Out-Null
    }
}

    [Console]::SetCursorPosition($InitialPosition.X,$InitialPosition.Y)
    Indicate-OK ; Write-host "$($DistributionGroupList.Collection.Count) Distribution groups parsed and ready to go!                     " -ForegroundColor Cyan
    if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)##################### Finished Shaping Distribution Group Array >>>>>>>>>>>>>>>>  $([System.Environment]::NewLine) " -ForegroundColor Yellow }

}


## In preparation for our worker-loops, we connect the required assets for the operation(s)
 if($Config.CreateUsers -or $Config.AssignLicenses){
     if($Config.Debug -eq $true){
        Write-host "$([System.Environment]::NewLine)<<<<<<<<<<<<<<<<< Requesting User input: Graph Credentials ##################### $([System.Environment]::NewLine)" -ForegroundColor Yellow
    }
        $Scope = 'User.ReadWrite.All'
        try{
            Connect-MgGraph -Scopes $Scope -NoWelcome
            if($Config.Debug -eq $true){
                VerifyConnected
                Write-host "$([System.Environment]::NewLine)##################### Recieved user input: Graph Credentials >>>>>>>>>>>>>>>>>$([System.Environment]::NewLine)" -ForegroundColor Yellow
            }
        }
        Catch{
            Indicate-Error ; Write-Host "Something went wrong. Did you close the authentication window?"
        }
            
}
    



if($Config.Debug -eq $true){Write-host "$([System.Environment]::NewLine)<<<<<<<<<<<<<<<<< Requesting User input: Graph Credentials ##################### $([System.Environment]::NewLine)" -ForegroundColor Yellow}    

    if($Config.ResetPasswords){
        $Scope = 'User.ReadWrite.All,UserAuthenticationMethod.ReadWrite.All'
        Connect-MgGraph -Scopes $Scope -NoWelcome
        VerifyConnected
    }

    if($Config.AssignProxyAddresses -or $Config.CreateContacts -or $Config.CreateDistributionGroups -or $Config.SetMailboxStartWeekConfig -or $Config.SetMailboxDanishTimeAndLanguage)
    {
        Connect-ExchangeOnline -ShowBanner:$false -TrackPerformance:$true -SkipLoadingCmdletHelp
        
    }

Read-host -Prompt "Paused..."




## Worker-Loop: Users
$UserList.Collection | ForEach-Object -Begin{

} -Process {
    ##################################
    # Create Users
    # Status: Working
    ##################################
    if($Config.CreateUsers){
       $_.createUserResult = Create-AzureUser($_)
    }
    ################################## 
    # Assign Proxy addresses
    # Status: Prototype done. Needs testing
    # Requires: Assigned exchange license
    ##################################
    if($Config.AssignProxyAddresses){
        $mailArray = @()
        foreach($email in $_.ProxyAddresses){
            if($email -ne ''){
                $mailArray += $email
            }
        }
        if($mailArray -ne $null){
            $stringbuilder = $stringbuilder + ($mailArray -join  ',')
            $ProxyAddressResult = Append-MailboxAliases -userid $_.UserPrincipalName -mailboxAliases $_.proxyAddresses
            
            if($ProxyAddressResult -eq $true){}
            else{}

            # Get-Mailbox -Identity '*' -ResultSize Unlimited | select UserPrincipalName, EmailAddresses
        }
        
    }

    ##################################
    # Not started
    # Assign Licenses
    ##################################
            if($Config.AssignLicenses){}
    ##################################
    # Reset passwords
    # Working
    ##################################
    if($Config.ResetPasswords){
    write-host "Resetting password for: $($_.UserPrincipalName)" -ForegroundColor Cyan
       $_.resetPasswordResult = Reset-AzureUserPassword -userid $_.UserPrincipalName -Password $_.PasswordProfile["Password"]
    }

} -End {
   
 }


## Worker-Loop: Contacts
if($Config.CreateContacts){
    $ContactList.Collection | % {
        Create-Contacts -name $_.DisplayName -address $_.SmtpAddress
    }
}

## Worker-Loop: Distribution groups
if($Config.CreateDistributionGroups){
    $DistributionGroupList.Collection | % {
        Create-DistributionGroup -GroupName $_.GroupName
    }
    $DistributionGroupList.Collection | % {
        $GroupName = $_.GroupName
        write-host "Processing:: Group name: $GroupName" -ForegroundColor Magenta
            Foreach($Member in $_.Collection){
                Write-host "Adding Member: $($member.MemberName) - Address: $($Member.SmtpAddress)"
                Add-DistributionGroupMember -Identity "$GroupName" -Member "$($Member.SmtpAddress)"
            }
    }
}

## Worker-Loop: Mailbox config - Language and Time
if($Config.$Config.SetMailboxDanishTimeAndLanguage){
    $UserList.Collection | % {
        ConfigureMailbox-DanishLanguageAndTimeFormat -User $_.UserPrincipalName
    }
}
## Worker-Loop: Mailbox config - Week days config
if($Config.SetMailboxStartWeekConfig){
    $UserList.Collection | % {
        ConfigureMailbox-DanishWeekStartConfiguration -User $_.UserPrincipalName
    }
}

## Disconnect the relevant ressources
if($Config.AssignLicenses -or $Config.CreateUsers -or $Config.ResetPasswords){
    Disconnect-MgGraph    
}
if($Config.AssignProxyAddresses -or $Config.CreateContacts -or $Config.CreateDistributionGroups){
    Disconnect-ExchangeOnline -Confirm:$false
}


OutputResults -ResultSet $UserList -createUsers $Config.createUsers -AssignLicenses $Config.AssignLicenses -AssignProxyAddresses $Config.AssignProxyAddresses -ResetPasswords $Config.ResetPasswords -outputpath $Config.OutputfilePath
