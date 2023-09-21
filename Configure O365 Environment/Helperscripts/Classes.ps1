
Class ConfigOptions{
    [bool]$debug = $true
    [bool]$UseTestData = $true
    [string]$UserCsvPath = ""
    [string]$ContactCsvPath
    [string]$DistributionGroupCsvPath
    [string]$OutputfilePath = ".\currentRun.csv"
    [bool]$CreateUsers = $false
    [bool]$AssignProxyAddresses = $false
    [bool]$AssignLicenses = $false
    [bool]$ResetPasswords = $false
    [bool]$CreateContacts = $false
    [bool]$CreateDistributionGroups = $false
    [bool]$SetMailboxDanishTimeAndLanguage = $false
    [bool]$SetMailboxStartWeekConfig = $false
    [bool]$HandleModules = $false

    [void]Print_Config(){
       If($this.AssignLicenses)      { Indicate-OK ; Write-host "Will assign licenses:        $($this.Translate_Option($this.AssignLicenses))" -ForegroundColor Cyan }
       if($this.CreateUsers)         { Indicate-OK ; Write-host "Will create users:           $($this.Translate_Option($this.CreateUsers))" -ForegroundColor Cyan }
       if($this.AssignProxyAddresses){ Indicate-OK ; Write-host "Will assign proxy addresses: $($this.Translate_Option($this.AssignProxyAddresses))" -ForegroundColor Cyan }
       if($this.ResetPasswords)      { Indicate-OK ; Write-host "Will reset user passwords:   $($this.Translate_Option($this.ResetPasswords))" -ForegroundColor Cyan }
       if($this.CreateContacts)      { Indicate-OK ; Write-host "Will create contacts:        $($this.Translate_Option($this.CreateContacts))" -ForegroundColor Cyan }
    }

    [string]Translate_Option([bool]$Option){
        if($Option){return "Yes"} 
        else{ return "No"}
    }

}


Class NewUser {
    [object]$PasswordProfile   = $PasswordProfile
    [bool]$AccountEnabled      = $true
    [string]$FirstName         = ""
    [string]$LastName          = ""
    [string]$DisplayName       = ""
    [string]$MailNickName      = ""
    [string[]]$ProxyAddresses  = ""
    [string]$UserPrincipalName = ""
    [bool]$OfficePakke         = $false
    [bool]$AdgangAdministration= $false
    [bool]$AdgangAlle          = $false
    [bool]$AdgangDirektion     = $false
    [bool]$AdgangHR            = $false
    [bool]$AdgangLager         = $false
    [bool]$AdgangSalg          = $false
    [bool]$AdgangVaerksted     = $false
    [string]$userid            = ""
    [string]$company           = "Gunner Due A/S"
    [bool]$resetPasswordResult = $false
    [bool]$createUserResult    = $false
    [bool]$AssignLicensesResult= $false
    [System.Collections.ArrayList]$AssignProxyAddressesSuccessResult = @()
    [System.Collections.ArrayList]$AssignProxyAddressesErrorResult = @()
}

Class UserList {
    [System.Collections.ArrayList]$Collection = @()
}

Class Contact {
    [string]$DisplayName = ""
    [string]$SmtpAddress = ""
}

Class ContactList {
    [System.Collections.ArrayList]$Collection = @()
}

Class DistributionGroup {
    [string]$GroupName
    [System.Collections.ArrayList]$Collection = @()
}

Class DistributionGroupList{
    [System.Collections.ArrayList]$Collection = @()
}

Class DistributionGroupMember{
    [string]$MemberName
    [string]$SmtpAddress
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