$ScriptInfo = @"
================================================================================
Audit-OutlookProfile.ps1 | v1.0
by Roman Zarka | github.com/zarkatech
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Initialize script environment

$UserName = $env:USERNAME; $ComputerName = $env:COMPUTERNAME
$AuditFile = "$env:USERPROFILE\Desktop\$UserName" + "_OutlookAudit.csv"
"`"UserName`",`"ComputerName`",`"ProfileName`",`"ProfileEmail`",`"StoreName`",`"StoreType`",`"StorePath`"" | Out-File $AuditFile -Encoding ascii

# --- Audit Outlook profile

Write-Host "INFO: Initialize Outlook COM..." -ForegroundColor Cyan
$Outlook = New-Object -ComObject Outlook.Application
$NameSpace = $Outlook.getNameSpace("MAPI")
Write-Host "INFO: Audit Outlook profile..." -ForegroundColor Cyan
$ProfileName = $NameSpace.CurrentProfileName
$Account = ($NameSpace.Accounts | Select DisplayName,SmtpAddress)
$Stores = ($NameSpace.Stores | Select DisplayName,ExchangeStoreType,FilePath)
ForEach ($Store in $Stores) { 
    $StoreType = $Store.ExchangeStoreType
    If ($Store.ExchangeStoreType -eq "0") { $StoreType = "PrimaryMailbox" }
    If ($Store.ExchangeStoreType -eq "1") { $StoreType = "DelegateMailbox" }
    If ($Store.ExchangeStoreType -eq "2") { $StoreType = "PublicFolder" }
    If ($Store.ExchangeStoreType -eq "3") { $StoreType = "NonExchange" }
    If ($Store.ExchangeStoreType -eq "4") { $StoreType = "AdditionalMailbox" }
    "`"$UserName`",`"$ComputerName`",`"$ProfileName`",`"$($Account.SmtpAddress)`",`"$($Store.DisplayName)`",`"$StoreType`",`"$($Store.FilePath)`"" | Out-File $AuditFile -Append
}

# --- Script end

$Outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
Remove-Variable Outlook
Write-Host "SUCCESS: Script Complete." -ForegroundColor Green
