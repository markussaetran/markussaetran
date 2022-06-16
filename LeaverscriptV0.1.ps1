#### BACKEND FUNCTIONS ###

#Modules
#Install-Module -Name Microsoft.Graph
#Install-Module -Name ExchangeOnlineManagement

Function Get-FileName($InitialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.FileName
}


Function Get-FileNameEx($InitialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.filter = "Excel Workbook (*.xlsx) | *.xlsx"
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.FileName
}

###########
### USER FUNCTIONS ###
$leaveruser = Read-Host "Username of User to Remove"
$leaverAD = Get-ADUser -Identity $leaveruser -Properties GivenName, SamAccountName, EmailAddress | select-object GivenName, SamAccountName, EmailAddress
$leaveremail = $leaverAD.EmailAddress
Write-Output $leaveremail


Function ConnectSessions(){
#Connect
Connect-MicrosoftTeams
Connect-ExchangeOnline
Connect-Graph -Scopes User.ReadWrite.All, Organization.Read.All
}

Function RemoveUser(){



#Remove Phone
Set-CsUser -Identity $leaveremail -OnPremLineURI $null -EnterpriseVoiceEnabled $false
Write-Host "Removing phone"


#SetSharedMailbox
#Remember to connect exchange
Set-Mailbox -Identity $leaveruser -Type Shared
Write-Host "Setting shared mailbox"
sleep (10)
Write-Host "Connecting"
sleep(10)
Write-Host "Making sure it is set before removing licenses"
sleep(10)
write-host "In the early stages of the script. Check if mailbox is set correctly in admin centre"
pause

#REMOVE ALL LICENSES
$licenseSKUs = Get-MgUserLicenseDetail -UserId $leaveremail | Sort-Object SkuID
$SkuID = $licenseSKUs.SkuId

foreach ($i in $SkuID) {
    #Write-Output $i, "Text"
    Set-MgUserLicense -UserId $leaveremail -RemoveLicenses @($i) -AddLicenses @{}
}

Write-Host "All licenses removed"


#Disable AD
Disable-ADAccount -Identity $LeaverAD.SamAccountName
Write-Host "AD account disabled"

#Move to expired OU
#Get-ADUser -Identity $leaveruser | Move-ADObject -TargetPath "OU=XX-O365-Expired,OU=O365-Users,DC=atlantic,DC=local"
}

#Disconnect sessions
Function DisconnectSessions(){
Write-Host "Disconnecting sessions"
Disconnect-MicrosoftTeams
Disconnect-ExchangeOnline
}