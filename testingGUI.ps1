
function AddPrinter { 
  # ADDING PRINTER LOGIC GOES HERE
  Write-Host "Connect Sessions was clicked"
}
function AddPrinter2 { 
  # ADDING PRINTER LOGIC GOES HERE
  Write-Host "Disonnect Sessions was clicked"
}

function AddPrinter3 { 
  # ADDING PRINTER LOGIC GOES HERE
  $outputext = $textBox.Text
  Write-Host "Remove user: $outputext"
}

Function ConnectSessions(){
#Connect
Write-Host "Connecting sessions"
Connect-MicrosoftTeams
Connect-ExchangeOnline
Connect-Graph -Scopes User.ReadWrite.All, Organization.Read.All
}

Function DisconnectSessions(){
Write-Host "Disconnecting sessions"
Disconnect-MicrosoftTeams
Disconnect-ExchangeOnline
}

Function RemoveUser(){
$RemoveUserBtn.BackColor         = "#7d9242"

$leaveruser = $textBox.Text
$leaverAD = Get-ADUser -Identity $leaveruser -Properties GivenName, SamAccountName, EmailAddress | select-object GivenName, SamAccountName, EmailAddress
$leaveremail = $leaverAD.EmailAddress
Write-Output $leaveremail

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

$RemoveUserBtn.BackColor         = "#a4ba67"
}

ConnectSessions

# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
# Create a new form
$LocalPrinterForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$LocalPrinterForm.ClientSize         = '500,300'
$LocalPrinterForm.text               = "Offboarding  - PowerShell GUI Testing"
$LocalPrinterForm.BackColor          = "#ffffff"
# Display the form




# Create a Title for our form. We will use a label for it.
$Title                           = New-Object system.Windows.Forms.Label
# The content of the label
$Title.text                      = "Offboarding script V0.1"
# Make sure the label is sized the height and length of the content
$Title.AutoSize                  = $true
# Define the minial width and height (not nessary with autosize true)
$Title.width                     = 25
$Title.height                    = 10
# Position the element
$Title.location                  = New-Object System.Drawing.Point(20,20)
# Define the font type and size
$Title.Font                      = 'Microsoft Sans Serif,13'
# Other elemtents
$Description                     = New-Object system.Windows.Forms.Label
$Description.text                = "This script will remove a user's phone number, set the mailbox status to shared mailbox, remove all licenses and disable AD account"
$Description.AutoSize            = $false
$Description.width               = 450
$Description.height              = 50
$Description.location            = New-Object System.Drawing.Point(20,50)
$Description.Font                = 'Microsoft Sans Serif,10'


$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20,100)
$textBox.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($textBox)

$RemoveUserBtn                   = New-Object system.Windows.Forms.Button
$RemoveUserBtn.BackColor         = "#a4ba67"
$RemoveUserBtn.text              = "Remove user"
$RemoveUserBtn.width             = 150
$RemoveUserBtn.height            = 30
$RemoveUserBtn.location          = New-Object System.Drawing.Point(180,100)
$RemoveUserBtn.Font              = 'Microsoft Sans Serif,10'
$RemoveUserBtn.ForeColor         = "#ffffff"
$RemoveUserBtn.Add_Click({ RemoveUser })



$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Quit"
$cancelBtn.width                 = 90
$cancelBtn.height                = 40
$cancelBtn.location              = New-Object System.Drawing.Point(380,250)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$LocalPrinterForm.CancelButton   = $cancelBtn
$LocalPrinterForm.Controls.Add($cancelBtn)

# ADD OTHER ELEMENTS ABOVE THIS LINE
# Add the elements to the form
$LocalPrinterForm.controls.AddRange(@($Title,$Description,$textBox,$RemoveUserBtn))
# THIS SHOULD BE AT THE END OF YOUR SCRIPT FOR NOW
# Display the form
$result = $LocalPrinterForm.ShowDialog()
if ($result –eq [System.Windows.Forms.DialogResult]::Cancel)
{
    DisconnectSessions
    write-output 'User pressed quit'
}



