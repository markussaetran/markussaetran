#Modules
#Import-Module ActiveDirectory
#Import-Module OneTimeSecret

#Account creation
#$NewUser = Read-Host "username" 
$global:setfullname = $false

#password generator
function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [int] $length
    )
    #$charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789{]+-[*=@:)}$^%;(_!&amp;#?>/|.'.ToCharArray()
    $charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
 
    $rng.GetBytes($bytes)
 
    $result = New-Object char[]($length)
 
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i]%$charSet.Length]
    }
 
    return (-join $result)
}
$pw = Get-RandomPassword 10
$securepassword = $pw | ConvertTo-SecureString -AsPlainText -Force


#Authorize OTS
Set-OTSAuthorizationToken -Username markuss@atlanticproductions.co.uk -APIKey eaa19d4e1a236f4fc6858eb8de009cafe03e6c96


#Validate username
function validateUsername(){
try {
    Get-ADUser -Identity $user
    $UserExists = $true
    Write-Host "Username is taken"
    
    #Remove comments after testing
    $ProceedCheck = Read-Host "Is this user already set up in AD? y/n"
    if ($ProceedCheck -eq "y"){
        #Enable-ADAccount -Identity $NewUser
        Write-Host "AD account enabled"
        Connect-ExchangeOnline
        #Set-Mailbox -Identity $leaveruser -Type Shared
        Write-Host "Setting mailbox to user"
        Disconnect-ExchangeOnline
    }
    if ($ProceedCheck -eq "n"){
    Write-Host "Setting username to full name"
    $global:setfullname = $true
    }
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException] {
    Write-Host "User does not exist. Proceeding with setup"
    $UserExists = $false
}
}

#Complex passwords

# Import System.Web assembly
#Add-Type -AssemblyName System.Web
# Generate random password
#$pw = [System.Web.Security.Membership]::GeneratePassword(10,0)

#make user
function makeUser(){

$first = Read-Host "First name" 
$last = Read-Host "Last name"
$user = $first.ToLower() + $last.Substring(0, [Math]::Min($last.Length, 1)).ToLower()

#Old user join. Keep in case new one breaks
#$user = -join($first, $last.Substring(0, [Math]::Min($last.Length, 1)));


validateUsername

if ($global:setfullname -eq $true){
$user = -join($first.ToLower(), $last.ToLower())
}


$ext = Read-Host "Extension"
$title = Read-Host "Title"

#Set this to a week in advance later, then check for year/month/date. Right now this is just a user input
$expirydate = Read-Host "Expiry date"
$expdate = Get-Date $expirydate

#Select company
Write-Host "Atlantic: 1"
Write-Host "Alchemy: 2"
Write-Host "Zoo: 3"
$company = Read-Host "Choose company from list: "
#join user and company

if ($company -eq 1){
$domain = "@atlanticproductions.co.uk"
$usercompany = "ATLANTIC PRODUCTIONS LIMITED"
$userOU = "OU=O365-Users,DC=atlantic,DC=local"
}
if ($company -eq 2){
$domain = "@alchemyimmersive.com"
$usercompany = "ATLANTIC PRODUCTIONS LIMITED"
$userOU = "OU=Alchemy,OU=O365-Users,DC=atlantic,DC=local"
}
if ($company -eq 3){
$domain = "@zoovfx.com"
$usercompany = "ZOO LIMITED"
$userOU = "OU=ZooVFX,OU=O365-Users,DC=atlantic,DC=local"
}

$useremail = -join($user, $domain)
Write-Host "Setting user email to", $useremail


$srcuser = Read-Host "Username to copy"

#Create a new user with the provided information and some static information

$newuserattributes = Get-ADUser -Identity $srcuser -Properties StreetAddress,City,State,PostalCode,Office,Department,Manager,wWWHomePage,Country
New-ADUser -Name "$first $last" -GivenName "$first" -Surname "$last" -SAMAccountName "$user" -Instance $newuserattributes -DisplayName "$First $Last" -EmailAddress $useremail -Description "Expires:$expirydate" -Title "$title" -Company $usercompany -AccountPassword $securepassword -UserPrincipalName "$user@atlantic.local" -Path "OU=O365-Users,DC=atlantic,DC=local" -OtherAttributes @{'TelephoneNumber'="+44 20 8735 9$ext"; 'ProxyAddresses'="SMTP:$useremail"} -Enabled $true
#New-ADUser "$first $last" -GivenName "$first" -Surname "$last" -DisplayName "$First $Last" -SamAccountName "$user" -Title "$title" -Description "$expirydate" -Office "Office Location" -EmailAddress $useremail -POBox "Suite Address" -City "City Name" -State "State Name" -PostalCode "Postal Code" -Country "Country Name" -UserPrincipalName "$user@atlantic.local" -HomeDirectory "\\domain.net\home$\$user" -Title "$title" -Company $usercompany -Path "OU=O365-Users,DC=atlantic,DC=local" -AccountPassword $pw -OtherAttributes @{'TelephoneNumber'="+44 20 8735 9$ext"; 'ProxyAddresses'="SMTP:$useremail"; 'Department'="$department"} -Enabled $true 

#Copy group membership of the source user above
Get-ADUser -Identity "$srcuser" -Properties memberof |
Select-Object -ExpandProperty memberof |
Add-ADGroupMember -Members "$user" -PassThru |
Select-Object -Property SamAccountName >$null



#Write-Host 'CHECK AD REPLICATION BEFORE CONTINUING!'
#pause


Set-ADAccountExpiration -Identity $user -DateTime $expdate.AddDays(1)
get-aduser $user | move-adobject -targetpath $userOU

Write-Host "$user is set up in AD"


$yesEmail = Read-Host "Send email? y/n"
if ($yesEmail -eq "y"){
    sendEmail
}
}


#write out
function sendEmail(){ 

$passphrase = "12345"
$emailTo = Read-Host "Email to: " 

$username = Get-ADUser -Identity $user -Properties GivenName, SamAccountName, EmailAddress | select-object GivenName, SamAccountName, EmailAddress
Write-Host 'Here are the details for the user' -ForegroundColor Green
$line0 = 'Hello',$username.GivenName
$line1 = ,'Welcome to Atlantic! Here are the details for your account'
$line2 = 'Username: ',$username.SamAccountName
Write-Host ''
$line3 = 'Email: ',$username.EmailAddress
Write-Host ''
$Secret = New-OTSSharedSecret -Secret $pw -Passphrase $passphrase
#$line4 = 'Password: ',$pw
$line4 = "Password accesible at: https://onetimesecret.com/secret/$($secret.SecretKey)"
$line5 = " "
$line6 = "To access your password, go to the link above and enter the passphrase, $passphrase"
$line7 = 'If you have any issues logging in, feel free to contact the IT team at IT@atlanticproductions.co.uk'
Write-Host ''


$MessageBody = "$line0`n`n$line1`n`n`n$line2`n`n$line3`n`n$line4 `n`n$line5`n`n$line6`n`n$line7"
$MessageBody2 = "Here is a guide on how to get started"

Write-Output $MessageBody

Send-MailMessage -SmtpServer eu-smtp-outbound-1.mimecast.com -From '<markuss@atlanticproductions.co.uk>' -To $emailTo -Subject 'Welcome to Atlantic' -body $MessageBody
#Send-MailMessage -SmtpServer eu-smtp-outbound-1.mimecast.com -From '<markuss@atlanticproductions.co.uk>' -To $username.EmailAddress -Subject 'Welcome to Atlantic Guide' -body $MessageBody2 -Attachments "C:\Users\markuss\Documents\Utilities\Powershell\2021-03-26 New user IT Guide.pdf"
}
