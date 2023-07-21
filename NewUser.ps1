Write-Host "Please be patient while back-end functions load silently"

# This gets domain administrator credentials
$UserCredential = Get-Credential -Message "Enter domain admin credentials"

# Connect to email server remote PowerShell
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://bapi64ml12/PowerShell -Authentication Kerberos
Import-PSSession $Session

# Connect to Exchange Online
$credentials = Get-Credential -Message "Enter Office 365 admin credentials"
Connect-ExchangeOnline -Credential $credentials

# Clears the shell to make it look organized
cls

# Prompt for user details
$FirstName = Read-Host "Enter the new employee's first name"
$LastName = Read-Host "Enter the new employee's last name"
$FirstInitial = $FirstName.Substring(0, 1).ToUpper()
$Login = "$FirstName$LastName"
$PhoneNumber = Read-Host "Enter phone number (XXX)-XXX-XXXX Ext. XXX"
$OnSite = Read-Host "Is this person on-site? (Yes/No)"
$Department = Read-Host "What department is this person in?"

# Password Generator
$Prefix = 'WelcomeToBAPI'
$Numbers = -join ((48..57) | Get-Random -Count 3 | % {[char]$_})
$Symbols = -join ((35..38)+(42)+(63..64)+(33) | Get-Random -Count 2 | % {[char]$_})
$Password = $Prefix + $Numbers + $Symbols
Write-Output "Temporary Password: $Password"

# Write Password to Desktop
$UserPath = "$($env:USERPROFILE)\Desktop"
cd\
cd $UserPath
New-Item $FirstInitial$LastName.txt
Set-Content -Path "$UserPath\$FirstInitial$LastName.txt" -Value "$Password"

# Convert to Secure String for use as password
$Password = ConvertTo-SecureString -String $Password -AsPlainText -Force

# New O365 Mailbox Creation
New-RemoteMailbox -Name "$firstName $lastName" -Password $Password -PrimarySmtpAddress $FirstInitial$lastname@bapisensors.com -UserPrincipalName $FirstInitial$lastname@bapihvac.lan 

# This pauses the script to wait for the On-Prem AD to create the user account. This script will fail without this.
Start-Sleep -Seconds 5

# Update on-premises AD attributes
$ADUser = Get-ADUser -Filter "Name -eq '$FirstName $LastName'"
$ADUser | Set-ADUser -GivenName $FirstName -Surname $LastName -OfficePhone $PhoneNumber

if ($OnSite -eq "Yes") {
    # Set on-site address
    Set-ADUser -Identity "$ADUser" -StreetAddress "750 N Royal Ave" -City Gays-Mills -State Wisconsin -PostalCode 54631 -Country US
} else {
    # Prompt for off-site address
    $Street = Read-Host "Enter street address"
    $City = Read-Host "Enter city"
    $State = Read-Host "Enter state/province"
    $PostalCode = Read-Host "Enter ZIP/postal code"
    $Country = Read-Host "Enter country/region (use US for United States)"
    
    # Update off-site address
    $ADUser | Set-ADUser -StreetAddress "$Street" -City $City -State $State -PostalCode $PostalCode -Country $Country
}

#Move user to their respective department in Users-BAPI OU
Move-ADObject -Identity $ADUser -TargetPath "OU=$Department,OU=Users-BAPI,DC=bapihvac,DC=lan"

# Connect to remote PowerShell session
$Session = New-PSSession -ComputerName BAPIDC-VIRT-ONE
Enter-PSSession -Session $Session

# Import ADSync module
Invoke-Command -Session $Session -ScriptBlock {
    Import-Module ADSync
}

# Get ADSync scheduler
Invoke-Command -Session $Session -ScriptBlock {
    Get-ADSyncScheduler
}

# Start ADSync synchronization cycle
Invoke-Command -Session $Session -ScriptBlock {
    Start-ADSyncSyncCycle -PolicyType Delta
}

# Exit the remote powershell session
Invoke-Command -Session $Session -ScriptBlock {
    Exit
}

# Creates email aliases
Set-RemoteMailbox $FirstInitial$lastname@bapisensors.com –EmailAddresses @{Add="$FirstInitial$LastName@bapihvac.com"}
Set-RemoteMailbox $FirstInitial$lastname@bapisensors.com –EmailAddresses @{Add="$FirstInitial$LastName@bapihvac.co.uk"}
Set-RemoteMailbox $FirstInitial$lastname@bapisensors.com –EmailAddresses @{Add="$FirstName$LastName@bapihvac.com"}
Set-RemoteMailbox $FirstInitial$lastname@bapisensors.com –EmailAddresses @{Add="$FirstName$LastName@bapihvac.co.uk"}

Write-Host "The user $FirstName $LastName has been successfully created. The user is in the Active Directory OU Users-BAPI/$Department. The user still needs to be assigned permissions and given an M365 license. The temporary password is in a text file on your desktop named $FirstInitial$LastName.txt"

Start-Sleep -Seconds 10

Write-Host -NoNewLine 'Press any key to exit...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');