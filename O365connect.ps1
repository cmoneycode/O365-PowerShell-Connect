<#---- Do these steps before running the script ----
1. Open PowerShell v3.0 or greater on Win7 or later OS
2. Find-Module MSOnline | Install-Module
3. Download and install the latest AzureAD module from the link in this tweet.
   Here's the link for the screenshot:
   https://www.powershellgallery.com/packages/AzureAD/2.0.0.155
4. Install-Module AzureAD
#>

#---- Script to Connect to Office 365 Exchange

#Create the UserCredential variable, with your O365 username/pw.
#The account must have Global Admin rights on O365.
$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -Credential $UserCredential -Authentication Basic -AllowRedirection

#Then you create a session, which later you want to manually Remove, when done.
Import-PSSession $session

#If you run into problems with the Connect-MsolService, 
#try without the $UserCredential variable.
Connect-MsolService -Credential $UserCredential

#When you are done with the session, don't forget to Remove it.
#Only 3 active sessions are allowed, so this is important.
Remove-PSSession $Session #use this to disconnect PSSession
