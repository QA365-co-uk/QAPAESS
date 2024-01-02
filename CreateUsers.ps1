<#
.SYNOPSIS
A Script to set up users in an Microsoft 365 tenant
.DESCRIPTION
This script file is designed to be used with custom QA Office 365 courses to 
set up test users.  It requires admin rights and the user will be prompted to 
sign in as an Office 365 Global admin account but all other aspects are automated.

The user information is stored as an array of hash tables in a variable called $Users,
(lines 72-94).

Static settings are stored in a seperate hash table (line 92)  and the two are combined  
in the loop that creates the users.

There are no parameters for this file as there are no inputs other than the Office 365 
Global admin account.
.NOTES
Author: Bret Appleby 
Version: 1.0

#>


[cmdletbinding()]
Param()

Write-Verbose "Check if NuGet is installed, if not install it"
$Packages = Get-PackageProvider
$NuGet = $null
$NuGet = $Packages | Where-Object Name -eq NuGet
If ($NuGet) {
    Write-Verbose "NuGet already installed - skipping install"
} Else {
    Write-Verbose "NuGet not installed - installing"
    Install-PackageProvider -Name NuGet -Force
    Write-Verbose "NuGet installed"
} #End If

Write-Verbose "Check if the AzaureAD module is installed, if not install it"
$Modules = Get-Module -ListAvailable
$AzureModule = $null
$AzureModule = $Modules | Where-Object Name -Like *AzureAD*
If ($AzureModule) {
    Write-Verbose "AzureAD Module present - skipping install"
} Else {
    Write-Verbose "AzureAD Module not present - installing"
    Install-Module -Name AzureADPreview -SkipPublisherCheck -Force
    Write-Verbose "AzureAD Module Installed"
} #End If

Write-Verbose "Enter Credentials and Check connection"
do {
    try {
        $Connect = $True
        $Cred = Get-Credential -Message "Office 365 Global admin account"
        Connect-AzureAD -Credential $Cred -ErrorAction Stop
    }
    Catch {
        $Connect = $False
        Write-Warning "Failed to connect with those credentials. Try again"
    } #End Try (Connect-AzureAD)
} until ($Connect -eq $True) #End Do
Write-Verbose "Connected to AzureAD as $($Cred.UserName)"

Write-Verbose "Extract domain name"
$DomainName = (Get-AzureADDomain | where name -like *onmicrosoft* | select -ExpandProperty name).split('.') | select -First 1
Write-Verbose "Domain name is $DomainName"

Write-Verbose "Set up array of users"
$Users = 
@{UserPrincipalName="Maker01@$DomainName.onmicrosoft.com";DisplayName="Matt Smith";GivenName="Matt";Surname="Smith";MailNickName="Matt"} ,
@{UserPrincipalName="Maker02@$DomainName.onmicrosoft.com";DisplayName="Ben Gregory";GivenName="Ben";Surname="Gregory";MailNickName="Ben"} ,
@{UserPrincipalName="Maker03@$DomainName.onmicrosoft.com";DisplayName="Al Martin";GivenName="Al";Surname="Martin";MailNickName="Al"},
@{UserPrincipalName="Maker04@$DomainName.onmicrosoft.com";DisplayName="Kelly Carey";GivenName="Kelly";Surname="Carey";MailNickName="Kelly"} ,
@{UserPrincipalName="Maker05@$DomainName.onmicrosoft.com";DisplayName="Susan Brown";GivenName="Susan";Surname="Brown";MailNickName="Susan"} ,
@{UserPrincipalName="Maker06@$DomainName.onmicrosoft.com";DisplayName="Vicky Slade";GivenName="Vicky";Surname="Slade";MailNickName="Vicky"} ,
@{UserPrincipalName="Maker07@$DomainName.onmicrosoft.com";DisplayName="Colin Montgomery";GivenName="Colin";Surname="Montgomery";MailNickName="Colin"} ,
@{UserPrincipalName="Maker08@$DomainName.onmicrosoft.com";DisplayName="Olivia Turner";GivenName="Olivia";Surname="Turner";MailNickName="Olivia"} ,
@{UserPrincipalName="Maker09@$DomainName.onmicrosoft.com";DisplayName="Edwina Carter";GivenName="Edwina";Surname="Carter";MailNickName="Edwina"} ,
@{UserPrincipalName="Maker10@$DomainName.onmicrosoft.com";DisplayName="Naomi Reed";GivenName="Naomi";Surname="Reed";MailNickName="Naomi"} ,
@{UserPrincipalName="Maker11@$DomainName.onmicrosoft.com";DisplayName="Kate Humphreys";GivenName="Kate";Surname="Humphreys";MailNickName="Kate"} ,
@{UserPrincipalName="Maker12@$DomainName.onmicrosoft.com";DisplayName="Dylan Cartwright";GivenName="Dylan";Surname="Cartwright";MailNickName="Dylan"} ,
@{UserPrincipalName="User01@$DomainName.onmicrosoft.com";DisplayName="Sofia Simmons";GivenName="Sofia";Surname="Simmons";MailNickName="Sofia"} ,
@{UserPrincipalName="User02@$DomainName.onmicrosoft.com";DisplayName="Charlie Briggs";GivenName="Charlie";Surname="Briggs";MailNickName="Charlie"} ,
@{UserPrincipalName="User03@$DomainName.onmicrosoft.com";DisplayName="Owen Barrett";GivenName="Owen";Surname="Barrett";MailNickName="Owen"},
@{UserPrincipalName="User04@$DomainName.onmicrosoft.com";DisplayName="Brooke Osborne";GivenName="Brooke";Surname="Osborne";MailNickName="Brooke"} ,
@{UserPrincipalName="User05@$DomainName.onmicrosoft.com";DisplayName="Bethany Parsons";GivenName="Bethany";Surname="Parsons";MailNickName="Bethany"} ,
@{UserPrincipalName="User06@$DomainName.onmicrosoft.com";DisplayName="Anthony Rose";GivenName="Anthony";Surname="Rose";MailNickName="Anthony"} ,
@{UserPrincipalName="User07@$DomainName.onmicrosoft.com";DisplayName="Harriet Kelly";GivenName="Harriet";Surname="Kelly";MailNickName="Harriet"} ,
@{UserPrincipalName="User08@$DomainName.onmicrosoft.com";DisplayName="Peter Hargreaves";GivenName="Peter";Surname="Hargreaves";MailNickName="Peter"} ,
@{UserPrincipalName="User09@$DomainName.onmicrosoft.com";DisplayName="Alex Osborne";GivenName="Alex";Surname="Osborne";MailNickName="Alex"} ,
@{UserPrincipalName="User10@$DomainName.onmicrosoft.com";DisplayName="Rosie Gilbert";GivenName="Rosie";Surname="Gilbert";MailNickName="Rosie"} ,
@{UserPrincipalName="User11@$DomainName.onmicrosoft.com";DisplayName="Harry Bishop";GivenName="Harry";Surname="Bishop";MailNickName="Harry"} ,
@{UserPrincipalName="User12@$DomainName.onmicrosoft.com";DisplayName="Lauren Schofield";GivenName="Lauren";Surname="Schofield";MailNickName="Lauren"}

Write-Verbose "Creating Password Profile"
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = "Apple#2411"
$passwordProfile.ForceChangePasswordNextLogin = $False
Write-Verbose "Password profile Object is $PasswordProfile"

#Setting up Variables
Write-Verbose "Setting up variables"
Write-Verbose "Setting static values"
$StaticSettings = @{PasswordProfile=$PasswordProfile;usagelocation='GB';AccountEnabled=$True}
Write-Verbose "`$StaticSettings is $StaticSettings"

Write-Verbose "Creating License Object"
$SKUID = Get-AzureADSubscribedSku | where skupartnumber -like *enterprise* | select -first 1 -ExpandProperty skuid
$License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$License.SkuId = $SKUID
$AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$AssignedLicenses.AddLicenses = $License
Write-Verbose "License object is $AssignedLicenses"

# Main loop to create users
Write-Verbose "Creating Users"
Foreach ($User in $Users)
{
    # Add static settings to current user
    $User += $StaticSettings
    
    # Create new user with splatting
    $NewUser = New-AzureADUser @User
    
    # Assign license to new user account
    Set-AzureADUserLicense -ObjectId $NewUser.ObjectId -AssignedLicenses $AssignedLicenses
    
    # Display User for trainer sanity purposes
    "Created User $($NewUser.DisplayName)"

} # End ForEach

