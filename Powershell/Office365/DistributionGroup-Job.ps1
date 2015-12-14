$ou = "OU=Departments,DC=ad,DC=proove,DC=com"
$pwFile = 'C:\Users\administrator\Scripts\pw.txt'

#Set and encrypt password
#read-host -assecurestring | convertfrom-securestring | out-file $pwFile

function addUsers ($email){
    $user | Foreach {Write-Host add-DistributionGroupMember -Identity $email -Member $_.UserPrincipalName
    add-DistributionGroupMember -Identity $email -Member $_.UserPrincipalName}
}

function delUsers ($email){
    $user | Foreach {Write-Host remove-DistributionGroupMember -Identity $email -Member $_.UserPrincipalName -Confirm:$False
    remove-DistributionGroupMember -Identity $email -Member $_.UserPrincipalName -Confirm:$False}
}

#Import PS Session
$username = "dtran@proove.com"
$password = cat $pwFile | convertto-securestring
$Credentials = new-object -typename System.Management.Automation.PSCredential `
         -argumentlist $username, $password
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credentials -Authentication Basic –AllowRedirection
Import-PSSession $Session

#Remove Disabled Users from ProoveEmp
$user = Get-ADUser -Filter {(Enabled -eq $False)} -SearchBase $ou
delUsers -email 'prooveemp@proove.com'

#Add Enabled Users to ProoveEmp
$user = Get-ADUser -Filter {(Enabled -eq $True)} -SearchBase $ou
addUsers -email 'prooveemp@proove.com'

#Remove Disabled Users from Irvine
$user = Get-ADUser -Filter {(Enabled -eq $False) -and (Office -eq 'Alton' -or Office -eq '26 Tech')} -SearchBase $ou
delUsers -email 'irvine@proove.com'

#Add Enabled Users to Irvine
$user = Get-ADUser -Filter {(Enabled -eq $True) -and (Office -eq 'Alton' -or Office -eq '26 Tech')} -SearchBase $ou
addUsers -email 'irvine@proove.com'

#Remove Disabled Users from 26 Tech
$user = Get-ADUser -Filter {(Enabled -eq $False) -and (Office -eq '26 Tech')} -SearchBase $ou
delUsers -email '26tech@proove.com'

#Add Enabled Users to 26 Tech
$user = Get-ADUser -Filter {(Enabled -eq $True) -and (Office -eq '26 Tech')} -SearchBase $ou
addUsers -email '26tech@proove.com'

#Remove Disabled Users from Alton
$user = Get-ADUser -Filter {(Enabled -eq $False) -and (Office -eq 'Alton')} -SearchBase $ou
delUsers -email 'alton@proove.com'

#Add Enabled Users to Alton
$user = Get-ADUser -Filter {(Enabled -eq $True) -and (Office -eq 'Alton')} -SearchBase $ou
addUsers -email 'alton@proove.com'

#Remove Disabled Users from Maryland
$user = Get-ADUser -Filter {(Enabled -eq $False) -and (Office -eq 'Maryland')} -SearchBase $ou
delUsers -email 'maryland@proove.com'

#Add Enabled Users to Maryland
$user = Get-ADUser -Filter {(Enabled -eq $True) -and (Office -eq 'Maryland')} -SearchBase $ou
addUsers -email 'maryland@proove.com'

#Remove Disabled Users from researchassistants
$user = Get-ADUser -Filter {(Enabled -eq $False)} -SearchBase "OU=Research Assistants,OU=Departments,DC=ad,DC=proove,DC=com"
delUsers -email 'researchassistants@proove.com'

#Add Enabled Users to researchassistants
$user = Get-ADUser -Filter {(Enabled -eq $True)} -SearchBase "OU=Research Assistants,OU=Departments,DC=ad,DC=proove,DC=com"
addUsers -email 'researchassistants@proove.com'

#Remove Disabled Users from accountmanagement
$user = Get-ADUser -Filter {(Enabled -eq $False)} -SearchBase "OU=Account Management,OU=Departments,DC=ad,DC=proove,DC=com"
delUsers -email 'accountmanagement@proove.com'

#Add Enabled Users to accountmanagement
$user = Get-ADUser -Filter {(Enabled -eq $True)} -SearchBase "OU=Account Management,OU=Departments,DC=ad,DC=proove,DC=com"
addUsers -email 'accountmanagement@proove.com'

#Remove Disabled Users from billing
$user = Get-ADUser -Filter {(Enabled -eq $False) -and (Office -eq 'Alton' -or Office -eq '26 Tech')} -SearchBase "OU=Financial,OU=Departments,DC=ad,DC=proove,DC=com"
#delUsers -email 'billing@proove.com'

#Add Enabled Users to billing
$user = Get-ADUser -Filter {(Enabled -eq $True) -and (Office -eq 'Alton' -or Office -eq '26 Tech')} -SearchBase "OU=Financial,OU=Departments,DC=ad,DC=proove,DC=com"
#addUsers -email 'billing@proove.com'

#Remove Disabled Users from billingmd
$user = Get-ADUser -Filter {(Enabled -eq $False) -and (Office -eq 'Maryland')} -SearchBase "OU=Financial,OU=Departments,DC=ad,DC=proove,DC=com"
delUsers -email 'billingmd@proove.com'

#Add Enabled Users to billingmd
$user = Get-ADUser -Filter {(Enabled -eq $True) -and (Office -eq 'Maryland')} -SearchBase "OU=Financial,OU=Departments,DC=ad,DC=proove,DC=com"
addUsers -email 'billingmd@proove.com'

#Terminate all live PS Sessions
$liveSessions = get-pssession
remove-pssession -session $liveSessions