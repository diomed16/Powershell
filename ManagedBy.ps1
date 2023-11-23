<#
.SYNOPSIS
	Sets managedby computer attribute
.DESCRIPTION
	This PowerShell script sets managedby computer attribute and iphostnumber
.LINK
	https://github.com/diomed16/PowerShell
.NOTES
	Author: Konstantin Shumakov
#>
$ADS_PROPERTY_CLEAR = 1
$ADS_PROPERTY_UPDATE = 2
$ADS_PROPERTY_APPEND = 3
$ADS_PROPERTY_DELETE = 4
$User = $env:UserName
$comp=$env:computername
$ldapPath = ([ADSI]"LDAP://rootDSE").defaultNamingContext
$LDAPSEARCH = New-Object System.DirectoryServices.DirectorySearcher 
$LDAPSEARCH.SearchRoot = "LDAP://"+"$ldapPath"
$LDAPSEARCH.Filter = "(&(objectClass=computer)(Name=$comp))"
$LDAPSEARCH_RESULTS = $LDAPSEARCH.FindAll()
$computer = foreach ($LINE in $LDAPSEARCH_RESULTS) {
$LINE_ENTRY = $LINE.GetDirectoryEntry()
$LINE_ENTRY | Select-Object -Property *
}
$LDAPSEARCH.Filter = "(&(objectClass=user)(samaccountname=$User))"
$LDAPSEARCH_RESULTS = $LDAPSEARCH.FindAll()
$User = foreach ($LINE in $LDAPSEARCH_RESULTS) {
$LINE_ENTRY = $LINE.GetDirectoryEntry()
$LINE_ENTRY | Select-Object -Property *
}
$env:HostIP = (
Get-NetIPConfiguration |
Where-Object {
$_.IPv4DefaultGateway -ne $null -and
$_.NetAdapter.Status -ne "Disconnected"
}
).IPv4Address.IPAddress
$ldapPath = $computer.path
$attributeName = "managedby"
$attributeValue = $user.distinguishedname
$compip=$env:HostIP
[ADSI]$objcomputer=$computer.path
$objcomputer.Putex($ADS_PROPERTY_UPDATE,$attributeName, @($attributeValue))
$objcomputer.SetInfo()
$objcomputer.Put("iPHostnumber", $compip)
$objcomputer.SetInfo()

