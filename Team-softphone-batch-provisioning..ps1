[CmdletBinding()]
param( [string] $UsersList = $(Read-Host -prompt `
    "Please input the CSV File Name with Location, Example: C:\Test.csv"))
$conf = Import-Csv $UsersList
$Conf | ForEach {
Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $_.LineURI -PhoneNumberType DirectRouting
Set-CsOnlineDialInConferencingUser -Identity $_.upn -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
Grant-CsDialoutPolicy -Identity $_.upn -PolicyName "DialoutCPCDomesticPSTNInternational"
}
$transcriptname = "Userstats" + `
    (Get-Date -format s).Replace(":","-") +".txt"
$CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
$outputname = "Stats" + $CurrentDate + ".csv"
Start-Transcript $transcriptname
$count = $conf.count
write-host "We have found" $count "Users" -foregroundcolor Yellow -backgroundcolor Black
$Conf | ForEach {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $outputname