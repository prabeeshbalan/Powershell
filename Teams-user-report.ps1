[CmdletBinding()]
param( [string] $UsersList = $(Read-Host -prompt `
    "Input the CSV File with Location"))
$conf = Import-Csv $UsersList -Delimiter ";"
$transcriptname = "Userstats" + `
    (Get-Date -format s).Replace(":","-") +".txt"
$CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
$outputname = "Stats" + $CurrentDate + ".csv"
Start-Transcript $transcriptname
$count = $conf.count
write-host "We have found" $count "Users" -foregroundcolor Yellow -backgroundcolor Black
$Conf | ForEach {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, TeamsMeetingPolicy, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $outputname