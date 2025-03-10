write-host "`nThis script assigns Teams soft phone for a single user using 'Direct Routing' for the region`n" -ForegroundColor Yellow
$UsersEmail = Read-Host -prompt "Please enter the email address to enable Teams EV"
$LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
try {
    Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType DirectRouting -ErrorAction Stop
    write-host "Teams EV has been enabled for $UsersEmail"
    Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
}
catch {
    <#Print the error and exit PS Script#>
    write-host "`n `nAn error occured: $_" -ForegroundColor Red 
    Exit
}