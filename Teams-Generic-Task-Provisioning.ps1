# Microsft Teams Admin Console General Day to Day task handling via script
# Author: Prabeesh Balan
# Description: This script helps to accomplish some day to day Microsft Teams administration function like Teams Softphone Activation and Removal, Teams International call, Meeting Recording, Audio-conferencing nubmer change etc


####################################################################################################

Do{
    # Main Menu for selection
    Write-Host "`n------------------------------------------------------- " -ForegroundColor Cyan
    Write-Host "Select the type of action you want to perform in Teams:" -ForegroundColor Cyan
    Write-Host "------------------------------------------------------- `n" -ForegroundColor Cyan
    Write-Host "1. Teams User Information" -ForegroundColor Yellow
    Write-Host "2. Teams Softphone Activation" -ForegroundColor Yellow
    Write-Host "3. Teams Softphone Removal" -ForegroundColor Yellow
    Write-Host "4. Teams International calling" -ForegroundColor Yellow
    Write-Host "5. Teams Meeting Recording" -ForegroundColor Yellow 
    Write-Host "6. Teams Audio-conferencing Number Update" -ForegroundColor Yellow
    Write-Host "7. Teams Softphone Caller ID " -ForegroundColor Yellow
    Write-Host "8. Exit`n" -ForegroundColor Red
    $selection = Read-Host "Please enter your choice (1, 2, 3, 4, 5, 6, 7 or 8)"
    
    switch ($selection) {

        '1' {
                # This sesseion retrieves user information / a report of bulk users 
                Write-Host "`nIs this a bulk request (Answer: Yes/No)?"
                $answer = Read-Host "Please enter your choice"

                if ($answer -ieq "No") 
                {
                    # This condition will retrieves a specific user information 
                    Write-Host " "
                    $upn = Read-Host -prompt "Please enter the user email address"
                    Get-CsOnlineUser -identity $upn | Format-List UserPrincipalName, DisplayName, SipAddress, City, UsageLocation, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, LineURI, OnlineVoiceRoutingPolicy, TeamsMeetingPolicy, HostingProvider, InterpretedUserType
                }elseif ($answer -ieq "Yes")
                {
                    # This condition will pull a report of bulk users based on the CSV file provided with UPN
                    # Make sure CSV file column name is "upn"

                    Write-Host " "
                    $UsersList = Read-Host -prompt "Please enter the CSV file path and name, Example: C:\folder\filename.csv"
                    $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                    $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"

                    # Import the list of users from the CSV file
                    $users = Import-Csv -Path $UsersList

                    # Initialize an array to store the report data
                    $report = @()

                    # Loop through each user and retrieve their details
                    foreach ($user in $users) {
                        $UserPrincipalName = $user.upn

                        # Get the Teams user details
                        $teamsUser = Get-CsOnlineUser -Identity $UserPrincipalName | Select-Object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, TeamsMeetingPolicy, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType

                        # Add the user's details to the report array
                        $report += [PSCustomObject]@{

                            DisplayName = $teamsUser.DisplayName
                            SoftPhoneNumber = $teamsUser.LineUri
                            Email = $teamsUser.UserPrincipalName
                            onpremsipaddress = $teamsUser.onpremsipaddress
                            sipaddress = $teamsUser.sipaddress
                            Enabled = $teamsUser.Enabled
                            TeamsUpgradeEffectiveMode = $teamsUser.TeamsUpgradeEffectiveMode
                            EnterpriseVoiceEnabled = $teamsUser.EnterpriseVoiceEnabled
                            TeamsMeetingPolicy = $teamsUser.TeamsMeetingPolicy
                            HostedVoiceMail = $teamsUser.HostedVoiceMail
                            OnlineVoiceRoutingPolicy = $teamsUser.OnlineVoiceRoutingPolicy
                            OnPremLineURI = $teamsUser.OnPremLineURI
                            OnlineDialinConferencingPolicy = $teamsUser.OnlineDialinConferencingPolicy
                            HostingProvider = $teamsUser.HostingProvider
                            InterpretedUserType = $teamsUser.InterpretedUserType
                        }
                    }
                    # Export the report to a CSV file
                    $report | Export-Csv -Path $OutputCsvPath -NoTypeInformation
                    Write-Host "`nReport generated successfully: $OutputCsvPath`n"
                }
                    
            }

        '2' {
                # This sesseion enables Teams Enterprise Voice for one/bulk user(s) based on their location available in the organization as on Feb 2025
                # This is the menu for location selection
                Write-Host "`n--------------------------------------------------------"
                Write-Host "Select the user location for Teams Softphone Activation:" -ForegroundColor Cyan
                Write-Host "--------------------------------------------------------`n"
                Write-Host "1. Australia" -ForegroundColor Yellow
                Write-Host "2. Canada" -ForegroundColor Yellow
                Write-Host "3. Dublin" -ForegroundColor Yellow
                Write-Host "4. Hong Kong" -ForegroundColor Yellow
                Write-Host "5. London" -ForegroundColor Yellow
                Write-Host "6. Singapore" -ForegroundColor Yellow
                Write-Host "7. United States`n" -ForegroundColor Yellow

                # Gets the location in to a variable
                $answer = Read-Host "Please enter your choice"
                Switch($answer){
                    
                    '1' {
                            # This session enabled Teams EV for the location: Australia
                            Write-Host "`nYou have selected the user location: Australia`n" -ForegroundColor Red
                                                       
                            # Gets the choice in to varialbe for sigle or bulk user(s)
                            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                            if($choice -ieq "No")
                            {
                                # Enabling single user teams EV when if condition satisfies the choice 1
                                #Gets the user upn and lineuri in to the variables
                                $UsersEmail = Read-Host -prompt "`nPlease enter the email address to enable Teams EV"
                                $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                                try {
                                        #Assigning toll and toll free nubmer to the user
                                        Set-CsOnlineDialInConferencingUser -Identity $UsersEmail -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                        
                                        #Sets Dial out policy to all destination
                                        Grant-CsDialoutPolicy -Identity $UsersEmail -PolicyName "DialoutCPCDomesticPSTNInternational"
                                        #Sets Voice routing policy to Australia
                                        Grant-CsOnlineVoiceRoutingPolicy $UsersEmail -PolicyName "Australia"
                                        #Assgning Teams EV number using the variables
                                        Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType OperatorConnect
                                        write-host "Teams EV has been enabled for $UsersEmail"
                                        #Displays user Information after Teams EV assignment completed
                                        Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                    }
                                catch {
                                        <#Print the error and exit PS Script#>
                                        write-host "`n `nAn error occured: $_" -ForegroundColor Red 
                                    }

                            } else {
                                    try {
                                            #Enabling Teams EV based on the CSV file provided
                                            #Make sure the input CSV file has two columns: upn and lineuri
                                            
                                            # Import the list of users and lineuri from the CSV file 
                                            $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                                            $conf = Import-Csv $Bulkevuserslist
                                            $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                                            
                                            #Adding currect date into the output csv file name
                                            $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                                            $Conf | ForEach-Object {
                                                #Goes through each object and performing the following tasks
                                                                                                
                                                Set-CsOnlineDialInConferencingUser -Identity $_.upn -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                                #Sets Dial out policy to all destination
                                                Grant-CsDialoutPolicy -Identity $_.upn -PolicyName "DialoutCPCDomesticPSTNInternational"
                                                Grant-CsOnlineVoiceRoutingPolicy $_.upn -PolicyName "Australia"
                                                Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $LineURI -PhoneNumberType OperatorConnect
                                                write-host "Teams EV has been enabled for $_.upn"
                                                Get-CsOnlineUser -identity $_.upn | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                            }
                                            $count = $conf.count
                                            write-host "Your file has " $count "Users" -foregroundcolor Yellow -backgroundcolor Black
                                            $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                        }
                                    catch {
                                            <#Print the error and exit PS Script#>
                                            write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                        }
                                    }  
                        }

                    '2' {
                            # This session enabled Teams EV for the location: Canada
                            Write-Host "`nYou have selected the user location: Canada`n" -ForegroundColor Red
                                                                                  
                            # Gets the choice in to varialbe for sigle or bulk user(s)
                            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                            
                            if($choice -ieq "No")
                            {
                                Write-Host "`nYou have selected the option: 1. For user Teams EV Activation`n" -ForegroundColor Yellow
                                $UsersEmail = Read-Host -prompt "`nPlease enter the email address to enable Teams EV"
                                $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                                try {
                                    Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType DirectRouting -ErrorAction Stop
                                    write-host "Teams EV has been enabled for $UsersEmail"
                                    Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                }
                                catch {
                                    <#Print the error and exit PS Script#>
                                    write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                }

                            } else {
                                try {
                                    $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                                    $conf = Import-Csv $Bulkevuserslist
                                    $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                                    $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                                    $Conf | ForEach-Object {
                                    Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $_.LineURI -PhoneNumberType DirectRouting
                                    }
                                    $count = $conf.count
                                    write-host "Your file has " $count "Users" -foregroundcolor Yellow -backgroundcolor Black
                                    $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                }
                                catch {
                                    <#Print the error and exit PS Script#>
                                    write-host "`n `nAn error occured: $_" -ForegroundColor Red 
                                }
                            }
                        }

                    '3' {
                            # This session enabled Teams EV for the location: Dublin
                            Write-Host "`nYou have selected the user location: Dublin`n" -ForegroundColor Red
                                                                                  
                            # Gets the choice in to varialbe for sigle or bulk user(s)
                            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                            if($choice -ieq "No")
                            {
                                $UsersEmail = Read-Host -prompt "`nPlease enter the email address to enable Teams EV"
                                $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                                try {
                                        Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType DirectRouting
                                        Set-CsOnlineDialInConferencingUser -Identity $UsersEmail -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                        Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $UsersEmail -PolicyName "EuropeEmergencyCalling"
                                        Grant-CsTeamsEmergencyCallingPolicy -Identity $UsersEmail -PolicyName "EuropeEmergencyCalling"
                                        Grant-CsTenantDialPlan -Identity $UsersEmail -PolicyName "Ireland"
                                        Grant-CsOnlineVoiceRoutingPolicy $UsersEmail -PolicyName "Europe"
                                        write-host "Teams EV has been enabled for $UsersEmail"
                                        Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                    }
                                catch {
                                        <#Print the error and exit PS Script#>
                                        write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                    }

                            } else {
                                    try {
                                            $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                                            $conf = Import-Csv $Bulkevuserslist
                                            $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                                            $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                                            $Conf | ForEach-Object {

                                                Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $_.LineURI -PhoneNumberType DirectRouting
                                                Set-CsOnlineDialInConferencingUser -Identity $_.upn -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                                Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $_.upn -PolicyName "EuropeEmergencyCalling"
                                                Grant-CsTeamsEmergencyCallingPolicy -Identity $_.upn -PolicyName "EuropeEmergencyCalling"
                                                Grant-CsTenantDialPlan -Identity $_.upn -PolicyName "Ireland"
                                                Grant-CsOnlineVoiceRoutingPolicy $_.upn -PolicyName "Europe"
                                                write-host "Teams EV has been enabled for $_.upn"
                                                Get-CsOnlineUser -identity $_.upn | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName 
                                            }
                                            $count = $conf.count
                                            write-host "Your file has " $count "Users" -foregroundcolor Yellow -backgroundcolor Black
                                            $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                        }
                                    catch {
                                            <#Print the error and exit PS Script#>
                                            write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                        }
                                    } 
                        }

                    '4' {
                            # This session enabled Teams EV for the location: Hong Kong
                            Write-Host "`nYou have selected the user location: Hong Kong`n" -ForegroundColor Red
                                                                                  
                            # Gets the choice in to varialbe for sigle or bulk user(s)
                            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                            if($choice -ieq "No")
                            {
                                $UsersEmail = Read-Host -prompt "`nPlease enter the email address to enable Teams EV"
                                $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                                try {
                                        
                                        Set-CsOnlineDialInConferencingUser -Identity $UsersEmail -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                        Grant-CsDialoutPolicy -Identity $UsersEmail -PolicyName "DialoutCPCDomesticPSTNInternational"
                                        Grant-CsOnlineVoiceRoutingPolicy $UsersEmail -PolicyName "HK"
                                        Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType OperatorConnect
                                        write-host "Teams EV has been enabled for $UsersEmail"
                                        Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                    }
                                catch {
                                        #Print the error and exit PS Script
                                        write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                    }

                            } else {
                                    try {
                                            $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                                            $conf = Import-Csv $Bulkevuserslist
                                            $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                                            $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                                            $Conf | ForEach-Object {

                                                Set-CsOnlineDialInConferencingUser -Identity $_.upn -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                                Grant-CsDialoutPolicy -Identity $_.upn -PolicyName "DialoutCPCDomesticPSTNInternational"
                                                Grant-CsOnlineVoiceRoutingPolicy $_.upn -PolicyName "HK"
                                                Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $_.LineURI -PhoneNumberType OperatorConnect
                                                write-host "Teams EV has been enabled for $_.upn"
                                                Get-CsOnlineUser -identity $_.upn | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                
                                            }
                                            $count = $conf.count
                                            write-host "Your file has " $count "Users" -foregroundcolor Yellow -backgroundcolor Black
                                            $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                        }
                                    catch {
                                            #Print the error and exit PS Script
                                            write-host "`n `nAn error occured: $_" -ForegroundColor Red 
                                        }
                                    }  
                        }

                    '5' {
                            # This session enabled Teams EV for the location: London
                            Write-Host "`nYou have selected the user location: London`n" -ForegroundColor Red
                                                                                  
                            # Gets the choice in to varialbe for sigle or bulk user(s)
                            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                            
                            if($choice -ieq "No")
                            {
                                $UsersEmail = Read-Host -prompt "`nPlease enter the email address to enable Teams EV"
                                $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                                try {
                                        Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType DirectRouting
                                        Set-CsOnlineDialInConferencingUser -Identity $UsersEmail -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                        Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $UsersEmail -PolicyName "EuropeEmergencyCalling"
                                        Grant-CsTeamsEmergencyCallingPolicy -Identity $UsersEmail -PolicyName "EuropeEmergencyCalling"
                                        Grant-CsTenantDialPlan -Identity $UsersEmail -PolicyName "UK"
                                        Grant-CsOnlineVoiceRoutingPolicy $UsersEmail -PolicyName "Europe"
                                        write-host "Teams EV has been enabled for $UsersEmail"
                                        Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                    }
                                catch {
                                        <#Print the error and exit PS Script#>
                                        write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                    }

                            } else {
                                try {
                                        $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                                        $conf = Import-Csv $Bulkevuserslist
                                        $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                                        $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                                        $Conf | ForEach-Object {

                                            Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $_.LineURI -PhoneNumberType DirectRouting
                                            Set-CsOnlineDialInConferencingUser -Identity $_.upn -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                            Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $_.upn -PolicyName "EuropeEmergencyCalling"
                                            Grant-CsTeamsEmergencyCallingPolicy -Identity $_.upn -PolicyName "EuropeEmergencyCalling"
                                            Grant-CsTenantDialPlan -Identity $_.upn -PolicyName "UK"
                                            Grant-CsOnlineVoiceRoutingPolicy $_.upn -PolicyName "Europe"
                                            write-host "Teams EV has been enabled for $_.upn"
                                            Get-CsOnlineUser -identity $_.upn | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName 
                                        }
                                        $count = $conf.count
                                        write-host "Your file has " $count "Users" -foregroundcolor Yellow -backgroundcolor Black
                                        $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                    }
                                catch {
                                        <#Print the error and exit PS Script#>
                                        write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                    }
                            }               
                        }
    
                    '6' {
                            # This session enabled Teams EV for the location: Singapore
                            Write-Host "`nYou have selected the user location: Singapore`n" -ForegroundColor Red
                                                                                  
                            # Gets the choice in to varialbe for sigle or bulk user(s)
                            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                            
                            if($choice -ieq "No")
                            {
                                $UsersEmail = Read-Host -prompt "`nPlease enter the email address to enable Teams EV"
                                $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                                try {
                                        
                                        Set-CsOnlineDialInConferencingUser -Identity $UsersEmail -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                        Grant-CsDialoutPolicy -Identity $UsersEmail -PolicyName "DialoutCPCDomesticPSTNInternational"
                                        Grant-CsOnlineVoiceRoutingPolicy $UsersEmail -PolicyName "Singapore"
                                        Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType Operator Connect
                                        write-host "Teams EV has been enabled for $UsersEmail"
                                        Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                    }
                                catch {
                                        <#Print the error and exit PS Script#>
                                        write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                    }

                            } else {
                                    try {
                                            $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                                            $conf = Import-Csv $Bulkevuserslist
                                            $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                                            $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                                            $Conf | ForEach-Object {

                                                Set-CsOnlineDialInConferencingUser -Identity $_.upn -ServiceNumber yourservicenumber -TollFreeServiceNumber yourtollfreenumber
                                                Grant-CsDialoutPolicy -Identity $_.upn -PolicyName "DialoutCPCDomesticPSTNInternational"
                                                Grant-CsOnlineVoiceRoutingPolicy $_.upn -PolicyName "Singapore"
                                                Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $LineURI -PhoneNumberType Operator Connect
                                                write-host "Teams EV has been enabled for $_.upn"
                                                Get-CsOnlineUser -identity $_.upn | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                
                                            }
                                            $count = $conf.count
                                            write-host "Your file has " $count "Users" -foregroundcolor Yellow -backgroundcolor Black
                                            $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                        }
                                    catch {
                                            <#Print the error and exit PS Script#>
                                            write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                        }
                                    }                     
                        }

                    '7' {
                            # This session enabled Teams EV for the location: United States
                            Write-Host "`nYou have selected the user location: United States`n" -ForegroundColor Red
                                                                                  
                            # Gets the choice in to varialbe for sigle or bulk user(s)
                            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                            if($choice -ieq "No")
                            {
                                $UsersEmail = Read-Host -prompt "`nPlease enter the email address to enable Teams EV"
                                $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                                try {
                                    Set-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI -PhoneNumberType DirectRouting -ErrorAction Stop
                                    write-host "Teams EV has been enabled for $UsersEmail"
                                    Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType, VoicePolicy, CountryOrRegionDisplayName
                                }
                                catch {
                                    <#Print the error and exit PS Script#>
                                    write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                }

                            } else {
                                try {
                                    $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                                    $conf = Import-Csv $Bulkevuserslist
                                    $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                                    $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                                    $Conf | ForEach-Object {
                                    Set-CsPhoneNumberAssignment -Identity $_.upn -PhoneNumber $_.LineURI -PhoneNumberType DirectRouting
                                    }
                                    $count = $conf.count
                                    write-host "`nYour file has "$count "user(s)" -foregroundcolor Yellow -backgroundcolor Black
                                    $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                }
                                catch {
                                    <#Print the error and exit PS Script#>
                                    write-host "`n `nAn error occured: $_" -ForegroundColor Red
                                }
                            }               
                        }
                        Default {
                            #default
                            Write-Host "`nInvalid choice, please select 1, 2, or 3." -ForegroundColor Red
                        }
                }
            }
                            
        '3' {
                #This session will help to remove Teams EV for a user/bulk
                Write-Host "`nYou have selected the option to removed Teams Softphone for user(s)`n" -ForegroundColor Yellow
                
                # Gets the choice in to varialbe for sigle or bulk user(s)
                $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                    if($choice -ieq "No")
                    {
                        $UsersEmail = Read-Host -prompt "Please enter user email address to remove Teams EV "
                        $LineURI = Read-Host -prompt "Please enter the LineURI for $UsersEmail"
                        $userlineURI = Get-CsOnlineuser -identity UsersEmail | Select-Object LineURI
                        If($LineURI -ieq $userlineURI)
                        {
                            Remove-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $LineURI
                            write-host "Teams EV ($LineURI) has been removed for $UsersEmail"
                            Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, EnterpriseVoiceEnabled, City, UsageLocation, LineURI, TeamsMeetingPolicy, OnlineVoiceRoutingPolicy
                    
                        } else {

                                write-host "The EV nubmer provided ($LineURI) is diffecrent from the actual EV number ($userlineURI) assigned for $UsersEmail"
                                $Question = write-host "Do you still want to remove currently assigned nubmer: $userlineURI (Yes/No) ?"
                                if($Question -ieq "Yes")
                                {
                                    Remove-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $userlineURI
                                    write-host "Teams EV ($userlineURI) has been removed for $UsersEmail"
                                    Get-CsOnlineUser -identity $UsersEmail | Format-List UserPrincipalName, DisplayName, SipAddress, EnterpriseVoiceEnabled, City, UsageLocation, LineURI, TeamsMeetingPolicy, OnlineVoiceRoutingPolicy
                        
                                } else {
                                    <# Action when all if and elseif conditions are false #>
                                    write-host "Thank you for the confirmation. No changes made on user EV"
                                }
                        }   
                        
                    }else {
                            # Action EV removal on bulk request
                            $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                            $conf = Import-Csv $Bulkevuserslist
                            $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                            $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                            $Conf | ForEach-Object {

                                Remove-CsPhoneNumberAssignment -Identity $UsersEmail -PhoneNumber $userlineURI
                                write-host "Teams EV has been disabled for $_.upn"
                            }
                            $count = $conf.count
                            write-host "Your file has $count users" -foregroundcolor Yellow -backgroundcolor Black
                            $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
                                        

                    }
            }

        '4' {
                #This session will help to enable Teams International calling for user(s)
                Write-Host "`nYou have selected the option to enable Teams International calling for user(s)`n" -ForegroundColor Yellow
                
                # Gets the choice in to varialbe for sigle or bulk user(s)
                $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                    Write-Host " "
                    if($choice -ieq "Yes")
                    {
                        #This session will enable Teams International calling for bulk user list and generates a report
                        $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                        $conf = Import-Csv $UsersList
                        $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                        $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                        $Conf | ForEach-Object {
                        Grant-CsOnlineVoiceRoutingPolicy -Identity $_.upn -PolicyName "International"
                        }
                        $count = $conf.count
                        write-host "International Calling feature has been enabled for $count users" -foregroundcolor Yellow -backgroundcolor Black
                        $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, TeamsMeetingPolicy, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath

                    } else {
                        #This session will enable Teams International calling for one user
                        $UsersEmail = Read-Host -prompt "Please enter user email address to enable International calling"
                        Grant-CsOnlineVoiceRoutingPolicy -Identity $UsersEmail -PolicyName "International"
                    }
                
            }

        '5' {
                #This session will help to enable Teams Meeeting recording feature for user(s)
                Write-Host "`nYou have selected the option to enable Teams Meeeting Recording feature for user(s)`n" -ForegroundColor Yellow
                
                # Gets the choice in to varialbe for sigle or bulk user(s)
                $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
                    Write-Host " "
                    if($choice -ieq "Yes")
                    {
                        #This sesession will enable Teams recording feature for bulk user list and generates a report
                        $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                        $conf = Import-Csv $UsersList
                        $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                        $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                        $Conf | ForEach-Object {
                        Grant-CsTeamsMeetingPolicy -Identity $_.upn -PolicyName "Recording Allowed"
                        }
                        $count = $conf.count
                        write-host "International Calling feature has been enabled for $count users" -foregroundcolor Yellow -backgroundcolor Black
                        $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, TeamsMeetingPolicy, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath

                    } else {
                        #This sesession will enable Teams recording feature for one user
                        $UsersEmail = Read-Host -prompt "Please enter user email address to enable International calling"
                        Grant-CsTeamsMeetingPolicy -Identity $UsersEmail -PolicyName "Recording Allowed"
                    }
                
        }

        '6' {
            # Audio confecensing nubmer assignment
            Write-Host "------------------------------------"
            Write-Host "This session is under development:" -ForegroundColor Cyan
            Write-Host "------------------------------------"
            Write-Host " "
            <#Write-Host "1. For Single User" -ForegroundColor Yellow
            Write-Host "2. For Bulk User CSV Report" -ForegroundColor Yellow
            Write-Host " "
             $choice = Read-Host "Please enter your choice"
                    Write-Host " "
                    if($choice -eq 1)
                    {

                    } else {
                        
                    }#>
        }

        '7'{

            #This session will help to set Teams Softphone caller ID for user(s)
            Write-Host "`nYou have selected the option to set Teams softphone Caller ID for user(s)`n" -ForegroundColor Yellow
            
            Write-Host "`n------------------------------------"
            Write-Host "      Select the Teams Caller ID      " -ForegroundColor Cyan
            Write-Host "------------------------------------`n"
            
            Write-Host "1. policyname" -ForegroundColor Yellow
            Write-Host "2. policyname US" -ForegroundColor Yellow
            Write-Host "3. policyname Contact Center Solutions Support`n" -ForegroundColor Yellow

            $CalledID = Read-Host "Please enter your choice (1,2 or 3)"
            Write-Host " "

            # Gets the choice in to varialbe for sigle or bulk user(s)
            $choice = Read-Host "Is this a bulk request (Answer: Yes/No)?"
            
            switch ($CalledID) {

                '1'{
                    if($choice -ieq "Yes")
                    {
                        #This sesession will enable Teams Caller ID for bulk user list and generates a report
                        # Make sure CSV file column name is "upn"
                        $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                        $conf = Import-Csv $UsersList
                        $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                        $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                        $Conf | ForEach-Object {
                        
                            Grant-CsCallingLineIdentity -Identity $_.upn -PolicyName "policyname"
                        }
                        $count = $conf.count
                        write-host "Called ID has been enabled for $count users" -foregroundcolor Yellow -backgroundcolor Black
                        $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, TeamsMeetingPolicy, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
    
                    } else {
                            #This sesession will enable Teams Caller ID for one user
                            $UsersEmail = Read-Host -prompt "Please enter user email address to enable Caller ID"
                            Grant-CsCallingLineIdentity -Identity $_.upn -PolicyName "policyname"
                        }
                    }
                '2'{
                    if($choice -ieq "Yes")
                    {
                        #This sesession will enable Teams Caller ID for bulk user list and generates a report
                        # Make sure CSV file column name is "upn"
                        $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                        $conf = Import-Csv $UsersList
                        $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                        $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                        $Conf | ForEach-Object {
                        
                            Grant-CsCallingLineIdentity -Identity $_.upn -PolicyName "policyname US"
                        }
                        $count = $conf.count
                        write-host "Called ID has been enabled for $count users" -foregroundcolor Yellow -backgroundcolor Black
                        $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, TeamsMeetingPolicy, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
    
                    } else {
                        #This sesession will enable Teams Caller ID for one user
                        $UsersEmail = Read-Host -prompt "Please enter user email address to enable Caller ID"
                        Grant-CsCallingLineIdentity -Identity $_.upn -PolicyName "policyname US"
                    }
                    }
                '3' {
                    if($choice -ieq "Yes")
                    {
                        #This sesession will enable Teams Caller ID for bulk user list and generates a report
                        # Make sure CSV file column name is "upn"
                        $Bulkevuserslist = Read-Host -prompt "`nPlease enter the CSV file path and name, Example: C:\folder\filename.csv"
                        $conf = Import-Csv $UsersList
                        $CurrentDate = Get-Date -Format "MM_dd_yyyy_HHmm"
                        $OutputCsvPath = "Teams_user_report_" + $CurrentDate + ".csv"
                        $Conf | ForEach-Object {
                        
                            Grant-CsCallingLineIdentity -Identity $_.upn -PolicyName "policyname Contact Center Solutions Support"
                        }
                        $count = $conf.count
                        write-host "Called ID has been enabled for $count users" -foregroundcolor Yellow -backgroundcolor Black
                        $Conf | ForEach-Object {get-CsonlineUser -Identity $_.upn} |select-object UserPrincipalName, DisplayName, onpremsipaddress,sipaddress,Enabled,TeamsUpgradeEffectiveMode,EnterpriseVoiceEnabled, TeamsMeetingPolicy, HostedVoiceMail,OnlineVoiceRoutingPolicy, LineURI, OnPremLineURI, OnlineDialinConferencingPolicy,HostingProvider, InterpretedUserType|export-csv -notype $OutputCsvPath
    
                    } else {
                        #This sesession will enable Teams Caller ID for one user
                        $UsersEmail = Read-Host -prompt "Please enter user email address to enable Caller ID"
                        Grant-CsCallingLineIdentity -Identity $_.upn -PolicyName "policyname Contact Center Solutions Support"
                    }
                    }
                Default {
                    #default
                    Write-Host "`nInvalid choice, please select 1, 2, or 3." -ForegroundColor Red
                }
            }
                
        }

        '8'{
            Write-Host "Exiting..." -ForegroundColor Red
            exit
        }
        default {
            Write-Host " "
            Write-Host "Invalid choice, please select 1, 2, 3, 4, 5, 6, 7 or 8." -ForegroundColor Red
        }
    }
    # Pause before showing the menu again
    if($selection -ne '8') {
        Write-Host " "
        Write-Host "Press any key to return to the main menu..." -ForegroundColor DarkMagenta
        Write-Host " "
        $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }

}while($selection -ne '8')