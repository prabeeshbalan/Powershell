# Import the Active Directory module
Import-Module ActiveDirectory

# Define an array of hashtables, each containing domain information
$Domains = @(
    @{
        Name = "Domain1"
        DomainController = "harrisbank.bmogc.net"
    },
    @{
        Name = "Domain2"
        DomainController = "ibg.adroot.bmogc.net"
    },
    @{
        Name = "Domain3"
        DomainController = "nesbittburns.ca"
    },
    @{
        Name = "Domain4"
        DomainController = "office.adroot.bmogc.net"
    },
    @{
        Name = "Domain5"
        DomainController = "pcd.nesbittburns.ca"
    },
    @{
        Name = "Domain6"
        DomainController = "percom.adroot.bmogc.net"
    }
)

# Prompt for the path to the CSV file (only once)
$CSVPath = Read-Host -Prompt "Enter the path to the CSV file containing group names"

# Import the CSV file
$Groups = Import-Csv -Path $CSVPath

# Function to process each domain
function Process-Domain {
    param(
        [string]$DomainName,
        [string]$DomainController
    )

    # Create an empty array to store the group information for the current domain
    $DomainGroupInfo = @()

    # Loop through each group in the CSV file
    foreach ($Group in $Groups) {
        # Get the group name from the "Group Name" column
        $GroupName = $Group."Group Name"

        # Search for the group
        try {
            # Get-ADGroup with -Properties Members is the key change
            $ADGroup = Get-ADGroup -Filter "name -like '$GroupName'" -Server $DomainController -Properties Description, ManagedBy, DistinguishedName, Members -ErrorAction Stop

            # Get the member count from the Members property
            $MemberCount = ($ADGroup.Members).Count

            # Create an object to store the group information
            $GroupObj = [PSCustomObject]@{
                Domain          = $DomainName
                Name            = $ADGroup.Name
                Description     = $ADGroup.Description
                ManagedBy       = $ADGroup.ManagedBy
                DistinguishedName = $ADGroup.DistinguishedName
                MemberCount     = $MemberCount
            }

            # Add the group object to the domain-specific array
            $DomainGroupInfo += $GroupObj
        }
        catch {
            Write-Verbose "Group '$GroupName' not found in $DomainName."
            # Handle the case where Members is null (group has no members)
            $MemberCount = 0
             $GroupObj = [PSCustomObject]@{
                Domain          = $DomainName
                Name            = $GroupName
                Description     = $null
                ManagedBy       = $null
                DistinguishedName = $null
                MemberCount     = $MemberCount
            }
           $DomainGroupInfo += $GroupObj

        }
    }

    return $DomainGroupInfo
}

# Create the output folder with an incrementing number
$OutputFolderBase = "Output-"
$OutputFolderNumber = 01
$OutputFolder = "{0}{1:d2}" -f $OutputFolderBase, $OutputFolderNumber
while (Test-Path $OutputFolder) {
    $OutputFolderNumber++
    $OutputFolder = "{0}{1:d2}" -f $OutputFolderBase, $OutputFolderNumber
}
New-Item -ItemType Directory -Path $OutputFolder

# Create an empty array to store all group information
$AllGroupInfo = @()

# Loop through each domain and process it
foreach ($Domain in $Domains) {
    $DomainName = $Domain.Name
    $DomainController = $Domain.DomainController

    # Process the domain and get the group information
    $DomainGroupInfo = Process-Domain -DomainName $DomainName -DomainController $DomainController

    # Add the domain-specific group info to the overall array
    $AllGroupInfo += $DomainGroupInfo
}

# Filter the group information to include only entries with a value in the "Name" field
$FilteredGroupInfo = $AllGroupInfo | Where-Object { $_.Name -ne $null }

# Export the filtered group information to a CSV file in the output folder
$CSVOutputPath = Join-Path -Path $OutputFolder -ChildPath "AllDomains.csv"
$FilteredGroupInfo | Export-Csv -Path $CSVOutputPath -NoTypeInformation

# --- Group Member Extraction ---

# Modified Get-ADGroupMemberFix function
Function Get-ADGroupMemberFix {
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0
        )]
        [string[]]
        $Identity,

        [Parameter(Mandatory = $true)]  # DomainController is now MANDATORY
        [string]
        $DomainController
    )
    process {
        foreach ($GroupIdentity in $Identity) {
            $Group = $null
            # Use the provided $DomainController with port 3268
            $Group = Get-ADGroup -Identity $GroupIdentity -Properties Member -Server "$DomainController`:3268"
            if (-not $Group) {
                Write-Verbose "Group '$GroupIdentity' not found on $DomainController."
                continue
            }
            Foreach ($Member in $Group.Member) {
                try {
                    $ADObject = Get-ADObject $Member -Server "$DomainController`:3268" -Properties UserPrincipalName, samaccountname -ErrorAction Stop
                    if ($ADObject.ObjectClass -eq "user") {
                        [PSCustomObject]@{
                            SamAccountName  = $ADObject.samaccountname
                            UPN             = $ADObject.UserPrincipalName
                            Username        = "{0}\{1}" -f ($GroupIdentity -replace '.* - (.*?) - .*', '$1'), $ADObject.samaccountname # Extract domain from group name
                            DistinguishedName = $Member  # Use the member DN directly
                        }
                    }
                }
                catch {
                    Write-Verbose "Error retrieving member: $Member on $DomainController"
                }
            }
        }
    }
}


# Import the "AllDomains.csv" file
$AllDomainsData = Import-Csv -Path $CSVOutputPath

# Loop through each group in the "AllDomains.csv" file
foreach ($Group in $AllDomainsData) {
    $DomainName = $Group.Domain
    $GroupName = $Group.Name

    # Get the domain controller for the current domain.  REQUIRED now.
    $DomainController = $Domains | Where-Object { $_.Name -eq $DomainName } | Select-Object -ExpandProperty DomainController
    if (-not $DomainController) {
        Write-Warning "Domain controller not found for domain: $DomainName"
        continue  # Skip this group if no DC is found.
    }

    # Use the Get-ADGroupMemberFix function and pipe the output to Export-Csv
    Get-ADGroupMemberFix -Identity $GroupName -DomainController $DomainController | Export-Csv -Path "$OutputFolder\$DomainName-$GroupName.csv" -NoTypeInformation
}

Write-Host "Script completed.  Output files saved in '$OutputFolder'."