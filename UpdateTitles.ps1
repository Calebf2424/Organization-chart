function Get-UsersFromOrgChart {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$File
    )

    # Import the Excel module
    Import-Module ImportExcel

    # Read the Excel file, treating the second row as the header
    $excelData = Import-Excel -Path $File -StartRow 2

    # Initialize an ArrayList to hold the user objects
    $users = New-Object System.Collections.ArrayList
    $managerCache = @{}  # Hashtable for caching known manager emails

    # First pass: Create the manager cache
    foreach ($row in $excelData) {
        $managerName = $row.'Name'
        $managerEmail = ''

        if ($managerName -and -not $managerCache.ContainsKey($managerName)) {
            # Check if "Primary Email Address" exists and is not null
            if (-not [string]::IsNullOrWhiteSpace($row.'Primary Email Address')) {
                $managerEmail = $row.'Primary Email Address'.ToLower()
            }

            # Add to cache if the email is valid
            if ($managerEmail -ne '' -and $managerEmail.EndsWith('@cityofgp.com')) {
                $managerCache[$managerName] = $managerEmail
            }
        }
    }

    # Second pass: Process the users
    foreach ($row in $excelData) {
        # Check if "Primary Email Address" exists and is not null
        if (-not [string]::IsNullOrWhiteSpace($row.'Primary Email Address')) {
            $UPN = $row.'Primary Email Address'.ToLower()
        } else {
            $UPN = ''
        }

        # Only proceed if UPN is not empty and ends with "@cityofgp.com"
        if ($UPN -ne '' -and $UPN.EndsWith('@cityofgp.com')) {
            # Find the manager's email address from the cache
            $managerName = $row.'Manager Name'
            $DirectReportUPN = ''
            if ($managerName -and $managerCache.ContainsKey($managerName)) {
                $DirectReportUPN = $managerCache[$managerName]
            }

            # Extract the title from "Business Title" up to the first underscore
            $businessTitle = $row.'Business Title'
            $Title = ''
            $Department = ''

            if ($businessTitle) {
                $splitTitle = $businessTitle.Split('_')

                # Extract Title and Department based on split
                if ($splitTitle.Count -ge 2) {
                    $Title = $splitTitle[0].Trim()
                    $Department = $splitTitle[1].Trim()
                } elseif ($splitTitle.Count -eq 1) {
                    $Title = $splitTitle[0].Trim()
                    $Department = ''
                }
            }

            # Create a PSCustomObject for the user
            $user = [PSCustomObject]@{
                UPN             = $UPN
                DirectReportUPN = $DirectReportUPN
                Title           = $Title
                Department      = $Department
            }

            # Add the user object to the ArrayList
            $users.Add($user) | Out-Null
        }
    }

    # Return the ArrayList of user objects
    return $users
}


function Get-ToBeUpdatedTitles {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$OrgChart,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$AdUsers
    )

    # Import the AD CSV file
    $adData = Import-Csv -Path $AdUsers

    # Create a hashtable for faster lookups
    $adLookup = @{}
    foreach ($adUser in $adData) {
        if ($adUser.UPN) {
            $adLookup[$adUser.UPN] = $adUser
        }
    }

    # Creates an array of users using OrgChart and Get-UsersFromOrgChart func
    $users = Get-UsersFromOrgChart -File $OrgChart
    $updated = New-Object System.Collections.ArrayList

    foreach($user in $users) {
        # Use the hashtable for a quick lookup
        if ($adLookup.ContainsKey($user.UPN)) {
            $adUser = $adLookup[$user.UPN]

            # If titles do not match, add to the updated array
            if ($adUser.Title -ne $user.Title) {
                $toAdd = [PSCustomObject]@{
                    UPN      = $user.UPN
                    OldTitle = $adUser.Title
                    NewTitle = $user.Title
                }

                # Add user to the to-be-updated array
                $updated.Add($toAdd) | Out-Null
            }
        }
    }

    # Return the array of users with mismatched titles
    return $updated
}
