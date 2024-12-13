# Define input and output file paths
$inputFile = "people.txt"
$outputFile = "enrichedpeople.csv"

# Function to perform LDAP query and retrieve user properties
function Get-ADUserProperties {
    param (
        [string]$Email
    )

    # Initialize properties
    $Title = ""
    $ManagerName = ""
    $ManagerEmail = ""

    try {
        # Create a DirectoryEntry object (root of the domain)
        $root = New-Object System.DirectoryServices.DirectoryEntry

        # Create a DirectorySearcher object
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = $root
        $searcher.Filter = "(&(objectClass=user)(mail=$Email))"
        $searcher.PropertiesToLoad.Add("title") | Out-Null
        $searcher.PropertiesToLoad.Add("manager") | Out-Null

        # Perform the search
        $result = $searcher.FindOne()

        if ($result -ne $null) {
            # Retrieve Title
            if ($result.Properties["title"]) {
                $Title = $result.Properties["title"][0]
            }

            # Retrieve Manager's Distinguished Name (DN)
            if ($result.Properties["manager"]) {
                $managerDN = $result.Properties["manager"][0]

                # Initialize DirectoryEntry for Manager
                $managerEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$managerDN")

                # Retrieve Manager's Display Name
                if ($managerEntry.Properties["displayName"].Count -gt 0) {
                    $ManagerName = $managerEntry.Properties["displayName"][0]
                }

                # Retrieve Manager's Email
                if ($managerEntry.Properties["mail"].Count -gt 0) {
                    $ManagerEmail = $managerEntry.Properties["mail"][0]
                }
            }
        } else {
            Write-Warning "No AD user found with email: $Email"
        }
    } catch {
        Write-Warning "AD lookup failed for email: $Email. $_"
    }

    return @{
        Title        = $Title
        ManagerName  = $ManagerName
        ManagerEmail = $ManagerEmail
    }
}

# Check if the input file exists
if (-not (Test-Path -Path $inputFile)) {
    Write-Error "Input file '$inputFile' does not exist. Please provide a valid file path."
    exit
}

# Read all lines from the input text file
try {
    $lines = Get-Content -Path $inputFile
} catch {
    Write-Error "Failed to read input file '$inputFile'. Ensure the file is accessible and not in use."
    exit
}

# Prepare an array to hold enriched data
$enrichedData = @()

# Define the regex pattern to parse each line
$regex = '^(?<Surname>[^,]+),\s*(?<GivenName>[^\s<]+)\s*<(?<Email>[^>]+)>$'

# Iterate through each line in the text file
foreach ($line in $lines) {
    # Trim any leading/trailing whitespace
    $originalEntry = $line.Trim()

    # Skip empty lines
    if ([string]::IsNullOrWhiteSpace($originalEntry)) {
        continue
    }

    # Use regex to parse the entry
    $matches = [regex]::Match($originalEntry, $regex)

    if ($matches.Success) {
        $surname = $matches.Groups['Surname'].Value.Trim()
        $givenName = $matches.Groups['GivenName'].Value.Trim()
        $email = $matches.Groups['Email'].Value.Trim()

        # Perform AD lookup using LDAP
        $adProperties = Get-ADUserProperties -Email $email
        $jobTitle = $adProperties.Title
        $managerName = $adProperties.ManagerName
        $managerEmail = $adProperties.ManagerEmail

        # Create a custom object with the enriched data
        $enrichedEntry = [PSCustomObject]@{
            Surname        = $surname
            GivenName      = $givenName
            Email          = $email
            JobTitle       = $jobTitle
            ManagerName    = $managerName
            ManagerEmail   = $managerEmail
        }

        # Add the enriched entry to the array
        $enrichedData += $enrichedEntry
    } else {
        Write-Warning "Failed to parse entry: '$originalEntry'"
    }
}

# Export the enriched data to a CSV file
try {
    $enrichedData | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
    Write-Host "Enriched data successfully exported to '$outputFile'"
} catch {
    Write-Error "Failed to export enriched data to CSV. $_"
}
