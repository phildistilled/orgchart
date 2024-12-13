# Filename: GenerateAsciiOrgChart.ps1

# Define the input CSV file path
$csvFile = "enrichedpeople.csv"

# Check if the input CSV file exists
if (-not (Test-Path -Path $csvFile)) {
    Write-Error "Input CSV file '$csvFile' does not exist. Please provide a valid file path."
    exit
}

# Import CSV data
try {
    $people = Import-Csv -Path $csvFile
    Write-Host "Successfully imported '$csvFile'. Total entries: $($people.Count)"
} catch {
    Write-Error "Failed to import CSV file '$csvFile'. Error: $_"
    exit
}

# Create a hashtable to map Email to person object for quick lookup
$emailToPerson = @{}
foreach ($person in $people) {
    $email = $person.Email.Trim().ToLower()
    $emailToPerson[$email] = $person
}

# Build a dictionary to map ManagerEmail to their direct reports
$managerToReports = @{}
foreach ($person in $people) {
    $managerEmail = $person.ManagerEmail.Trim().ToLower()
    $employeeEmail = $person.Email.Trim().ToLower()
    
    if (![string]::IsNullOrWhiteSpace($managerEmail)) {
        if ($managerEmail -eq $employeeEmail) {
            # ManagerEmail points to themselves; treat as top-level manager
            continue
        }

        if ($emailToPerson.ContainsKey($managerEmail)) {
            if ($managerToReports.ContainsKey($managerEmail)) {
                $managerToReports[$managerEmail] += $person
            } else {
                $managerToReports[$managerEmail] = @($person)
            }
        }
        else {
            # ManagerEmail points to someone not in the list; treat the employee as top-level manager
            # We'll identify top-level managers in the next step
            continue
        }
    }
}

# Identify top-level managers (those whose ManagerEmail does not match any Email in the CSV)
$topLevelManagers = $people | Where-Object { 
    [string]::IsNullOrWhiteSpace($_.ManagerEmail) -or 
    -not $emailToPerson.ContainsKey($_.ManagerEmail.Trim().ToLower())
}

# Debugging: Check if any top-level managers are found
if ($topLevelManagers.Count -eq 0) {
    Write-Warning "No top-level managers found (employees without a ManagerEmail or with ManagerEmail not in the CSV). Please check your CSV data."
} else {
    Write-Host "Identified $($topLevelManagers.Count) top-level manager(s)."
}

# Function to recursively build the hierarchy as strings
function Get-OrgChartString {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Person,
        
        [int]$Level = 0,
        [bool]$IsLast = $false,
        [ref]$VisitedEmails
    )
    
    # Initialize the output string
    $output = ""
    
    # Define indentation based on level
    $indent = ""
    for ($i = 0; $i -lt $Level; $i++) {
        if ($VisitedEmails.Value[$i]) {
            $indent += "    "
        } else {
            $indent += "│   "
        }
    }
    
    # Define branch lines
    if ($Level -gt 0) {
        if ($IsLast) {
            $indent += "└── "
        } else {
            $indent += "├── "
        }
    }
    
    # Create the line for the current person
    $line = "$indent$($Person.GivenName) $($Person.Surname) [$($Person.JobTitle)]"
    $output += "$line`n"
    
    # Mark this level as having the last entry if applicable
    if ($Level -ge $VisitedEmails.Value.Count) {
        $VisitedEmails.Value += $IsLast
    } else {
        $VisitedEmails.Value[$Level] = $IsLast
    }
    
    # Check for circular references
    $personEmail = $Person.Email.Trim().ToLower()
    if ($VisitedEmails.Value.Contains($personEmail)) {
        # Indicate a circular reference
        $output += "$indent[Circular Reference Detected for $($Person.GivenName) $($Person.Surname)]`n"
        Write-Warning "Circular reference detected for $($Person.GivenName) $($Person.Surname)"
        return $output
    }
    
    # Add the current person's email to the visited list
    $VisitedEmails.Value += $personEmail
    
    # If the person has direct reports, sort them and recurse
    if ($managerToReports.ContainsKey($personEmail)) {
        $directReports = $managerToReports[$personEmail] | Sort-Object GivenName
        $count = $directReports.Count
        for ($j = 0; $j -lt $count; $j++) {
            $report = $directReports[$j]
            $isLast = ($j -eq ($count - 1))
            $output += Get-OrgChartString -Person $report -Level ($Level + 1) -IsLast $isLast -VisitedEmails ([ref]$VisitedEmails.Value)
        }
    }
    
    return $output
}

# Initialize the VisitedEmails list to track visited employees and prevent circular references
$VisitedEmails = @()

# Build the org chart string
$chartOutput = ""
foreach ($manager in $topLevelManagers | Sort-Object GivenName) {
    $chartOutput += Get-OrgChartString -Person $manager -Level 0 -IsLast $false -VisitedEmails ([ref]$VisitedEmails)
}

# Check if chartOutput is empty
if ([string]::IsNullOrWhiteSpace($chartOutput)) {
    Write-Warning "Organizational chart is empty. Please verify your CSV data."
} else {
    # Output the chart to the console
    Write-Host "`nOrganizational Chart:`n" -ForegroundColor Cyan
    Write-Host $chartOutput
    
    # Export the chart to a text file
    $outputFile = "OrgChart.txt"
    try {
        $chartOutput | Out-File -FilePath $outputFile -Encoding UTF8
        Write-Host "Organizational chart exported to '$outputFile'."
    } catch {
        Write-Warning "Failed to export organizational chart to '$outputFile'. Error: $_"
    }
}
