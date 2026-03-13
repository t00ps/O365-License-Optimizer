# Import the Active Directory module
Import-Module ActiveDirectory

# Define the inactivity threshold (in months)
$inactivityMonths = 3  # Change this value to any number of months for the inactivity period

# Calculate the threshold date for inactivity (3 months ago)
$inactivityThresholdDate = (Get-Date).AddMonths(-$inactivityMonths)

# Define custom start and end dates for filtering (set your desired date range here)
$startDate = "1500-01-01"  # Set the start date here
$endDate = "2026-01-01"    # Set the end date here

# Convert the start and end dates to DateTime objects
$startDate = [datetime]$startDate
$endDate = [datetime]$endDate

# Set the threshold date for 1500-01-01
$thresholdDateForCreation = [datetime]"1500-01-01"

# Specify the Active Directory path (convert domain\ou\ou\ou\ou\ou to DN format)
$searchBase = "OU=,OU=,OU=,OU=,OU=,DC=,DC=,DC="

# Initialize the counter for inactive users
$userCount = 0
# Initialize the total months of inactivity
$totalInactiveMonths = 0

# Collect user inactivity data
$inactiveUsers = Get-ADUser -Filter * -SearchBase $searchBase -Properties lastLogonTimestamp, whenCreated | ForEach-Object {
    # Convert lastLogonTimestamp from a filetime to a readable datetime
    if ($_.lastLogonTimestamp -ne 0) {
        $lastLogonDate = [datetime]::FromFileTime($_.lastLogonTimestamp)
    } else {
        $lastLogonDate = $null
    }

    # Get the account creation date
    $accountCreationDate = $_.whenCreated

    # Check if the user is inactive for more than the specified months and if the inactivity period falls between the start and end dates
    if (($lastLogonDate -lt $inactivityThresholdDate -or $lastLogonDate -eq $null) -and 
        $lastLogonDate -ge $startDate -and $lastLogonDate -le $endDate) {

        # Increment the counter if the user is inactive and the logon date is between the start and end dates
        $userCount++

        # Handle "never logged in" case
        if ($lastLogonDate -eq [datetime]"1601-01-01 01:00:00 AM") {
            $lastLogonDateText = "never"
        } else {
            $lastLogonDateText = $lastLogonDate
        }

        # Determine which date to use (last logon or account creation date)
        if ($lastLogonDate -lt $thresholdDateForCreation -or $lastLogonDate -eq $null) {
            # If the last logon is before 01-01-2020 or never logged in, use the account creation date
            $inactiveDuration = (New-TimeSpan -Start $accountCreationDate -End (Get-Date)).Days / 30
        } else {
            # Otherwise, use the last logon date
            $inactiveDuration = (New-TimeSpan -Start $lastLogonDate -End (Get-Date)).Days / 30
        }

        # Handle the case for extremely high inactive months (like 5165 or more)
        if ($inactiveDuration -ge 5165) {
            $inactiveDurationText = "-"
        } else {
            $inactiveDurationText = [math]::Round($inactiveDuration)
        }

        # Accumulate the total number of months of inactivity
        $totalInactiveMonths += [math]::Round($inactiveDuration)

        # Output the user information
        [PSCustomObject]@{
            Name            = $_.Name
            SamAccountName  = $_.SamAccountName
            LastLogon       = $lastLogonDateText
            AccountCreation = $accountCreationDate
            InactiveMonths  = $inactiveDurationText  # Set inactive months as "-" if over 5165 months
        }
    }
} | Sort-Object LastLogon

# Generate the HTML table from the user data
$htmlTable = $inactiveUsers | ConvertTo-Html -Property Name, SamAccountName, LastLogon, AccountCreation, InactiveMonths -Head "<style>table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid black; padding: 8px; text-align: left; }</style>" -Body "<h2>Inactive User Report</h2><p>The following users have not logged in for more than $inactivityMonths months:</p>"

# Output the total number of inactive users and the total months of inactivity
$body = "<p>There are $userCount inactive users who have not logged in for more than $inactivityMonths months between $startDate and $endDate. The total months of inactivity across all users is $totalInactiveMonths.</p>"

# Append the HTML table to the body
$body += $htmlTable

# Define the email parameters
$from = "mailfrom@domain.com"  # Set the sender email address
$to = "mailto@domain.com"  # Set the recipient email group address
$subject = "Inactive User Report"
$smtpServer = "mail.xy"  # Set your SMTP server

# Send the email with HTML content
Send-MailMessage -From $from -To $to -Subject $subject -Body $body -SmtpServer $smtpServer -BodyAsHtml
