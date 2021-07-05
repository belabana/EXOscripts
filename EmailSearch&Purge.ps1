<#
.NOTES
    ***I AM STILL WORKING ON THIS SCRIPT, IT IS PARTIALLY TESTED. PLEASE WAIT UNTIL NEXT VERSION.***
    Name: EmailSearch&Purge.ps1
    Author: Bela Bana | https://github.com/belabana
    Request: I needed to provide my colleagues with an informative script which is able to purge a malicious email from Exchange Online mailboxes.
    Classification: Public
    Disclaimer: Author does not take responsibility for any unexpected outcome that could arise from using this script.
                Please always test it in a virtual lab or UAT environment before executing it in production environment.
        
.SYNOPSIS
    Search and purge an exact email from all mailboxes in Exchange Online.

.DESCRIPTION
    This script will help you find an email by different criteria and purge it from all users mailbox.
    Transcript and timestamps are added to calculate runtime and write output to a .log file. Path: "C:\temp"
    It needs to be executed from a computer with Exchange Online PowerShell V2 module with either a Global or Exchange Administrator account.

.PARAMETER IPPSAdminCredential
    Credentials of a user with Global or Exchange Online Administrator access. MFA should be disabled for the user.

.PARAMETER SearchName
    The name of your search. This is how your email search will be shown up in IPPS.

.PARAMETER Sender
    The sender of the email which you need to search and purge.

.PARAMETER Subject
    The subject of the email which you need to search and purge.

.PARAMETER ReceiveDate
    The date when the email was delivered.
    Syntax reference for exact date: 5/11/2021
    Syntax reference for date interval: 6/18/2021..3/18/2021
#>
param (
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $IPPSAdminCredential,

    [parameter(Mandatory=$true)]
    [System.String] $SearchName = $( Read-Host "Enter the name of your search i.e. 'Find phishing emails from John Doe' - mandatory" ),

    [parameter(Mandatory=$true)]
    [System.String] $Sender = $( Read-Host "Enter the sender's email address i.e. 'john.doe@phishingstar.com' - mandatory" ),

    [parameter(Mandatory=$false)]
    [System.String] $Subject = $( Read-Host "Enter the subject line of the targeted email i.e. 'Wedding invitation' - optional" ),

    [parameter(Mandatory=$false)]
    [System.String] $ReceiveDate = $( Read-Host "Enter the exact date of email delivery i.e. '5/11/2021' or interval '6/18/2021..3/18/2021' - optional" )
)
begin {
$TranscriptPath = "C:\temp\EmailSearch&Purge_$(Get-Date -Format yyyy-MM-dd-HH-mm).log"
Start-Transcript -Path $TranscriptPath
Write-Host -ForegroundColor Yellow "Script started: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
#Functions
Function GetComplianceSearch {
    do {
        $SearchResults = Get-ComplianceSearch $SearchName # or -Identity $Search.Identity
        if ($SearchResults.Status -ne "Completed") {
            Write-Host -ForegroundColor Yellow "Searching is still in progress. Please wait.."
            Start-Sleep -Seconds 15
        }
    } until ($SearchResults.Status -eq "Completed")
    Write-Host -ForegroundColor Green "Searching has been completed."
    Write-Host -ForegroundColor Yellow "There are" + $SearchResults.Items + "emails that match the specified criteria."
}
Function AddNameToSearchAction {
    param(
        [Parameter(Mandatory)]
        [string]$SearchActionName = $( Read-Host "Enter the name for compliance search action i.e. '<Ticket ref>-Purge malicious emails' - mandatory" )
    )
}
#Request confirmation to connect to the Security & Compliance Center
$Username = $IPPSAdminCredential.UserName
Write-Warning "Would you like to connect to the Security & Compliance Center with the following user: `n$UserName" -WarningAction Inquire
}
process {
#Start a timer to measure runtime
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
#Connect to IPPS
try {
    Connect-IPPSSession -Credential $IPPSAdminCredential
    Write-Host -ForegroundColor Green "Connected to Security & Compliance Center."
}
catch {
    Write-Host -ForegroundColor Red "Unable to connect to Security & Compliance Center. Terminating process."
    Write-Error -Message "$_" -ErrorAction Stop
    return;
}
#Search for the email by the specified criteria
if ($Subject -ne $null) {
    if ($ReceiveDate -ne $null) {
        $Search = New-ComplianceSearch -Name $SearchName -ExchangeLocation All -ContentMatchQuery '(Received:$ReceiveDate) AND (Subject:$Subject) AND (From:$Sender)'
    }
    else {
        $Search = New-ComplianceSearch -Name $SearchName -ExchangeLocation All -ContentMatchQuery '(Subject:$Subject) AND (From:$Sender)'
    }
}
else {
    $Search = New-ComplianceSearch -Name $SearchName -ExchangeLocation All -ContentMatchQuery '(From:$Sender)' 
 }
#Start compliance search
Write-Host -ForegroundColor Yellow "Starting a compliance search with the specified criteria.." 
try {
    Start-ComplianceSearch -Identity $Search.Identity
    Start-Sleep -Seconds 15
}
catch {
    Write-Host -ForegroundColor Red "Unable to start compliance search with the given criteria. Terminating process."
    Write-Error -Message "$_" -ErrorAction Stop
    return;
}
#Get number of items before purge
GetComplianceSearch
Write-Warning "Would you like to purge them from the mailboxes?" -WarningAction Inquire
#Purge emails that meet the criteria
Write-Host -ForegroundColor Yellow "Attempting to purge emails.."
AddNameToSearchAction
try {
    New-ComplianceSearchAction -SearchName $SearchActionName -Purge -PurgeType HardDelete
} 
catch {
    Write-Host -ForegroundColor Red "Unable to purge emails. Terminating process."
    Write-Error -Message "$_" -ErrorAction Stop
    return;
}
#Check status of compliance action
do {
    $PurgeResults = Get-ComplianceSearchAction -SearchName $SearchActionName
    if ($PurgeResults.Status -ne "Completed") {
         Write-Host -ForegroundColor Yellow "Purging is still in progress. Please wait.."
         Start-Sleep -Seconds 15
    }
} until ($PurgeResults.Status -eq "Completed")
Write-Host -ForegroundColor Green "Purging has been completed."
#Start another compliance search to verify results
Write-Host -ForegroundColor Yellow "Starting a new compliance search to ensure zero results.." 
try {
    Start-ComplianceSearch -Identity $Search.Identity
    Start-Sleep -Seconds 15
}
catch {
    Write-Host -ForegroundColor Red "Unable to start compliance search to verify zero results after purge. Please check manually."
    Write-Error -Message "$_" -ErrorAction Continue
    return;
}
#Get number of items after purge
GetComplianceSearch
#Disconnect from EXO
Disconnect-ExchangeOnline
Write-Host -ForegroundColor Yellow "Script ended: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
$StopWatch.Stop()
Write-Host -ForegroundColor Yellow "Script completed in" $Stopwatch.Elapsed.Hours "hours and" $Stopwatch.Elapsed.Minutes "minutes and" $Stopwatch.Elapsed.Seconds "seconds."
Stop-Transcript
}
