<#
.NOTES
    Name: Create_remote_shared_mailbox.ps1
    Author: Bela Bana | https://github.com/belabana
    Request: I needed to automate the creation of a new site with a remote shared mailbox in a hybrid Exchange environment.
    Classification: Public
    Disclaimer: Author does not take responsibility for any unexpected outcome that could arise from using this script.
                Please always test it in a virtual lab or UAT environment before executing it in production environment.
    Variables: You need to set your servers where I used "FQDNOFYOURONPREMISEEXCHANGESERVER", "YOURPRIMARYDOMAINCONTROLLER", "YOURSECONDARYDOMAINCONTROLLER", "YOURADCONNECTSERVER" references.    

.SYNOPSIS
    Create a new site with a remote shared mailbox.

.DESCRIPTION
    This script will reduce the time required for service desk engineers to process new site creation requests.
    It automates a list of steps to create the site in Active Directory, onpremise Exchange and Exchange Online.
    Transcript and timestamps are also added to calculate runtime and write output to a .log file. Path: “C:\temp”
    It should be executed from onpremise Exchange server with a user who is an onpremise Exchange and Exchange Online administrator.

.PARAMETER Site
    The 4-digit site number which will be assigned to the new site.

.PARAMETER SiteName
    The name of the site to be shown in the global address list.

.PARAMETER EmailAddress
    The desired email address required to be created for the new site.

.PARAMETER ADAdminCredential
    Credentials of a user with Domain and Exchange administrator access.
#>
param(
    [parameter(Mandatory=$true)]
    [System.String] $Site,

    [parameter(Mandatory=$true)]
    [System.String] $SiteName,

    [parameter(Mandatory=$true)]
    [System.String] $EmailAddress,

    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $ADAdminCredential
)
begin {
$TranscriptPath = "C:\temp\Create_remote_shared_mailbox_$(Get-Date -Format yyyy-MM-dd-HH-mm).log"
Start-Transcript -Path $TranscriptPath
Write-Host -ForegroundColor Yellow "Script started: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
#Variables
[int]$StartTimer = (Get-Date).Second
$ExchangeServerFQDN = "FQDNOFYOURONPREMISEEXCHANGESERVER"
$DCControllers = "YOURPRIMARYDOMAINCONTROLLER","YOURSECONDARYDOMAINCONTROLLER"
$AADServer = "YOURADCONNECTSERVER"

#Connect to onpremise Exchange server
Function Connect-ExchangeOnprem ($ExchangeServerFQDN, $ADAdminCredential) {
    Write-Host -ForegroundColor Yellow "Connecting to onpremise Exchange server.."
    $so = New-PSSessionOption -SkipCACheck:$true -SkipCNCheck:$true -SkipRevocationCheck:$true
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServerFQDN/PowerShell/ -Authentication Kerberos -Credential $ADAdminCredential -SessionOption $so
    try {       
        Import-PSSession $Session -AllowClobber -DisableNameChecking | Out-Null
        Write-Host -ForegroundColor Green "Connected to onpremise Exchange server."
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to connect to onpremise Exchange server. Terminating process."
        Write-Error -Message "$_" -ErrorAction Stop
        return;
    }
}
#Replicate Active Directory
Function Replicate-AD ([string[]]$DCControllers,$ADAdminCredential) { 
    Write-Host -ForegroundColor Yellow "Replicating Active Directory.."          
    ForEach ($Controller in $DCControllers) { 
        $Controller
        $x=0
        Start-Sleep -Seconds 3
        While ($x -lt 3) { 
            $x = $x+1
            Invoke-Command -ComputerName $Controller -Credential $ADAdminCredential -ScriptBlock { cmd /c "repadmin /syncall /AdeP" }
            Start-Sleep -Seconds 5
        }
    }
    Write-Host -ForegroundColor Green "AD replication completed."
}
#Sync to O365
Function Sync-AD ($AADServer,$ADAdminCredential) {
    Write-Host -ForegroundColor Yellow "Running AD Delta Sync.."
    Start-Sleep -Seconds 15 #Give onpremise Exchange some time before sync 
    Invoke-Command -ComputerName $AADServer -Credential $ADAdminCredential -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
    Start-Sleep -Seconds 15 #Allow O365 web GUI to pickup the changes after sync
    Write-Host -ForegroundColor Green "AD Delta Sync completed."
}
#Confirm if site exists in AD
$SiteExists = $null
if ($Site) {
    Write-Host -ForegroundColor Yellow "Checking if public folder exists in Exchange.."
    try {
        $SiteExists = Get-MailPublicFolder -Identity \Site\$Site -ErrorAction Stop #Please check path to sites in your Exchange environemnt
        Write-Host -ForegroundColor Red "Site number: $Site is already taken."
    }
    catch {
        Write-Host -ForegroundColor Green "Site number: $Site is available."
    }
}
}
process {
if ($SiteExists -ne $null) {
    Write-Host -ForegroundColor Red "Terminating process."
    Exit
}
else {
    #Create Remote Shared Mailbox onpremise Exchange
    Connect-ExchangeOnprem -ExchangeServerFQDN $ExchangeServerFQDN -ADAdminCredential $ADAdminCredential
    Write-Host -ForegroundColor Yellow "Creating a remote shared mailbox for the site.."
    try {
        $Global:ErrorActionPreference = 'Stop'
        New-RemoteMailbox -Shared -Name $SiteName -UserPrincipalName $EmailAddress | Out-Null
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to create remote shared mailbox. Terminating process."
        Write-Error -Message "$_" -ErrorAction Stop
        return;
    }
    #Replicate Active Directory (optional, currently disabled)
    #Replicate-AD -Controllers $Controllers -ADAdminCredential $ADAdminCredential | out-null
    #Write-Host -ForegroundColor Yellow "Waiting for AD to replicate.."
    #Sync changes to O365
    Sync-AD -AADServer $AADServer -ADAdminCredential $ADAdminCredential  | out-null
}
}
end {
#Confirm if site email address has been created
$SiteExists = $null
if ($Site) {
    Write-Host -ForegroundColor Yellow "Checking if site email address has been created.."
    try {
        $SiteExists = Get-RemoteMailbox -Identity $EmailAddress -ErrorAction Stop | out-null
        Write-Host -ForegroundColor Green "Site email address has been found."
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to find email address for the new site."
    }
}
if ($Session -ne $null) {
    Remove-PSSession $Session;
}
Write-Host -ForegroundColor Yellow "Script ended: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
[int]$EndTimer = (Get-Date).Second
Write-Host -ForegroundColor Yellow "Script completed in $([Math]::Abs($StartTimer - $EndTimer)) seconds."
Stop-Transcript
}
