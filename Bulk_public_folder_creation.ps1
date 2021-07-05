<#
.NOTES
    Name: Bulk_public_folder_creation.ps1
    Author: Bela Bana | https://github.com/belabana
    Request: I needed to automate the creation of new sites with mail enabled public folders in a hybrid Exchange environment.
    Classification: Public
    Disclaimer: Author does not take responsibility for any unexpected outcome that could arise from using this script.
                Please always test it in a virtual lab or UAT environment before executing it in production environment.
    Variables: You need to set your server and database where I used "FQDNOFYOURONPREMISEEXCHANGESERVER" and "YOURMAILBOXDATABASE" references.    

.SYNOPSIS
    Create multiple sites with mail enabled public folders using a CSV file.

.DESCRIPTION
    This script will reduce the time required for service desk engineers to process new site creation requests.
    It processes a comma-separated values (CSV) file called "Bulk_public_folder_creation.csv" in which the required new sites are listed.
    The CSV file is mandatory and it has to be in the folder with path "C:\temp".
    The script automates a list of steps to create a mail enabled public folder for each site.
    Transcript and timestamps are also added to calculate runtime and write output to a .log file. Path: "C:\temp"
    It should be executed with credentials of an Exchange Administrator from the onpremise Exchange server.

.PARAMETER ADAdminCredential
    Credentials of a user with Domain and Exchange administrator access.
#>
param(
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $ADAdminCredential
)
begin {
$TranscriptPath = "C:\temp\Bulk_public_folder_creation_$(Get-Date -Format yyyy-MM-dd-HH-mm).log"
Start-Transcript -Path $TranscriptPath
Write-Host -ForegroundColor Yellow "Script started: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
#Variables
[int]$StartTimer = (Get-Date).Second
$file = "C:\temp\Bulk_public_folder_creation.csv"
$ExchangeServerFQDN = "FQDNOFYOURONPREMISEEXCHANGESERVER"
$Global:ErrorActionPreference = 'Stop'

#Confirm if Bulk_public_folder_creation.csv file exists
Write-Host -ForegroundColor Yellow "Checking if Bulk_public_folder_creation.csv file exists.."
if (Test-Path -Path $file -PathType Leaf) {
    Write-Host -ForegroundColor Green "File [$file] was found."
}
else {
    Write-Host -ForegroundColor Red "The file [$file] was not found. Terminating process."
    Exit
}
#Functions
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
        Write-Error -Message "$_"
        return;
    }
}
}
process {
#Importing sites from Bulk_public_folder_creation.csv file
Write-Host -ForegroundColor Yellow "Importing sites from file.."
$Sitelist = Import-CSV $file
#Connecting to Exchange server
Connect-ExchangeOnprem -ExchangeServerFQDN $ExchangeServerFQDN -ADAdminCredential $ADAdminCredential
#Creating sites with mail enabled public folders
Write-Host -ForegroundColor Yellow "Creating sites with public folders.."

ForEach ($Site in $Sitelist) {
    $Sitenumber = $Site.SiteNumber
    Write-Host $Sitenumber
    if ($Sitenumber) {
        Write-Host -ForegroundColor Yellow "Checking if sitenumber [$Sitenumber] is available.."
        try {
            $SiteExists = Get-MailPublicFolder -Identity \Site\$Sitenumber #Please check path to sites in your Exchange environemnt
            Write-Host -ForegroundColor Red "Site number [$Sitenumber] is already taken."
        }
        catch {
            Write-Host -ForegroundColor Green "Site number [$Sitenumber] is available."
            try {
                New-PublicFolder -Name "$Sitenumber" -Path \Site -Mailbox YOURMAILBOXDATABASE | Out-Null
                Enable-MailPublicFolder -Identity "\Site\$Sitenumber"
                Write-Host -ForegroundColor Green "Public folder for site [$Sitenumber] has been created."
            }
            catch { 
                Write-Host -ForegroundColor Red "Unable to create public folder for site [$Sitenumber]. Terminating process."
            }
        }
    }
}
}
end {
#Confirm if mail enabled public folders have been created
$SiteExists = $null
ForEach ($Site in $Sitelist) {
    $Sitenumber = $Site.SiteNumber
    if ($Sitenumber) {
        Write-Host -ForegroundColor Yellow "Checking if public folder for site [$Sitenumber] is created.."
        try {
            $SiteExists = Get-MailPublicFolder -Identity \Site\$Sitenumber -ErrorAction Stop | out-null
            Write-Host -ForegroundColor Green "Public folder for site [$Sitenumber] was found."
        }
        catch {
            Write-Host -ForegroundColor Red "Unable to find public folder for site [$Sitenumber]."
        }
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
