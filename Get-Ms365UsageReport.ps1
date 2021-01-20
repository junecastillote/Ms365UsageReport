
#Requires -Version 5.1
#Requires -PSEdition Desktop
#Requires -Modules @{ ModuleName="MSAL.PS"; ModuleVersion="4.16.0.4" }
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="2.0.3" }

<#PSScriptInfo

.VERSION 1.2.2

.GUID 19fea2a0-ff5a-4f00-8d15-4e721d5c3c7b

.AUTHOR June Castillote

.COMPANYNAME June Castillote

.COPYRIGHT june.castillote@gmail.com

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/Ms365UsageReport

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

Include:
* [√] Reports > Usage > Active Users
* [√] Reports > Usage > Activations > Users

Add Config item:
    * [√] tenantName
    * [√] msGraphAuthType
    * [√] msGraphAppID
    * [√] msGraphAppKey
    * [√] msGraphAppCertificateThumbprint
    * [√] exchangeAuthType
    * [√] exchangeAppID
    * [√] exchangeAppCertificateThumbprint
    * [√] exchangeCredentialFile

Changed:
    * [√] Revert Get-ExoMailbox to Get-Mailbox due to REST-related issues.

.PRIVATEDATA

#>

<#
.SYNOPSIS
    Short description
.DESCRIPTION
    Microsoft 365 Usage Reporting Script using Microsoft Graph API and Exchange Online PowerShell V2
.EXAMPLE
    PS C:\> .\Get-Ms365UsageReport.ps1 -Config .\config.json -GraphApiAccessToken <accesstoken>

.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>

[cmdletbinding()]
param (
    [Parameter(Mandatory)]
    [string]$Config
)
#Region Functions
Function LogEnd {
    $txnLog = ""
    Do {
        try {
            Stop-Transcript | Out-Null
        }
        catch [System.InvalidOperationException] {
            $txnLog = "stopped"
        }
    } While ($txnLog -ne "stopped")
}

Function LogStart {
    param (
        [Parameter(Mandatory = $true)]
        [string]$logPath
    )
    LogEnd
    Start-Transcript $logPath -Force | Out-Null
}
#EndRegion Functions

While (Get-PsSession -Name ExchangeOnline*) {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

LogEnd

#Enable TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$WarningPreference = "SilentlyContinue"
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition

# $headerParams = @{'Authorization' = "Bearer $($GraphApiAccessToken)" }

# Create transcript folder
$logFolder = "$($script_root)\transcript"
$logFile = "$($logFolder)\log_$(Get-Date -format dd-MMM-yyyy_H_mm_ss).txt"
if (!(Test-Path $logFolder)) {
    Write-Output "$(Get-Date) : Creating Transcript folder $($logFolder)"
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}

#region config

# Import configuration
try {
    $Config = (Resolve-Path $Config -ErrorAction STOP).Path.ToString()
}
catch {
    Write-Output "$(Get-Date) : [X] Cannot open the configuration file. Make sure that the file is accessible and valid."
    LogEnd
    return $null
}

$options = Get-Content $Config -Raw | ConvertFrom-Json

$transLog = $options.parameters.transLog


Write-Output "$(Get-Date) : Using configuration from $($Config)"

$enabledReport = @()

# Select reports from config
# Parameters
$saveRawData = $options.parameters.saveRawData

# Developer
$graphApiVersion = $options.developer.graphApiVersion

# License
$reportLicenseAssigned = $options.reports.license
if ($reportLicenseAssigned) { $enabledReport += "License" }

# MS365 Active Users
$reportMs365ActiveUsers = $options.reports.ms365ActiveUsers
if ($reportMs365ActiveUsers) { $enabledReport += "Users" }

# MS365 Activation
$reportMs365ActivationUsers = $options.reports.ms365ActivationUsers
if ($reportMs365ActivationUsers) { $enabledReport += "Activations" }

# Exchange
$reportMailboxUsageAndProvisioning = $options.reports.exchangeMailbox
$reportOffice365GroupsProvisioning = $options.reports.Office365Groups
$reportEmailAppUsage = $options.reports.exchangeApp
$reportMailTraffic = $options.reports.exchangeMailTraffic
$reportTopMailTraffic = $options.reports.exchangeTopMailTraffic
$reportATPDetections = $options.reports.exchangeATPDetections

if (    $reportMailboxUsageAndProvisioning -or `
        $reportOffice365GroupsProvisioning -or `
        $reportEmailAppUsage -or `
        $reportTopMailTraffic -or `
        $reportMailTraffic -or `
        $reportATPDetections) {
    $enabledReport += "Exchange"
}

# Sharepoint
$reportSPO = $options.reports.sharepoint
if ($reportSPO) { $enabledReport += "SharePoint" }
# Onedrive
$reportOneDrive = $options.reports.onedrive
if ($reportOneDrive) { $enabledReport += "OneDrive" }
# SkypeForBusiness
$reportSkypeForBusiness = $options.reports.SkypeForBusiness
if ($reportSkypeForBusiness) { $enabledReport += "Skype for Business" }
# Teams
$reportTeams = $options.reports.teams
if ($reportTeams) { $enabledReport += "Microsoft Teams" }



# Check if there's any enabled report. If none, stop transcript and exit script.
if (!$enabledReport) {
    Write-Output "$(Get-Date) : [X] There are no reports enabled in your configuration file. Make sure to enable reports first then try again."
    LogEnd
    return $null
}

$enabledReportList = $enabledReport -join ","

Write-Output "$(Get-Date) : Enabled reports are - $enabledReportList"
#endregion

# Start Transcript Logging
if ($transLog) {
    Write-Output "$(Get-Date) : Transcript - $($logFile)"
    LogStart -logPath $logFile
}
Write-Output "$(Get-Date) : Script Start"

#Region MS Grap Authentication

# tenantName check
$tenantName = $options.auth.tenantName
if (!($tenantName)) {
    Write-Output "$(Get-Date) : [X] The tenantName value is missing from the configuration file $Config"
    LogEnd
    return $null
}

# msGraphAppKey check
$msGraphAppID = $options.auth.msGraphAppID
if (!($msGraphAppID)) {
    Write-Output "$(Get-Date) : [X] The msGraphAppID value is missing from the configuration file $Config"
    LogEnd
    return $null
}

# if msGraphAuthType = 1 (Certificate)
if ($options.auth.msGraphAuthType -eq 1) {
    $msGraphAppCertificateThumbprint = $options.auth.msGraphAppCertificateThumbprint
    if (!($msGraphAppCertificateThumbprint)) {
        Write-Output "$(Get-Date) : [X] The msGraphAppCertificateThumbprint is missing from the configuration file $Config."
        LogEnd
        return $null
    }
    else {
        try {
            Write-Output "$(Get-Date) : Trying to acquire an access token using certificate [$msGraphAppCertificateThumbprint]."
            $oAuth = Get-MsalToken -ClientId $msGraphAppID -TenantId $tenantName -ClientCertificate (Get-Item Cert:\CurrentUser\My\$msGraphAppCertificateThumbprint) -ErrorAction STOP
            $headerParams = @{'Authorization' = "Bearer $($oAuth.AccessToken)" }
            Write-Output "$(Get-Date) : [$([Char]8730)] Graph Graph API access token acquired."
        }
        catch {
            Write-Output "$(Get-Date) : [X] Failed to get access token."
            Write-Output "$(Get-Date) : $($_.Exception.Message)"
            LogEnd
            return $null
        }
    }
}

# if msGraphAuthType = 2 (App Key)
if ($options.auth.msGraphAuthType -eq 2) {
    # if msGraphAppKey is missing
    $msGraphAppKey = $options.auth.msGraphAppKey
    if (!($msGraphAppKey)) {
        Write-Output "$(Get-Date) : [X] The msGraphAppKey is missing from the configuration file $Config."
        LogEnd
        return $null
    }
    else {
        try {
            Write-Output "$(Get-Date) : Trying to acquire an access token using app key."
            $msGraphAppKeySecured = new-object securestring
            $msGraphAppKey.ToCharArray() | ForEach-Object { $msGraphAppKeySecured.AppendChar($_) }
            $oAuth = Get-MsalToken -ClientId $msGraphAppID -TenantId $tenantName -ClientSecret $msGraphAppKeySecured -ErrorAction STOP
            $headerParams = @{'Authorization' = "Bearer $($oAuth.AccessToken)" }
            Write-Output "$(Get-Date) : [$([Char]8730)] Graph Graph API access token acquired."
        }
        catch {
            Write-Output "$(Get-Date) : [X] Failed to get access token."
            Write-Output "$(Get-Date) : $($_.Exception.Message)"
            LogEnd
            return $null
        }
    }
}
#EndRegion
#Region Exchange Authentication
if ($enabledReport -contains 'Exchange') {
    $exchangeAuthType = $options.auth.exchangeAuthType
    if ($exchangeAuthType) {
        if ($exchangeAuthType -eq 1) {
            $exchangeAppID = $options.auth.exchangeAppID
            $exchangeAppCertificateThumbprint = $options.auth.exchangeAppCertificateThumbprint
            if (!($exchangeAppID) -or !($exchangeAppCertificateThumbprint)) {
                Write-Output "$(Get-Date) : [X] The exchangeAppID or exchangeAppCertificateThumbprint values is missing from the configuration file $Config."
                LogEnd
                return $null
            }
            else {
                try {
                    Write-Output "$(Get-Date) : Trying to connect Exchange Online PowerShell using app certificate [$exchangeAppCertificateThumbprint]."
                    Connect-ExchangeOnline -AppId $exchangeAppID -Organization $tenantName -CertificateThumbprint $exchangeAppCertificateThumbprint -ShowBanner:$false -ErrorAction STOP
                    Write-Output "$(Get-Date) : [$([Char]8730)] Connected to Exchange Online PowerShell."
                }
                catch {
                    Write-Output "$(Get-Date) : [X] Failed to connect to Exchange Online PowerShell. [X]"
                    Write-Output "$(Get-Date) : $($_.Exception.Message)"
                    LogEnd
                    return $null
                }
            }
        }
        elseif ($exchangeAuthType -eq 2) {
            $exchangeCredentialFile = $options.auth.exchangeCredentialFile
            if (!($exchangeCredentialFile)) {
                Write-Output "$(Get-Date) : [X] The exchangeCredentialFile value is missing from the configuration file $Config.`nUpdate your configuration to point exchangeCredentialFile to the right location of the credential file."
                LogEnd
                return $null
            }
            else {
                try {
                    $exchangeCredential = Import-CliXml $exchangeCredentialFile -ErrorAction STOP
                    Write-Output "$(Get-Date) : Trying to connect Exchange Online PowerShell using credential."
                    Connect-ExchangeOnline -Organization $tenantName -Credential $exchangeCredential -ShowBanner:$false -ErrorAction STOP
                    Write-Output "$(Get-Date) : [$([Char]8730)] Connected to Exchange Online PowerShell."
                }
                catch {
                    Write-Output "$(Get-Date) : [X] Failed to connect to Exchange Online PowerShell."
                    Write-Output "$(Get-Date) : $($_.Exception.Message)"
                    LogEnd
                    return $null
                }
            }
        }
        else {
            Write-Output "$(Get-Date) : [X] The exchangeAuthType value is not valid.`nValid values as 1, 2.`n * 1 = App + Certificate`n2 = Credential"
            LogEnd
            return $null
        }
    }
    else {
        Write-Output "$(Get-Date) : [X] The exchangeAuthType value is not valid.`nValid values as 1, 2.`n * 1 = App + Certificate`n2 = Credential"
        LogEnd
        return $null
    }
}
#EndRegion

#organization details
$uri = "https://graph.microsoft.com/beta/organization`?`$select=displayname"
$organizationName = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams).Value.displayname
Write-Output "$(Get-Date) : Your organization name is $organizationName"

# Set report period
[int]$dPeriod = $options.parameters.period

[datetime]$today = (Get-Date).Date
$startDate = ($today).AddDays(-$dPeriod)
$endDate = $today
Write-Output "$(Get-Date) : Setting Report Period to $dPeriod Days [$($startDate)] - [$($endDate)]"
$fileSuffix = ('{0:yyyy-MMM-dd_}' -f ($startDate)) + ('{0:yyyy-MMM-dd}' -f ($endDate))

# Create report folder for this period (if it does not exist)
$reportFolder = "$($script_root)\reports\$($organizationName)\$fileSuffix"
if (!(Test-Path $reportFolder)) {
    Write-Output "$(Get-Date) : Creating Reports folder $($reportFolder)"
    New-Item -ItemType Directory -Path $reportFolder | Out-Null
}

# Empty the report folder
Get-ChildItem -Path "$($reportFolder)\*" -Exclude debug.log | Remove-Item -Force

# HTML report header
$mailSubject = "[$($organizationName)] Microsoft 365 Usage Report for the period of " + ("{0:yyyy-MM-dd}" -f $startDate ) + " to " + ("{0:yyyy-MM-dd}" -f $endDate)
$html = "<html>"
$html += "<head><title>$($mailSubject)</title>"
$html += "<meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
$html += '<style type="text/css">' + "`n"
$html += @'
#heading
	{
		font-family:"Segoe UI Light";
		width:fit-content;
        border-collapse:collapse;
		padding:3px 7px 2px 7px;
        font-size:1.8em;
		text-align:left;
		background-color:#fff;
		color:#000;
	}
#section
	{
		font-family:"Segoe UI";
		width:fit-content;
		border-collapse:collapse;
		font-size:1.2em;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
        color:#000;
	}
#data
	{
		font-family:'Segoe UI';
		border-collapse:collapse;
	}
#data td, #data th
	{
		font-size:0.8em;
		border:1px solid #DDD;
		padding:3px 7px 2px 7px;
		width:fit-content;
	}
#data th
	{
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
		color:#000; text-align:left;
	}
#data td
	{
		padding-top:5px;
		padding-bottom:4px;
		text-align:left;

	}
'@
$html += "`n"
$html += '</style>'
$html += "</head><body>"
$html += '<table id="heading">'
$html += '<tr><th>' + $organizationName + '</th></tr>'
$html += '<tr><th>Microsoft 365 Usage Report</th></tr>'
$html += '<tr><th>' + ("{0:yyyy-MM-dd}" -f $startDate ) + ' to ' + ("{0:yyyy-MM-dd}" -f $endDate) + '</th></tr>'
$html += '</table>'

#==============================================
# Licenses Assigned Report
#==============================================
if ($reportLicenseAssigned) {
    Write-Output "$(Get-Date) : Processing Assigned License Report"
    Write-Output "$(Get-Date) :      --> Getting Office 365 user count and assigned licenses"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOffice365ActiveUserDetail(period='D$($dPeriod)')"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    $license = "" | Select-Object TotalUsers, TotalUsersLicensed, TotalUsersUnlicensed, Exchange, Sharepoint, OneDrive, SkypeForBusiness, Yammer, Teams
    $license.TotalUsers = $raw.Count
    $license.TotalUsersLicensed = ($raw | Where-Object { $_."Assigned Products" }).count
    $license.TotalUsersUnlicensed = ($raw | Where-Object { !($_."Assigned Products") }).count
    $license.Exchange = ($raw | Where-Object { $_."Has Exchange License" -eq $true }).count
    $license.Sharepoint = ($raw | Where-Object { $_."Has Sharepoint License" -eq $true }).count
    $license.OneDrive = ($raw | Where-Object { $_."Has OneDrive License" -eq $true }).count
    $license.SkypeForBusiness = ($raw | Where-Object { $_."Has Skype For Business License" -eq $true }).count
    $license.Yammer = ($raw | Where-Object { $_."Has Yammer License" -eq $true }).count
    $license.Teams = ($raw | Where-Object { $_."Has Teams License" -eq $true }).count

    $html += '<hr><table id="section"><tr><th>Users and Assigned Licenses</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Total Users</th><td>' + ("{0:N0}" -f $license.TotalUsers) + '</td></tr>'
    $html += '<tr><th>Licensed Users</th><td>' + ("{0:N0}" -f $license.TotalUsersLicensed) + '</td></tr>'
    $html += '<tr><th>Unlicensed Users</th><td>' + ("{0:N0}" -f $license.TotalUsersUnlicensed) + '</td></tr>'
    $html += '<tr><th>Exchange</th><td>' + ("{0:N0}" -f $license.Exchange) + '</td></tr>'
    $html += '<tr><th>Sharepoint</th><td>' + ("{0:N0}" -f $license.Sharepoint) + '</td></tr>'
    $html += '<tr><th>OneDrive</th><td>' + ("{0:N0}" -f $license.OneDrive) + '</td></tr>'
    $html += '<tr><th>SkypeForBusiness</th><td>' + ("{0:N0}" -f $license.SkypeForBusiness) + '</td></tr>'
    $html += '<tr><th>Yammer</th><td>' + ("{0:N0}" -f $license.Yammer) + '</td></tr>'
    $html += '<tr><th>Teams</th><td>' + ("{0:N0}" -f $license.Teams) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $raw | Export-Csv "$($reportFolder)\raw_Office365ActiveUserDetail.csv" -NoTypeInformation
    }
}

#==============================================
# MS365 Active Users Count Report
#==============================================

if ($reportMs365ActiveUsers) {
    Write-Output "$(Get-Date) : Processing Office 365 Active Users Report"
    Write-Output "$(Get-Date) :      --> Getting Office 365 active user count per service"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOffice365ServicesUserCounts(period='D$($dPeriod)')"
    $result = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    $raw = "" | Select-Object Office365Active, ExchangeActive, OneDriveActive, SharePointActive, SkypeforBusinessActive, YammerActive, TeamsActive, Office365inActive, ExchangeinActive, OneDriveinActive, SharePointinActive, SkypeforBusinessinActive, YammerinActive, TeamsinActive
    $raw.Office365Active = $result."office 365 active"
    $raw.ExchangeActive = $result."exchange Active"
    $raw.OneDriveActive = $result."oneDrive Active"
    $raw.SharePointActive = $result."sharePoint Active"
    $raw.SkypeforBusinessActive = $result."skype For Business Active"
    $raw.YammerActive = $result."yammer Active"
    $raw.TeamsActive = $result."teams Active"

    $raw.Office365inActive = $result."office 365 inactive"
    $raw.ExchangeinActive = $result."exchange inActive"
    $raw.OneDriveinActive = $result."oneDrive inActive"
    $raw.SharePointinActive = $result."sharePoint inActive"
    $raw.SkypeforBusinessinActive = $result."skype For Business inActive"
    $raw.YammerinActive = $result."yammer inActive"
    $raw.TeamsinActive = $result."teams inActive"

    $html += '<hr><table id="section"><tr><th>Office 365 Active Users</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Service</th><th>Active</th><th>Inactive</th></tr>'
    $html += '<tr><th>Office 365</th><td>' + ("{0:N0}" -f [int]$raw.Office365Active) + '</td><td>' + ("{0:N0}" -f [int]$raw.Office365inActive) + '</td></tr>'
    $html += '<tr><th>Exchange</th><td>' + ("{0:N0}" -f [int]$raw.ExchangeActive) + '</td><td>' + ("{0:N0}" -f [int]$raw.ExchangeinActive) + '</td></tr>'
    $html += '<tr><th>OneDrive</th><td>' + ("{0:N0}" -f [int]$raw.OneDriveActive) + '</td><td>' + ("{0:N0}" -f [int]$raw.OneDriveinActive) + '</td></tr>'
    $html += '<tr><th>Sharepoint</th><td>' + ("{0:N0}" -f [int]$raw.SharepointActive) + '</td><td>' + ("{0:N0}" -f [int]$raw.SharepointinActive) + '</td></tr>'
    $html += '<tr><th>SkypeForBusiness</th><td>' + ("{0:N0}" -f [int]$raw.SkypeForBusinessActive) + '</td><td>' + ("{0:N0}" -f $raw.SkypeForBusinessinActive) + '</td></tr>'
    $html += '<tr><th>Yammer</th><td>' + ("{0:N0}" -f [int]$raw.YammerActive) + '</td><td>' + ("{0:N0}" -f [int]$raw.YammerinActive) + '</td></tr>'
    $html += '<tr><th>Teams</th><td>' + ("{0:N0}" -f [int]$raw.TeamsActive) + '</td><td>' + ("{0:N0}" -f [int]$raw.TeamsinActive) + '</td></tr>'

    $html += '</table>'

    if ($saveRawData) {
        $raw | Export-Csv "$($reportFolder)\raw_Office365ServicesUserCounts.csv" -NoTypeInformation
    }
}

#==============================================
# MS365 Activations Users Count Report
#==============================================

if ($reportMs365ActivationUsers) {
    Write-Output "$(Get-Date) : Processing Office 365 Activations Users Count Report"
    Write-Output "$(Get-Date) :      --> Getting Office 365 app activations count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOffice365ActivationsUserCounts"
    $result = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    $html += '<hr><table id="section"><tr><th>Office 365 Activations Users</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Product Type</th><th>Assigned</th><th>Activated</th><th>Shared Computer Activation</th></tr>'

    foreach ($detail in $result) {
        $html += '<tr><th>' + ($detail."product Type") + '</th>
        <td>' + ("{0:N0}" -f [int]$detail.assigned) + '</td>
        <td>' + ("{0:N0}" -f [int]$detail.activated) + '</td>
        <td>' + ("{0:N0}" -f [int]$detail."shared Computer Activation") + '</td>
        </tr>'
    }
    $html += '</table>'

    if ($saveRawData) {
        $result | Export-Csv "$($reportFolder)\raw_Office365ActivationsUserCounts.csv" -NoTypeInformation
    }
}

#==============================================
# Exchange Online Report
#==============================================
if ($reportMailboxUsageAndProvisioning) {
    #get mailbox usage
    Write-Output "$(Get-Date) : Processing Mailbox Usage and Provisioning Report"
    Write-Output "$(Get-Date) :      --> Getting Exchange Online mailbox usage and provisioning details"

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getMailboxUsageDetail(period='D$($dPeriod)')"
    $result = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $mailboxUsageAndProvisioningData = @()
    foreach ($detail in $result) {
        $raw = "" | Select-Object UserPrincipalName, DisplayName, IsDeleted, DeletedDate, CreatedDate, LastActivityDate, StorageUsedByte, IssueWarningQuotaByte, ProhibitSendQuotaByte, ProhibitSendReceiveQuotaByte, IsBelow25Percent, IsOverQuota, IsInActive
        $raw.UserPrincipalName = $detail."User Principal Name"
        $raw.DisplayName = $detail."Display Name"
        $raw.IsDeleted = $detail."Is Deleted"
        if ($detail."Deleted Date") { $raw.DeletedDate = [datetime]$detail."Deleted Date" }
        if ($detail."Created Date") { $raw.CreatedDate = [datetime]$detail."Created Date" }
        if ($detail."Last Activity Date") { $raw.LastActivityDate = [datetime]$detail."Last Activity Date" }
        $raw.StorageUsedByte = [double]$detail."Storage Used (Byte)"
        $raw.IssueWarningQuotaByte = [double]$detail."Issue Warning Quota (Byte)"
        $raw.ProhibitSendQuotaByte = [double]$detail."Prohibit Send Quota (Byte)"
        $raw.ProhibitSendReceiveQuotaByte = [double]$detail."Prohibit Send/Receive Quota (Byte)"

        if (!($raw.LastActivityDate)) {
            $raw.IsInActive = $true
        }
        elseif ($raw.LastActivityDate -lt $startDate) {
            $raw.IsInActive = $true
        }
        else {
            $raw.IsInActive = $false
        }

        if ((($raw.StorageUsedByte / $raw.ProhibitSendReceiveQuotaByte) * 100) -lt 25) {
            $raw.IsBelow25Percent = $true
        }
        else {
            $raw.IsBelow25Percent = $false
        }

        if ($raw.StorageUsedByte -ge $raw.ProhibitSendReceiveQuotaByte) {
            $raw.IsOverQuota = $true
        }
        else {
            $raw.IsOverQuota = $false
        }
        $mailboxUsageAndProvisioningData += $raw
    }

    # Get deleted mailbox
    Write-Output "$(Get-Date) :      --> Getting list of deleted mailboxes"
    # v1.2.1 - changed back to Get-Mailbox
    $deletedMailbox = Get-Mailbox -ResultSize Unlimited -SoftDeletedMailbox -Filter "WhenSoftDeleted -ge '$startDate'" |
    Select-Object UserPrincipalName, WhenSoftDeleted |
    Sort-Object UserPrincipalName

    $exchangeMailboxStatus = "" | Select-Object ActiveMailbox, InActiveMailbox, CreatedMailbox, DeletedMailbox
    $exchangeMailboxStatus.ActiveMailbox = ($mailboxUsageAndProvisioningData | Where-Object { $_.IsInActive -eq $false }).count
    $exchangeMailboxStatus.InactiveMailbox = ($mailboxUsageAndProvisioningData | Where-Object { $_.IsInActive }).count
    $exchangeMailboxStatus.CreatedMailbox = ($mailboxUsageAndProvisioningData | Where-Object { $_.CreatedDate -ge $today.AddDays(-$dPeriod) }).count
    $exchangeMailboxStatus.DeletedMailbox = $deletedMailbox.count

    $html += '<hr><table id="section"><tr><th>Exchange Mailbox Status</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Active Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.ActiveMailbox) + '</td></tr>'
    $html += '<tr><th>Inactive Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.InactiveMailbox) + '</td></tr>'
    $html += '</table>'

    $html += '<hr><table id="section"><tr><th>Exchange Mailbox Provisioning</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Created Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.CreatedMailbox) + '</td></tr>'
    $html += '<tr><th>Deleted Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.deletedMailbox) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $result | Export-Csv "$($reportFolder)\raw_getMailboxUsageDetail.csv" -NoTypeInformation
        $mailboxUsageAndProvisioningData | Export-csv "$($reportFolder)\raw_MailboxUsageDetail.csv" -NoTypeInformation
        $deletedMailbox | Export-Csv "$($reportFolder)\raw_exchangeDeletedMailbox.csv" -NoTypeInformation
    }

    # Get quota status
    Write-Output "$(Get-Date) : Processing Mailbox Quota Report"
    Write-Output "$(Get-Date) :      --> Getting Exchange Online mailbox quota status"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getMailboxUsageQuotaStatusMailboxCounts(period='D$($dPeriod)')"
    $raw = ((Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv)[0]
    $quota = "" | Select-Object UnderLimit, WarningIssued, SendProhibited, SendReceiveProhibited, Below25Percent
    [int]$quota.UnderLimit = $raw."Under Limit"
    [int]$quota.WarningIssued = $raw."Warning Issued"
    [int]$quota.SendProhibited = $raw."Send Prohibited"
    [int]$quota.SendReceiveProhibited = $raw."Send/Receive Prohibited"
    [int]$quota.Below25Percent = ($mailboxUsageAndProvisioningData | Where-Object { $_.IsBelow25Percent }).count


    # EXO Storage Used
    Write-Output "$(Get-Date) :      --> Getting Exchange Online storage usage (tenant)"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getMailboxUsageStorage(period='D$($dPeriod)')"
    $exoStorage = ((Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv)

    $html += '<hr><table id="section"><tr><th>Exchange Mailbox Storage and Quota</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Storage Used (TB)</th><td>' + ("{0:N2}" -f (($exoStorage[0].'Storage Used (Byte)') / 1TB)) + '</td></tr>'
    $html += '<tr><th>Below 25% Used</th><td>' + ("{0:N0}" -f $quota.Below25Percent) + '</td></tr>'
    $html += '<tr><th>Under Limit</th><td>' + ("{0:N0}" -f $quota.underLimit) + '</td></tr>'
    $html += '<tr><th>Warning Issued</th><td>' + ("{0:N0}" -f $quota.WarningIssued) + '</td></tr>'
    $html += '<tr><th>Send Prohibited</th><td>' + ("{0:N0}" -f $quota.SendProhibited) + '</td></tr>'
    $html += '<tr><th>Send/Receive Prohibited</th><td>' + ("{0:N0}" -f $quota.SendReceiveProhibited) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $raw | Export-Csv "$($reportFolder)\raw_MailboxUsageQuotaStatusMailboxCounts.csv" -NoTypeInformation
    }
}

# Email app report
if ($reportEmailAppUsage) {
    Write-Output "$(Get-Date) : Processing Email App Report"
    Write-Output "$(Get-Date) :      --> Getting Exchange Online email app distribution count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getEmailAppUsageAppsUserCounts(period='D$($dPeriod)')"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    $html += '<hr><table id="section"><tr><th>Exchange Mail App Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Mail for Mac</th><td>' + ("{0:N0}" -f [int]$raw."Mail for Mac") + '</td></tr>'
    $html += '<tr><th>Outlook for Mac</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Mac") + '</td></tr>'
    $html += '<tr><th>Outlook for Windows</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Windows") + '</td></tr>'
    $html += '<tr><th>Outlook for Mobile</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Mobile") + '</td></tr>'
    $html += '<tr><th>Other for Mobile</th><td>' + ("{0:N0}" -f [int]$raw."Other for Mobile") + '</td></tr>'
    $html += '<tr><th>Outlook for Web</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Web") + '</td></tr>'
    $html += '<tr><th>POP3 App</th><td>' + ("{0:N0}" -f [int]$raw."POP3 App") + '</td></tr>'
    $html += '<tr><th>IMAP4 App</th><td>' + ("{0:N0}" -f [int]$raw."IMAP4 App") + '</td></tr>'
    $html += '<tr><th>SMTP App</th><td>' + ("{0:N0}" -f [int]$raw."SMTP App") + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $raw | Export-Csv "$($reportFolder)\raw_exchangeMailAppUsage.csv" -NoTypeInformation
    }
}

# Microsoft 365 Groups report
if ($reportOffice365GroupsProvisioning) {
    Write-Output "$(Get-Date) : Processing Office 365 Groups Report"

    # Get all current Microsoft 365 Groups only
    Write-Output "$(Get-Date) :      --> Getting all Office 365 groups"
    $liveGroups = @()
    $uri = "https://graph.microsoft.com/$graphApiVersion/groups`?`$filter=groupTypes/any(c:c+eq+'Unified')`&`$select=mailNickname,deletedDateTime,createdDateTime"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams)
    if ($raw.value) {
        $liveGroups += $raw.value
        while (($raw.'@odata.nextLink')) {
            $raw = (Invoke-RestMethod -Method Get -Uri ($raw.'@odata.nextLink') -Headers $headerParams)
            $liveGroups += $raw.value
        }
    }
    Write-Output "$(Get-Date) :      --> Getting list of deleted Office 365 groups"
    # Get all deleted Microsoft 365 Groups only
    $deletedGroups = @()
    $uri = "https://graph.microsoft.com/$graphApiVersion/directory/deletedItems/microsoft.graph.group`?`$filter=groupTypes/any(c:c+eq+'Unified')`&`$select=mailNickname,deletedDateTime,createdDateTime"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams)
    if ($raw.value) {
        $deletedGroups += $raw.value
        while (($raw.'@odata.nextLink')) {
            $raw = (Invoke-RestMethod -Method Get -Uri ($raw.'@odata.nextLink') -Headers $headerParams)
            $deletedGroups += $raw.value
        }
    }

    $o365Groups = "" | Select-Object LiveGroups, CreatedGroups, DeletedGroups
    $o365Groups.LiveGroups = $liveGroups.count
    $o365Groups.CreatedGroups = ($liveGroups | Where-Object { ([datetime]$_.createdDateTime) -ge $startDate }).Count
    $o365Groups.DeletedGroups = ($deletedGroups | Where-Object { ([datetime]$_.deletedDateTime) -ge $startDate }).Count

    $html += '<hr><table id="section"><tr><th>Microsoft 365 Groups</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Current Groups</th><td>' + ("{0:N0}" -f $o365Groups.LiveGroups) + '</td></tr>'
    $html += '<tr><th>Created Groups</th><td>' + ("{0:N0}" -f $o365Groups.CreatedGroups) + '</td></tr>'
    $html += '<tr><th>Deleted Groups</th><td>' + ("{0:N0}" -f $o365Groups.DeletedGroups) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $liveGroups | Export-Csv "$($reportFolder)\raw_exchangeOffice365LiveGroups.csv" -NoTypeInformation
        $deletedGroups | Export-Csv "$($reportFolder)\raw_exchangeOffice365DeletedGroups.csv" -NoTypeInformation -Append
    }
}

# Mail traffic (inbound/outbound)
if ($reportMailTraffic) {
    Write-Output "$(Get-Date) : Processing Mail Traffic Report"
    Write-Output "$(Get-Date) :      --> Getting mail traffic data"
    $mailTrafficData = Get-MailTrafficReport -StartDate $startDate -EndDate $endDate -AggregateBy Summary

    $mailTraffic = "" | Select-Object Inbound, Outbound, Malware, Spam
    $mailTraffic.Inbound = ($mailTrafficData | Where-Object { $_.Direction -eq "Inbound" -AND ($_.EventType -eq 'goodmail' -or $_.EventType -like "spam*" -or $_.EventType -eq 'malware' -or $_.EventType -eq 'TransportRuleHits') } | Measure-Object MessageCount -sum).Sum
    $mailTraffic.Outbound = ($mailTrafficData | Where-Object { $_.Direction -eq "Outbound" -AND ($_.EventType -eq 'goodmail' -or $_.EventType -like "spam*" -or $_.EventType -eq 'malware' -or $_.EventType -eq 'TransportRuleHits') } | Measure-Object MessageCount -sum).Sum
    $mailTraffic.Spam = ($mailTrafficData | Where-Object { $_.Direction -eq "Inbound" -AND $_.EventType -like "spam*" } | Measure-Object MessageCount -sum).Sum
    $mailTraffic.Malware = ($mailTrafficData | Where-Object { $_.Direction -eq "Inbound" -AND $_.EventType -eq 'malware' } | Measure-Object MessageCount -sum).Sum

    $html += '<hr><table id="section"><tr><th>Mail Traffic</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Outbound</th><td>' + ("{0:N0}" -f $mailTraffic.Outbound) + '</td></tr>'
    $html += '<tr><th>Inbound</th><td>' + ("{0:N0}" -f $mailTraffic.inbound) + '</td></tr>'
    $html += '</table>'

    $html += '<hr><table id="section"><tr><th>Malware and Spam Detection</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Spam</th><td>' + ("{0:N0}" -f $mailTraffic.Spam) + '</td></tr>'
    $html += '<tr><th>Malware</th><td>' + ("{0:N0}" -f $mailTraffic.Malware) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $mailTrafficData | Export-Csv "$($reportFolder)\raw_exchangeMailTraffic.csv" -NoTypeInformation
    }
}

# ATP Mail detections
if ($reportATPDetections) {
    Write-Output "$(Get-Date) : Processing ATP Mail Detection Report"

    $atpSafeLinks_splat = @{
        StartDate   = $startDate
        EndDate     = (Get-Date $endDate).AddDays(-1)
        AppNameList = @('Email Client', 'Outlook')
    }
    Write-Output "$(Get-Date) :      --> Getting ATP SafeLinks blocked message count"
    $atpSafeLinks = Get-SafeLinksAggregateReport @atpSafeLinks_splat | Where-Object { $_.Action -eq 'Blocked' }
    Write-Output "$(Get-Date) :      --> Getting ATP SafeAttachments blocked message count"
    $atpSafeAttachments = Get-MailTrafficATPReport -StartDate $startDate -EndDate $endDate -EventType 'ATP safe attachments'
    $atpMailTraffic = "" | Select-Object atpSafeLinks, atpSafeAttachments
    $atpMailTraffic.atpSafeLinks = $atpSafeLinks.MessageCount
    $atpMailTraffic.atpSafeAttachments = ($atpSafeAttachments | measure-object messagecount -sum).sum

    $html += '<hr><table id="section"><tr><th>ATP Mail Detection</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Blocked by ATP Safe Links</th><td>' + ("{0:N0}" -f $atpMailTraffic.atpSafeLinks) + '</td></tr>'
    $html += '<tr><th>Blocked by ATP Safe Attachments</th><td>' + ("{0:N0}" -f $atpMailTraffic.atpSafeAttachments) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $atpSafeLinks | Export-Csv "$($reportFolder)\raw_exchangeAtpSafeLinks.csv" -NoTypeInformation
        $atpSafeAttachments | Export-Csv "$($reportFolder)\raw_exchangeAtpSafeAttachments.csv" -NoTypeInformation
    }
}

# Top 10 mail traffic reports
if ($reportTopMailTraffic) {
    Write-Output "$(Get-Date) : Processing Top Mail Traffic Report"

    # Top mail report - ALL
    Write-Output "$(Get-Date) :      --> Getting Top Mail Traffic Raw Data"
    $topMailReport = @()
    $pageSize = 5000
    $page = 0
    Do {
        $page++
        $temp = Get-MailTrafficTopReport -StartDate $startDate -EndDate $endDate -pagesize $pageSize -page $page -AggregateBy Summary
        $topMailReport = $topMailReport + $temp
    } While ($temp.count -eq $pageSize)
    $topMailReport | Add-Member -MemberType ScriptProperty -Name TotalMessage -Value { [int]$this.MessageCount }

    # Top 10 Spam Recipients
    Write-Output "$(Get-Date) :      --> Getting Top 10 Spam Recipients"
    $top10SpamRecipient = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopSpamUser' }
    $top10SpamRecipient = $top10SpamRecipient | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Malware Recipients
    Write-Output "$(Get-Date) :      --> Getting Top 10 Malware Recipients"
    $top10MalwareRecipient = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopMalwareUser' }
    $top10MalwareRecipient = $top10MalwareRecipient | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Mail Senders
    Write-Output "$(Get-Date) :      --> Getting Top 10 Mail Senders"
    $top10MailSender = $topMailReport | Where-Object { $_.Direction -eq 'Outbound' -And $_.EventType -eq 'TopMailUser' }
    $top10MailSender = $top10MailSender | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Mail Recipients
    Write-Output "$(Get-Date) :      --> Getting Top 10 Mail Recipients"
    $top10MailRecipient = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopMailUser' }
    $top10MailRecipient = $top10MailRecipient | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Malware
    Write-Output "$(Get-Date) :      --> Getting Top 10 Malware"
    $top10Malware = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopMalware' }
    $top10Malware = $top10Malware | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 mail sender
    $html += '<hr><table id="section"><tr><th>Top 10 Mail Senders</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>User</th><th>Message Count</th></tr>'
    foreach ($mailSender in $top10MailSender) {
        $html += '<tr><td>' + $mailSender.Name + '</td><td>' + ("{0:N0}" -f $mailSender.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 mail recipients
    $html += '<hr><table id="section"><tr><th>Top 10 Mail Recipients</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>User</th><th>Message Count</th></tr>'
    foreach ($mailRecipient in $top10MailRecipient) {
        $html += '<tr><td>' + $mailRecipient.Name + '</td><td>' + ("{0:N0}" -f $mailRecipient.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 Spam Recipients
    $html += '<hr><table id="section"><tr><th>Top 10 Spam Recipients</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>User</th><th>Message Count</th></tr>'
    foreach ($spamRecipient in $top10SpamRecipient) {
        $html += '<tr><td>' + $spamRecipient.Name + '</td><td>' + ("{0:N0}" -f $spamRecipient.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 Malware Recipients
    $html += '<hr><table id="section"><tr><th>Top 10 Malware Recipients</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>User</th><th>Message Count</th></tr>'
    foreach ($malwareRecipient in $top10MalwareRecipient) {
        $html += '<tr><td>' + $malwareRecipient.Name + '</td><td>' + ("{0:N0}" -f $malwareRecipient.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 Malware
    $html += '<hr><table id="section"><tr><th>Top 10 Malware</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>User</th><th>Message Count</th></tr>'
    foreach ($malware in $top10Malware) {
        $html += '<tr><td>' + $malware.Name + '</td><td>' + ("{0:N0}" -f $malware.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    if ($saveRawData) {
        $topMailReport | Export-Csv "$($reportFolder)\raw_exchangeTopMailTraffic.csv" -NoTypeInformation
    }
}
#==============================================
# Sharepoint Report
#==============================================
if ($reportSPO) {

    Write-Output "$(Get-Date) : Processing SharePoint Report"
    Write-Output "$(Get-Date) :      --> Getting SharePoint Sites Usage"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSharePointSiteUsageDetail(period='D$($dPeriod)')"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $raw | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this."Last Activity Date" }
    $spoSites = "" | Select-Object Total, Active, Inactive
    $spoSites.Total = ($raw | Where-Object { $_.'Is Deleted' -eq $false }).Count
    $spoSites.Inactive = ($raw | Where-Object { $_.LastActivityDate -lt ($today.AddDays(-$dPeriod)) -and $_.'Is Deleted' -eq $false }).Count
    $spoSites.Active = ($raw | Where-Object { $_.LastActivityDate -ge ($today.AddDays(-$dPeriod)) -and $_.'Is Deleted' -eq $false }).Count

    Write-Output "$(Get-Date) :      --> Getting SharePoint Storage Usage"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSharePointSiteUsageStorage(period='D$($dPeriod)')"
    $spoStorage = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $spoStorage | Add-Member -MemberType ScriptProperty -Name ReportDate -Value { [datetime]$this."Report Date" }
    $spoStorage | Add-Member -MemberType ScriptProperty -Name Inactive -Value { [int]$this.Total - [int]$this.Active }
    $spoStorage | Add-Member -MemberType ScriptProperty -Name UsedGB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1gb) }
    $spoStorage | Add-Member -MemberType ScriptProperty -Name UsedTB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1tb) }
    $spoStorage = $spoStorage | Sort-Object ReportDate | Select-Object -Last 1

    $html += '<hr><table id="section"><tr><th>SharePoint Site and Storage Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Storage Used (TB)</th><td>' + ("{0:N2}" -f [decimal]$spoStorage.UsedTB) + '</td></tr>'
    $html += '<tr><th>Total Sites</th><td>' + ("{0:N0}" -f [int]$spoSites.Total) + '</td></tr>'
    $html += '<tr><th>Active Sites</th><td>' + ("{0:N0}" -f [int]$sposites.Active) + '</td></tr>'
    $html += '<tr><th>InActive Sites</th><td>' + ("{0:N0}" -f [int]$sposites.inactive) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $raw | Export-Csv "$($reportFolder)\raw_sharePointSiteUsageDetail.csv" -NoTypeInformation
    }
}

#==============================================
# OneDrive Report
#==============================================
if ($reportOneDrive) {
    Write-Output "$(Get-Date) : Processing OneDrive Report"
    Write-Output "$(Get-Date) :      --> Getting OneDrive Usage"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOneDriveUsageAccountDetail(period='D$($dPeriod)')"
    $getOneDriveUsageAccountDetail = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $getOneDriveUsageAccountDetail | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this."Last Activity Date" }
    $oneDriveSites = "" | Select-Object Total, Active, Inactive
    $oneDriveSites.Total = ($getOneDriveUsageAccountDetail).Count
    $oneDriveSites.Inactive = ($getOneDriveUsageAccountDetail | Where-Object { $_.LastActivityDate -lt ($today.AddDays(-$dPeriod)) }).Count
    $oneDriveSites.Active = ($getOneDriveUsageAccountDetail | Where-Object { $_.LastActivityDate -ge ($today.AddDays(-$dPeriod)) }).Count

    Write-Output "$(Get-Date) :      --> Getting OneDrive Storage Usage"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOneDriveUsageStorage(period='D$($dPeriod)')"
    $oneDriveStorage = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name ReportDate -Value { [datetime]$this."Report Date" }
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name Inactive -Value { [int]$this.Total - [int]$this.Active }
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name UsedGB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1gb) }
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name UsedTB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1tb) }
    $oneDriveStorage = $oneDriveStorage | Sort-Object ReportDate | Select-Object -Last 1

    $html += '<hr><table id="section"><tr><th>OneDrive Site Count and Storage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Storage Used (TB)</th><td>' + ("{0:N2}" -f [decimal]$oneDriveStorage.UsedTB) + '</td></tr>'
    $html += '<tr><th>Total Sites</th><td>' + ("{0:N0}" -f [int]$oneDriveSites.Total) + '</td></tr>'
    $html += '<tr><th>Active Sites</th><td>' + ("{0:N0}" -f [int]$oneDriveSites.Active) + '</td></tr>'
    $html += '<tr><th>InActive Sites</th><td>' + ("{0:N0}" -f [int]$oneDriveSites.inactive) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $getOneDriveUsageAccountDetail | Export-Csv "$($reportFolder)\raw_oneDriveUsageAccountDetail.csv" -NoTypeInformation
        $oneDriveSites | Export-Csv "$($reportFolder)\raw_oneDriveSites.csv" -NoTypeInformation
        $oneDriveStorage | Export-Csv "$($reportFolder)\raw_oneDriveStorage.csv" -NoTypeInformation
    }
}
#==============================================
# Sype for Business Report
#==============================================
if ($reportSkypeForBusiness) {
    Write-Output "$(Get-Date) : Processing Skype For Business Report"
    # Total minutes (audio/video)
    # Organized minutes
    Write-Output "$(Get-Date) :      --> Getting SfB organized minutes"
    $uri1 = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessOrganizerActivityMinuteCounts(period='D$($dPeriod)')"
    $sfb1 = (Invoke-RestMethod -Method Get -Uri $uri1 -Headers $headerParams) | ConvertFrom-Csv

    if ($saveRawData) {
        $sfb1 | Export-Csv "$($reportFolder)\raw_SkypeForBusinessOrganizerActivityMinuteCounts.csv" -NoTypeInformation
    }

    # Participant minutes
    Write-Output "$(Get-Date) :      --> Getting SfB participant minutes"
    $uri2 = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessParticipantActivityMinuteCounts(period='D$($dPeriod)')"
    $sfb2 = (Invoke-RestMethod -Method Get -Uri $uri2 -Headers $headerParams) | ConvertFrom-Csv

    if ($saveRawData) {
        $sfb2 | Export-Csv "$($reportFolder)\raw_SkypeForBusinessParticipantActivityMinuteCounts.csv" -NoTypeInformation
    }

    # Peer to peer minutes
    Write-Output "$(Get-Date) :      --> Getting SfB p2p minutes"
    $uri3 = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessPeerToPeerActivityMinuteCounts(period='D$($dPeriod)')"
    $sfb3 = (Invoke-RestMethod -Method Get -Uri $uri3 -Headers $headerParams) | ConvertFrom-Csv

    if ($saveRawData) {
        $sfb3 | Export-Csv "$($reportFolder)\raw_SkypeForBusinessPeerToPeerActivityMinuteCounts.csv" -NoTypeInformation
    }

    # Assemble object
    $SfbMinutes = "" | Select-Object Organized, Participated, PeerToPeer, Total
    $SfbMinutes.Organized = ($sfb1 | Measure-Object "Audio/Video" -sum).sum
    $SfbMinutes.Participated = ($sfb2 | Measure-Object "Audio/Video" -sum).sum
    $SfbMinutes.PeerToPeer = (($sfb3 | Measure-Object "Audio" -sum).sum + ($sfb3 | Measure-Object "Video" -sum).sum)
    $SfbMinutes.Total = $SfbMinutes.Organized + $SfbMinutes.Participated + $SfbMinutes.PeerToPeer

    # Active user, conference and p2p sessions
    # Active User Count
    Write-Output "$(Get-Date) :      --> Getting SfB active user count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessActivityUserDetail(period='D$($dPeriod)')"
    $activeUserCount = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $activeUserCount | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this."Last Activity Date" }
    if ($saveRawData) {
        $activeUserCount | Export-Csv "$($reportFolder)\raw_SkypeForBusinessActivityUserDetail.csv" -NoTypeInformation
    }
    $activeUserCount = ($activeUserCount | Where-Object { $_.LastActivityDate -ge $startDate -and $_.LastActivityDate -le $endDate }).Count

    # Conference count
    Write-Output "$(Get-Date) :      --> Getting SfB conference count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessOrganizerActivityCounts(period='D$($dPeriod)')"
    $conferenceCount = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    if ($saveRawData) {
        $conferenceCount | Export-Csv "$($reportFolder)\raw_SkypeForBusinessOrganizerActivityCounts.csv" -NoTypeInformation
    }

    $conferenceCount = $conferenceCount | Measure-Object -Property IM, "Audio/Video", "App Sharing", Web, "Dial-in/out 3rd Party", "Dial-in/out Microsoft" -sum
    $conferenceCount = ($conferenceCount | Measure-Object Sum -sum).Sum

    # Peer to peer count
    Write-Output "$(Get-Date) :      --> Getting SfB p2p activity count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessPeerToPeerActivityCounts(period='D$($dPeriod)')"
    $peerTOpeerCount = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $peerTOpeerCount | Add-Member -MemberType ScriptProperty -Name ReportDate -Value { [datetime]$this."Report Date" }
    $peerTOpeerCount = $peerTOpeerCount | Where-Object { $_.ReportDate -ge $startDate -and $_.ReportDate -le $endDate }
    if ($saveRawData) {
        $peerTOpeerCount | Export-Csv "$($reportFolder)\raw_SkypeForBusinessPeerToPeerActivityCounts.csv" -NoTypeInformation
    }
    $peerTOpeerCount = $peerTOpeerCount | Measure-Object -Property IM, Audio, Video, "App Sharing", "File Transfer" -sum
    $peerTOpeerCount = ($peerTOpeerCount | Measure-Object Sum -sum).Sum

    $sfbCount = "" | Select-Object ActiveUser, Conference, PeerToPeer
    $sfbCount.ActiveUser = $activeUserCount
    $sfbCount.Conference = $conferenceCount
    $sfbCount.PeerToPeer = $peerTOpeerCount

    $html += '<hr><table id="section"><tr><th>Skype For Business Activities and Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Audio & Video Minutes</th><td>' + ("{0:N0}" -f $sfbMinutes.Total) + '</td></tr>'
    $html += '<tr><th>Active Users</th><td>' + ("{0:N0}" -f $sfbCount.ActiveUser) + '</td></tr>'
    $html += '<tr><th>Conferences</th><td>' + ("{0:N0}" -f $sfbCount.Conference) + '</td></tr>'
    $html += '<tr><th>Peer-To-Peer Sessions</th><td>' + ("{0:N0}" -f $sfbCount.PeerToPeer) + '</td></tr>'
    $html += '</table>'

    # Device usage distribution
    Write-Output "$(Get-Date) :      --> Getting SfB device usage distribution count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessDeviceUsageDistributionUserCounts(period='D$($dPeriod)')"
    $sfbDevices = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    if ($saveRawData) {
        $sfbDevices | Export-Csv "$($reportFolder)\raw_SkypeForBusinessDeviceUsageDistributionUserCounts.csv" -NoTypeInformation
    }
    $html += '<hr><table id="section"><tr><th>Skype For Business Device Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Windows</th><td>' + ("{0:N0}" -f [int]$sfbDevices.Windows) + '</td></tr>'
    $html += '<tr><th>Windows Phone</th><td>' + ("{0:N0}" -f [int]$sfbDevices."Windows Phone") + '</td></tr>'
    $html += '<tr><th>Android Phone</th><td>' + ("{0:N0}" -f [int]$sfbDevices."Android Phone") + '</td></tr>'
    $html += '<tr><th>iPhone</th><td>' + ("{0:N0}" -f [int]$sfbDevices.iPhone) + '</td></tr>'
    $html += '<tr><th>iPad</th><td>' + ("{0:N0}" -f [int]$sfbDevices.iPad) + '</td></tr>'
    $html += '</table>'
}

#==============================================
# Microsoft Teams Report
#==============================================
if ($reportTeams) {
    Write-Output "$(Get-Date) : Processing MS Teams Report"
    Write-Output "$(Get-Date) :      --> Getting Teams active user count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getTeamsUserActivityUserDetail(period='D$($dPeriod)')"
    $TeamsUserActivityUserDetail = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $TeamsUserActivityUserDetail | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this.'Last Activity Date' }
    Write-Output "$(Get-Date) :      --> Getting Teams user activity count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getTeamsUserActivityCounts(period='D$($dPeriod)')"
    $TeamsUserActivityCounts = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    Write-Output "$(Get-Date) :      --> Getting Teams device usage distribution count"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getTeamsDeviceUsageDistributionUserCounts(period='D$($dPeriod)')"
    $TeamsDeviceUsageDistributionUserCounts = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    # Teams Users
    $html += '<hr><table id="section"><tr><th>Teams Users</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Teams Users</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'Is Licensed' -eq 'Yes' }).count) + '</td></tr>'
    $html += '<tr><th>Active Teams Users</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'Is Licensed' -eq 'Yes' -and $_.LastActivityDate -ge $startDate }).count) + '</td></tr>'
    $html += '<tr><th>Inctive Teams Users</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'Is Licensed' -eq 'Yes' -and $_.LastActivityDate -lt $startDate }).count) + '</td></tr>'
    $html += '<tr><th>Guest Users</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'User Principal Name' -match '#EXT#' }).count) + '</td></tr>'
    $html += '</table>'

    # Teams User Activity
    $html += '<hr><table id="section"><tr><th>Teams User Activity</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Total 1:1 Calls</th><td>' + ("{0:N0}" -f ($TeamsUserActivityCounts.Calls | Measure-Object -Sum).Sum) + '</td></tr>'
    $html += '<tr><th>Total Team Chat Messages</th><td>' + ("{0:N0}" -f ($TeamsUserActivityCounts.'Team Chat Messages' | Measure-Object -Sum).Sum) + '</td></tr>'
    $html += '<tr><th>Total Private Chat Messages</th><td>' + ("{0:N0}" -f ($TeamsUserActivityCounts.'Private Chat Messages' | Measure-Object -Sum).Sum) + '</td></tr>'
    $html += '</table>'

    # Teams Device Usage
    $html += '<hr><table id="section"><tr><th>Teams Device Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th>Windows</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Windows)) + '</td></tr>'
    $html += '<tr><th>Mac</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Mac)) + '</td></tr>'
    $html += '<tr><th>Web</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Web)) + '</td></tr>'
    $html += '<tr><th>iOS</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.iOS)) + '</td></tr>'
    $html += '<tr><th>Android Phone</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.'Android Phone')) + '</td></tr>'
    $html += '<tr><th>Windows Phone</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.'Windows Phone')) + '</td></tr>'
    $html += '<tr><th>Chrome OS</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.'Chrome OS')) + '</td></tr>'
    $html += '<tr><th>Linux</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Linux)) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $TeamsUserActivityUserDetail | Export-Csv "$($reportFolder)\raw_TeamsUserActivityUserDetail.csv" -NoTypeInformation
        $TeamsUserActivityCounts | Export-Csv "$($reportFolder)\raw_TeamsUserActivityCounts.csv" -NoTypeInformation
        $TeamsUserActivityCounts | Export-Csv "$($reportFolder)\raw_TeamsDeviceUsageDistributionUserCounts.csv" -NoTypeInformation
    }
}

# Build the HTML report
$html += '<p><table id="section">'
$html += '<tr><th>----END of REPORT----</th></tr></table></p>'
$html += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
$html += '<b>[ENABLED REPORTS]</b><br />'
$html += ($enabledReportList).Split(",") -join "<br />"
$html += '<br /><br />'
$html += '<b>[REPORT DETAILS]</b><br />'
$html += 'Report Period: ' + $dPeriod + ' days<br />'
$html += 'Generated from Computer: ' + ($env:COMPUTERNAME) + '<br />'
$html += 'Script File: ' + $MyInvocation.MyCommand.Definition + '<br />'
$html += 'Config File: ' + $Config + '<br />'
$html += '</p><p>'
$html += '<a href="' + ($scriptInfo.PROJECTURI) + '">Ms365UsageReport v.' + ($scriptInfo.Version) + '</a></p>'
$html += '</body></html>'
$html | Out-File "$($reportFolder)\report.html"

Write-Output "$(Get-Date) : Reports saved to $($reportFolder)"

$sendEmail = $options.mail.sendEmail
if ($sendEmail) {
    Write-Output "$(Get-Date) : Sending email report"
    $fromAddress = $options.mail.fromAddress
    try {
        #message
        $mailBody = @{
            message = @{
                subject                = $mailSubject
                body                   = @{
                    contentType = "HTML"
                    content     = $html
                }

                internetMessageHeaders = @(
                    @{
                        name  = "X-Mailer"
                        value = "Ms365UsageReport by June Castillote"
                    }
                )
            }
        }

        # To address
        if ($options.mail.toAddress) {
            [array]$toAddress = $options.mail.toAddress.Split(",")
            # create JSON-format recipients
            $toAddressJSON = @()
            $toAddress | ForEach-Object {
                $toAddressJSON += @{EmailAddress = @{Address = $_ } }
            }
            $mailBody.message += @{
                toRecipients = @(
                    $ToAddressJSON
                )
            }
        }

        # Cc address
        if ($options.mail.ccAddress) {
            [array]$ccAddress = $options.mail.ccAddress.Split(",")
            # create JSON-format recipients
            $ccAddressJSON = @()
            $ccAddress | ForEach-Object {
                $ccAddressJSON += @{EmailAddress = @{Address = $_ } }
            }
            $mailBody.message += @{
                ccRecipients = @(
                    $ccAddressJSON
                )
            }
        }

        # Bcc address
        if ($options.mail.bccAddress) {
            [array]$bccAddress = $options.mail.bccAddress.Split(",")
            # create JSON-format recipients
            $bccAddressJSON = @()
            $bccAddress | ForEach-Object {
                $bccAddressJSON += @{EmailAddress = @{Address = $_ } }
            }
            $mailBody.message += @{
                bccRecipients = @(
                    $bccAddressJSON
                )
            }
        }

        # attach log file if debug is enabled
        if ($transLog) {
            LogEnd
            [string]$base64_logFile = [convert]::ToBase64String((Get-Content $logFile -Encoding byte))
            $mailBody.message += @{
                attachments = @(
                    @{
                        "@odata.type"  = "#microsoft.graph.fileAttachment"
                        "name"         = "transcript.txt"
                        "contentBytes" = $base64_logFile
                    }
                )
            }
        }
        $mailBody = $mailBody | ConvertTo-JSON -Depth 4
        $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint('https://graph.microsoft.com')
        $mailApiUri = "https://graph.microsoft.com/$graphApiVersion/users/$($fromAddress)/sendmail"
        Invoke-RestMethod -Method Post -Uri $mailApiUri -Body $mailbody -Headers $headerParams -ContentType application/json -ErrorAction STOP
        $null = $ServicePoint.CloseConnectionGroup("")
        Write-Output "$(Get-Date) : [$([Char]8730)] Sending Complete"
    }
    catch {
        $null = $ServicePoint.CloseConnectionGroup("")
        Write-Output "$(Get-Date) : [X] Sending failed"
        Write-Output "$(Get-Date) : $($_.Exception)"
        LogEnd
        [System.GC]::Collect()
        return $null
    }
}

Write-Output "$(Get-Date) : Script End"
LogEnd
[System.GC]::Collect()