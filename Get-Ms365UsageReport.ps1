
#Requires -Version 5.1

<#PSScriptInfo

.VERSION 1.4

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
    [Parameter(Mandatory = $true)]
    [string]$Config,
    [parameter(Mandatory)]
    [string]$GraphApiAccessToken
)
#Region Functions
Function Stop-TxnLogging {
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

Function Start-TxnLogging {
    param (
        [Parameter(Mandatory = $true)]
        [string]$logPath
    )
    Stop-TxnLogging
    Start-Transcript $logPath -Force | Out-Null
}
#EndRegion Functions

Stop-TxnLogging

#Enable TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$WarningPreference = "SilentlyContinue"
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition
$headerParams = @{'Authorization' = "Bearer $($GraphApiAccessToken)" }

# Create transcript folder
$logFolder = "$($script_root)\transcript"
$logFile = "$($logFolder)\log_$(Get-Date -format dd-MMM-yyyy_H_mm_ss).txt"
if (!(Test-Path $logFolder)) {
    Write-Verbose "$(Get-Date) : Creating Transcript folder $($logFolder)"
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}

# Import configuration
try {
    $Config = (Resolve-Path $Config -ErrorAction STOP).Path.ToString()
}
catch {
    Write-Verbose "$(Get-Date) : [SCRIPT TERMINATED] Cannot open the configuration file. Make sure that the file is accessible and valid."
    Stop-TxnLogging
    return $null
}

$options = Get-Content $Config -Raw | ConvertFrom-Json
$transLog = $options.parameters.transLog

# Start Transcript Logging
if ($transLog) {
    Write-Verbose "$(Get-Date) : Transcript - $($logFile)"
    Start-TxnLogging -logPath $logFile
}
Write-Verbose "$(Get-Date) : Script Start"

#organization details
$uri = "https://graph.microsoft.com/beta/organization`?`$select=displayname"
$organizationName = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams).Value.displayname
Write-Verbose "$(Get-Date) : Your organization name is $organizationName"

# Set report period
[int]$dPeriod = $options.parameters.period

[datetime]$today = (Get-Date).Date
$startDate = ($today).AddDays(-$dPeriod)
$endDate = $today
Write-Verbose "$(Get-Date) : Setting Report Period to $dPeriod Days [$($startDate)] - [$($endDate)]"
$fileSuffix = ('{0:yyyy-MMM-dd_}' -f ($startDate)) + ('{0:yyyy-MMM-dd}' -f ($endDate))

# Create report folder for this period (if it does not exist)
$reportFolder = "$($script_root)\reports\$($organizationName)\$fileSuffix"
if (!(Test-Path $reportFolder)) {
    Write-Verbose "$(Get-Date) : Creating Reports folder $($reportFolder)"
    New-Item -ItemType Directory -Path $reportFolder | Out-Null
}

# Empty the report folder
Get-ChildItem -Path "$($reportFolder)\*" -Exclude debug.log | Remove-Item -Force

#region config
Write-Verbose "$(Get-Date) : Using configuration from $($Config)"

$enabledReport = @()

# Select reports from config
# Parameters
$saveRawData = $options.parameters.saveRawData

# Developer
$graphApiVersion = $options.developer.graphApiVersion

# License
if ($reportLicenseAssigned = $options.reports.license) { $enabledReport += "License" }

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
if ($reportSPO = $options.reports.sharepoint) { $enabledReport += "SharePoint" }
# Onedrive
if ($reportOneDrive = $options.reports.onedrive) { $enabledReport += "OneDrive" }
# SkypeForBusiness
if ($reportSkypeForBusiness = $options.reports.SkypeForBusiness) { $enabledReport += "Skype for Business" }
# Teams
if ($reportTeams = $options.reports.teams) { $enabledReport += "Microsoft Teams" }
# Check if there's any enabled report. If none, stop transcript and exit script.
if (!$enabledReport) {
    Write-Verbose "$(Get-Date) : [SCRIPT TERMINATED] There are no reports enabled in your configuration file. Make sure to enable reports first then try again."
    Stop-TxnLogging
    return $null
}

$enabledReportList = $enabledReport -join ","

Write-Verbose "$(Get-Date) : Enabled reports are - $enabledReportList"
#endregion

# HTML report header
$mailSubject = "[$($organizationName)] Microsoft 365 Usage Report for the period of " + ("{0:MMM-dd-yyyy}" -f $startDate ) + " to " + ("{0:MMMM-dd-yyyy}" -f $endDate)
$html = "<html>"
$html += "<head><title>$($mailSubject)</title>"
$html += "<meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
$html += '<style type="text/css">'
$html += @'
#HeadingInfo
	{
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	}
#HeadingInfo td, #HeadingInfo th
	{
		font-size:0.8em;
		padding:3px 7px 2px 7px;
	}
#HeadingInfo th
	{
		font-size:2.0em;
		font-weight:normal;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
		color:#000;
	}
#section
	{
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	}
#section th.data
	{
		font-size:1.2em;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
		color:#000;
	}
#data
	{
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	}
#data td, #data th
	{
		font-size:0.8em;
		border:1px solid #DDD;
		padding:3px 7px 2px 7px;
	}
#data th
	{
		font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
		color:#000; text-align:left;
	}
#data td
	{ 	font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
		text-align:left;
	}
'@
$html += '</style>'
$html += "</head><body>"
$html += '<table id="HeadingInfo">'
$html += '<tr><th>Microsoft 365 Usage Report</th></tr>'
$html += '<tr><th>' + ("{0:MMM-dd-yyyy}" -f $startDate ) + ' to ' + ("{0:MMMM-dd-yyyy}" -f $endDate) + '</th></tr>'
$html += '</table>'

#==============================================
# Licenses Assigned Report
#==============================================
if ($reportLicenseAssigned) {
    Write-Verbose "$(Get-Date) : Processing Assigned License Report"
    #$enabledReport += "License Assignment"

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOffice365ActiveUserDetail(period='D" + $($dPeriod) + "')"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    $license = "" | Select-Object Exchange, Sharepoint, OneDrive, SkypeForBusiness, Yammer, Teams
    $license.Exchange = ($raw | Where-Object { $_."Has Exchange License" -eq $true }).count
    $license.Sharepoint = ($raw | Where-Object { $_."Has Sharepoint License" -eq $true }).count
    $license.OneDrive = ($raw | Where-Object { $_."Has OneDrive License" -eq $true }).count
    $license.SkypeForBusiness = ($raw | Where-Object { $_."Has Skype For Business License" -eq $true }).count
    $license.Yammer = ($raw | Where-Object { $_."Has Yammer License" -eq $true }).count
    $license.Teams = ($raw | Where-Object { $_."Has Teams License" -eq $true }).count

    $html += '<hr><table id="section"><tr><th class="data">Assigned Licenses</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Exchange</th><td>' + ("{0:N0}" -f $license.Exchange) + '</td></tr>'
    $html += '<tr><th width="15%">Sharepoint</th><td>' + ("{0:N0}" -f $license.Sharepoint) + '</td></tr>'
    $html += '<tr><th width="15%">OneDrive</th><td>' + ("{0:N0}" -f $license.OneDrive) + '</td></tr>'
    $html += '<tr><th width="15%">SkypeForBusiness</th><td>' + ("{0:N0}" -f $license.SkypeForBusiness) + '</td></tr>'
    $html += '<tr><th width="15%">Yammer</th><td>' + ("{0:N0}" -f $license.Yammer) + '</td></tr>'
    $html += '<tr><th width="15%">Teams</th><td>' + ("{0:N0}" -f $license.Teams) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        #$license | Export-Csv "$($reportFolder)\summary_assignedLicenses.csv" -NoTypeInformation
        $raw | Export-Csv "$($reportFolder)\raw_Office365ActiveUserDetail.csv" -NoTypeInformation
    }
}

#==============================================
# Exchange Online Report
#==============================================
if ($reportMailboxUsageAndProvisioning) {
    #get mailbox usage
    Write-Verbose "$(Get-Date) : Processing Mailbox Usage and Provisioning Report"
    #$enabledReport += "Exchange Online"

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getMailboxUsageDetail(period='D" + $($dPeriod) + "')"
    $result = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams -ContentType application/json) | ConvertFrom-Csv
    $mailboxUsageAndProvisioningData = @()
    foreach ($detail in $result) {
        $raw = "" | Select-Object UserPrincipalName, DisplayName, IsDeleted, DeletedDate, CreatedDate, LastActivityDate, ItemCount, StorageUsedByte, IssueWarningQuotaByte, ProhibitSendQuotaByte, ProhibitSendReceiveQuotaByte, IsBelow25Percent, IsOverQuota, IsInActive
        $raw.UserPrincipalName = $detail."User Principal Name"
        $raw.DisplayName = $detail."Display Name"
        $raw.IsDeleted = $detail."Is Deleted"
        if ($detail."Deleted Date") { $raw.DeletedDate = [datetime]$detail."Deleted Date" }
        if ($detail."Created Date") { $raw.CreatedDate = [datetime]$detail."Created Date" }
        if ($detail."Last Activity Date") { $raw.LastActivityDate = [datetime]$detail."Last Activity Date" }
        $raw.ItemCount = [double]$detail."Item Count"
        $raw.StorageUsedByte = [double]$detail."Storage Used (Byte)"

        # sometimes the 'Issue Warning Quota (Byte)' property is empty. If so, we need to get it from the mailbox using Get-Mailbox
        if (!($detail.'Issue Warning Quota (Byte)')) {
            $mailbox = Get-Mailbox ($detail."User Principal Name") | Select-Object IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota
            $raw.IssueWarningQuotaByte = [math]::Round(($mailbox.IssueWarningQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "")), 2)
            $raw.ProhibitSendQuotaByte = [math]::Round(($mailbox.ProhibitSendQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "")), 2)
            $raw.ProhibitSendReceiveQuotaByte = [math]::Round(($mailbox.ProhibitSendReceiveQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "")), 2)
        }
        else {
            $raw.IssueWarningQuotaByte = [double]$detail."Issue Warning Quota (Byte)"
            $raw.ProhibitSendQuotaByte = [double]$detail."Prohibit Send Quota (Byte)"
            $raw.ProhibitSendReceiveQuotaByte = [double]$detail."Prohibit Send/Receive Quota (Byte)"
        }

        if (!($raw.LastActivityDate)) {
            $raw.IsInActive = $true
        }
        elseif ($raw.LastActivityDate -lt $startDate) {
            $raw.IsInActive = $true
        }
        # elseif ((New-TimeSpan -Start $raw.LastActivityDate -End $today).Days -ge 30) {
        #     $raw.IsInActive = $true
        # }
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
    Write-Verbose "$(Get-Date) : Getting list of deleted mailboxes"
    $deletedMailbox = Get-Mailbox -ResultSize Unlimited -SoftDeletedMailbox -Filter "WhenSoftDeleted -ge '$startDate'" |
    Select-Object UserPrincipalName, WhenSoftDeleted |
    Sort-Object UserPrincipalName

    $exchangeMailboxStatus = "" | Select-Object ActiveMailbox, InActiveMailbox, CreatedMailbox, DeletedMailbox
    $exchangeMailboxStatus.ActiveMailbox = ($mailboxUsageAndProvisioningData | Where-Object { $_.IsInActive -eq $false }).count
    $exchangeMailboxStatus.InactiveMailbox = ($mailboxUsageAndProvisioningData | Where-Object { $_.IsInActive }).count
    $exchangeMailboxStatus.CreatedMailbox = ($mailboxUsageAndProvisioningData | Where-Object { $_.CreatedDate -ge $today.AddDays(-$dPeriod) }).count
    $exchangeMailboxStatus.DeletedMailbox = $deletedMailbox.count

    $html += '<hr><table id="section"><tr><th class="data">Exchange Mailbox Status</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Active Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.ActiveMailbox) + '</td></tr>'
    $html += '<tr><th width="15%">Inactive Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.InactiveMailbox) + '</td></tr>'
    $html += '</table>'

    $html += '<hr><table id="section"><tr><th class="data">Exchange Mailbox Provisioning</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Created Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.CreatedMailbox) + '</td></tr>'
    $html += '<tr><th width="15%">Deleted Mailbox</th><td>' + ("{0:N0}" -f $exchangeMailboxStatus.deletedMailbox) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        #$exchangeMailboxStatus | Export-Csv "$($reportFolder)\summary_exchangeUsageAndProvisioning.csv" -NoTypeInformation
        $result | Export-Csv "$($reportFolder)\raw_getMailboxUsageDetail.csv" -NoTypeInformation
        $mailboxUsageAndProvisioningData | Export-csv "$($reportFolder)\raw_MailboxUsageDetail.csv" -NoTypeInformation
        $deletedMailbox | Export-Csv "$($reportFolder)\raw_exchangeDeletedMailbox.csv" -NoTypeInformation
    }

    # Get quota status
    Write-Verbose "$(Get-Date) : Processing Mailbox Quota Report"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getMailboxUsageQuotaStatusMailboxCounts(period='D" + $($dPeriod) + "')"
    $raw = ((Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv)[0]
    $quota = "" | Select-Object UnderLimit, WarningIssued, SendProhibited, SendReceiveProhibited, Below25Percent
    [int]$quota.UnderLimit = $raw."Under Limit"
    [int]$quota.WarningIssued = $raw."Warning Issued"
    [int]$quota.SendProhibited = $raw."Send Prohibited"
    [int]$quota.SendReceiveProhibited = $raw."Send/Receive Prohibited"
    [int]$quota.Below25Percent = ($mailboxUsageAndProvisioningData | Where-Object { $_.IsBelow25Percent }).count


    # EXO Storage Used
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getMailboxUsageStorage(period='D" + $($dPeriod) + "')"
    $exoStorage = ((Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv)

    $html += '<hr><table id="section"><tr><th class="data">Exchange Mailbox Storage and Quota</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Storage Used (TB)</th><td>' + ("{0:N2}" -f (($exoStorage[0].'Storage Used (Byte)') / 1TB)) + '</td></tr>'
    $html += '<tr><th width="15%">Below 25% Used</th><td>' + ("{0:N0}" -f $quota.Below25Percent) + '</td></tr>'
    $html += '<tr><th width="15%">Under Limit</th><td>' + ("{0:N0}" -f $quota.underLimit) + '</td></tr>'
    $html += '<tr><th width="15%">Warning Issued</th><td>' + ("{0:N0}" -f $quota.WarningIssued) + '</td></tr>'
    $html += '<tr><th width="15%">Send Prohibited</th><td>' + ("{0:N0}" -f $quota.SendProhibited) + '</td></tr>'
    $html += '<tr><th width="15%">Send/Receive Prohibited</th><td>' + ("{0:N0}" -f $quota.SendReceiveProhibited) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        #$quota | Export-Csv "$($reportFolder)\summary_MailboxUsageQuota.csv" -NoTypeInformation
        $raw | Export-Csv "$($reportFolder)\raw_MailboxUsageQuotaStatusMailboxCounts.csv" -NoTypeInformation
        #$exoStorage | Export-Csv "$($reportFolder)\summary_MailboxUsageStorage.csv" -NoTypeInformation
    }
}

# Email app report
if ($reportEmailAppUsage) {
    Write-Verbose "$(Get-Date) : Processing Email App Report"

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getEmailAppUsageAppsUserCounts(period='D" + $($dPeriod) + "')"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    $html += '<hr><table id="section"><tr><th class="data">Exchange Mail App Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Mail for Mac</th><td>' + ("{0:N0}" -f [int]$raw."Mail for Mac") + '</td></tr>'
    $html += '<tr><th width="15%">Outlook for Mac</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Mac") + '</td></tr>'
    $html += '<tr><th width="15%">Outlook for Windows</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Windows") + '</td></tr>'
    $html += '<tr><th width="15%">Outlook for Mobile</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Mobile") + '</td></tr>'
    $html += '<tr><th width="15%">Other for Mobile</th><td>' + ("{0:N0}" -f [int]$raw."Other for Mobile") + '</td></tr>'
    $html += '<tr><th width="15%">Outlook for Web</th><td>' + ("{0:N0}" -f [int]$raw."Outlook for Web") + '</td></tr>'
    $html += '<tr><th width="15%">POP3 App</th><td>' + ("{0:N0}" -f [int]$raw."POP3 App") + '</td></tr>'
    $html += '<tr><th width="15%">IMAP4 App</th><td>' + ("{0:N0}" -f [int]$raw."IMAP4 App") + '</td></tr>'
    $html += '<tr><th width="15%">SMTP App</th><td>' + ("{0:N0}" -f [int]$raw."SMTP App") + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $raw | Export-Csv "$($reportFolder)\raw_exchangeMailAppUsage.csv" -NoTypeInformation
    }
}

# Microsoft 365 Groups report
if ($reportOffice365GroupsProvisioning) {
    Write-Verbose "$(Get-Date) : Processing Office 365 Groups Report"

    # Get all current Microsoft 365 Groups only
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

    $html += '<hr><table id="section"><tr><th class="data">Microsoft 365 Groups</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Current Groups</th><td>' + ("{0:N0}" -f $o365Groups.LiveGroups) + '</td></tr>'
    $html += '<tr><th width="15%">Created Groups</th><td>' + ("{0:N0}" -f $o365Groups.CreatedGroups) + '</td></tr>'
    $html += '<tr><th width="15%">Deleted Groups</th><td>' + ("{0:N0}" -f $o365Groups.DeletedGroups) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        #$o365Groups | Export-Csv "$($reportFolder)\summary_exchangeOffice365Groups.csv" -NoTypeInformation
        $liveGroups | Export-Csv "$($reportFolder)\raw_exchangeOffice365LiveGroups.csv" -NoTypeInformation
        $deletedGroups | Export-Csv "$($reportFolder)\raw_exchangeOffice365DeletedGroups.csv" -NoTypeInformation -Append
    }
}

# Mail traffic (inbound/outbound)
if ($reportMailTraffic) {
    Write-Verbose "$(Get-Date) : Processing Mail Traffic Report"
    $mailTrafficData = Get-MailTrafficReport -StartDate $startDate -EndDate $endDate -AggregateBy Summary

    $mailTraffic = "" | Select-Object Inbound, Outbound, Malware, Spam
    $mailTraffic.Inbound = ($mailTrafficData | Where-Object { $_.Direction -eq "Inbound" -AND ($_.EventType -eq 'goodmail' -or $_.EventType -like "spam*" -or $_.EventType -eq 'malware' -or $_.EventType -eq 'TransportRuleHits') } | Measure-Object MessageCount -sum).Sum
    $mailTraffic.Outbound = ($mailTrafficData | Where-Object { $_.Direction -eq "Outbound" -AND ($_.EventType -eq 'goodmail' -or $_.EventType -like "spam*" -or $_.EventType -eq 'malware' -or $_.EventType -eq 'TransportRuleHits') } | Measure-Object MessageCount -sum).Sum
    $mailTraffic.Spam = ($mailTrafficData | Where-Object { $_.Direction -eq "Inbound" -AND $_.EventType -like "spam*" } | Measure-Object MessageCount -sum).Sum
    $mailTraffic.Malware = ($mailTrafficData | Where-Object { $_.Direction -eq "Inbound" -AND $_.EventType -eq 'malware' } | Measure-Object MessageCount -sum).Sum

    $html += '<hr><table id="section"><tr><th class="data">Mail Traffic</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Outbound</th><td>' + ("{0:N0}" -f $mailTraffic.Outbound) + '</td></tr>'
    $html += '<tr><th width="15%">Inbound</th><td>' + ("{0:N0}" -f $mailTraffic.inbound) + '</td></tr>'
    $html += '</table>'

    $html += '<hr><table id="section"><tr><th class="data">Malware and Spam Detection</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Spam</th><td>' + ("{0:N0}" -f $mailTraffic.Spam) + '</td></tr>'
    $html += '<tr><th width="15%">Malware</th><td>' + ("{0:N0}" -f $mailTraffic.Malware) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        #$mailTraffic | Export-Csv "$($reportFolder)\summary_exchangeMailTraffic.csv" -NoTypeInformation
        $mailTrafficData | Export-Csv "$($reportFolder)\raw_exchangeMailTraffic.csv" -NoTypeInformation
    }
}

# ATP Mail detections
if ($reportATPDetections) {
    Write-Verbose "$(Get-Date) : Processing ATP Mail Detection Report"

    $atpSafeLinks_splat = @{
        StartDate   = $startDate
        EndDate     = (Get-Date $endDate).AddDays(-1)
        AppNameList = @('Email Client', 'Outlook')
    }
    $atpSafeLinks = Get-SafeLinksAggregateReport @atpSafeLinks_splat | Where-Object { $_.Action -eq 'Blocked' }
    $atpSafeAttachments = Get-MailTrafficATPReport -StartDate $startDate -EndDate $endDate -EventType 'ATP safe attachments'
    $atpMailTraffic = "" | Select-Object atpSafeLinks, atpSafeAttachments
    $atpMailTraffic.atpSafeLinks = $atpSafeLinks.MessageCount
    $atpMailTraffic.atpSafeAttachments = ($atpSafeAttachments | measure-object messagecount -sum).sum

    $html += '<hr><table id="section"><tr><th class="data">ATP Mail Detection</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Blocked by ATP Safe Links</th><td>' + ("{0:N0}" -f $atpMailTraffic.atpSafeLinks) + '</td></tr>'
    $html += '<tr><th width="15%">Blocked by ATP Safe Attachments</th><td>' + ("{0:N0}" -f $atpMailTraffic.atpSafeAttachments) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        $atpSafeLinks | Export-Csv "$($reportFolder)\raw_exchangeAtpSafeLinks.csv" -NoTypeInformation
        $atpSafeAttachments | Export-Csv "$($reportFolder)\raw_exchangeAtpSafeAttachments.csv" -NoTypeInformation
        #$atpMailTraffic | Export-Csv "$($reportFolder)\summary_atpMailTraffic.csv" -NoTypeInformation
    }
}

# Top 10 mail traffic reports
if ($reportTopMailTraffic) {
    Write-Verbose "$(Get-Date) : Processing Top Mail Traffic Report"

    # Top mail report - ALL
    Write-Verbose "$(Get-Date) : Getting Top Mail Traffic Raw Data"
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
    Write-Verbose "$(Get-Date) : Getting Top 10 Spam Recipients"
    $top10SpamRecipient = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopSpamUser' }
    $top10SpamRecipient = $top10SpamRecipient | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Malware Recipients
    Write-Verbose "$(Get-Date) : Getting Top 10 Malware Recipients"
    $top10MalwareRecipient = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopMalwareUser' }
    $top10MalwareRecipient = $top10MalwareRecipient | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Mail Senders
    Write-Verbose "$(Get-Date) : Getting Top 10 Mail Senders"
    $top10MailSender = $topMailReport | Where-Object { $_.Direction -eq 'Outbound' -And $_.EventType -eq 'TopMailUser' }
    $top10MailSender = $top10MailSender | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Mail Recipients
    Write-Verbose "$(Get-Date) : Getting Top 10 Mail Recipients"
    $top10MailRecipient = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopMailUser' }
    $top10MailRecipient = $top10MailRecipient | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 Malware
    Write-Verbose "$(Get-Date) : Getting Top 10 Malware"
    $top10Malware = $topMailReport | Where-Object { $_.Direction -eq 'Inbound' -And $_.EventType -eq 'TopMalware' }
    $top10Malware = $top10Malware | Sort-Object MessageCount -Descending | Select-Object -First 10

    # Top 10 mail sender
    $html += '<hr><table id="section"><tr><th class="data">Top 10 Mail Senders</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">User</th><th>Message Count</th></tr>'
    foreach ($mailSender in $top10MailSender) {
        $html += '<tr><td width="15%">' + $mailSender.Name + '</td><td>' + ("{0:N0}" -f $mailSender.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 mail recipients
    $html += '<hr><table id="section"><tr><th class="data">Top 10 Mail Recipients</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">User</th><th>Message Count</th></tr>'
    foreach ($mailRecipient in $top10MailRecipient) {
        $html += '<tr><td width="15%">' + $mailRecipient.Name + '</td><td>' + ("{0:N0}" -f $mailRecipient.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 Spam Recipients
    $html += '<hr><table id="section"><tr><th class="data">Top 10 Spam Recipients</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">User</th><th>Message Count</th></tr>'
    foreach ($spamRecipient in $top10SpamRecipient) {
        $html += '<tr><td width="15%">' + $spamRecipient.Name + '</td><td>' + ("{0:N0}" -f $spamRecipient.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 Malware Recipients
    $html += '<hr><table id="section"><tr><th class="data">Top 10 Malware Recipients</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">User</th><th>Message Count</th></tr>'
    foreach ($malwareRecipient in $top10MalwareRecipient) {
        $html += '<tr><td width="15%">' + $malwareRecipient.Name + '</td><td>' + ("{0:N0}" -f $malwareRecipient.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    # Top 10 Malware
    $html += '<hr><table id="section"><tr><th class="data">Top 10 Malware</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">User</th><th>Message Count</th></tr>'
    foreach ($malware in $top10Malware) {
        $html += '<tr><td width="15%">' + $malware.Name + '</td><td>' + ("{0:N0}" -f $malware.MessageCount) + '</td></tr>'
    }
    $html += '</table>'

    if ($saveRawData) {
        $topMailReport | Export-Csv "$($reportFolder)\raw_exchangeTopMailTraffic.csv" -NoTypeInformation
        # $top10SpamRecipient | Export-Csv "$($reportFolder)\summary_top10SpamRecipient.csv" -NoTypeInformation
        # $top10MalwareRecipient | Export-Csv "$($reportFolder)\summary_top10MalwareRecipient.csv" -NoTypeInformation
        # $top10MailSender | Export-Csv "$($reportFolder)\summary_top10MailSender.csv" -NoTypeInformation
        # $top10MailRecipient | Export-Csv "$($reportFolder)\summary_top10MailRecipient.csv" -NoTypeInformation
        # $top10Malware | Export-Csv "$($reportFolder)\summary_top10Malware.csv" -NoTypeInformation
    }
}
#==============================================
# Sharepoint Report
#==============================================
if ($reportSPO) {

    Write-Verbose "$(Get-Date) : Processing SharePoint Report"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSharePointSiteUsageDetail(period='D" + $($dPeriod) + "')"
    $raw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $raw | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this."Last Activity Date" }
    $spoSites = "" | Select-Object Total, Active, Inactive
    $spoSites.Total = ($raw | Where-Object { $_.'Is Deleted' -eq $false }).Count
    $spoSites.Inactive = ($raw | Where-Object { $_.LastActivityDate -lt ($today.AddDays(-$dPeriod)) -and $_.'Is Deleted' -eq $false }).Count
    $spoSites.Active = ($raw | Where-Object { $_.LastActivityDate -ge ($today.AddDays(-$dPeriod)) -and $_.'Is Deleted' -eq $false }).Count

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSharePointSiteUsageStorage(period='D" + $($dPeriod) + "')"
    $spoStorage = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $spoStorage | Add-Member -MemberType ScriptProperty -Name ReportDate -Value { [datetime]$this."Report Date" }
    $spoStorage | Add-Member -MemberType ScriptProperty -Name Inactive -Value { [int]$this.Total - [int]$this.Active }
    $spoStorage | Add-Member -MemberType ScriptProperty -Name UsedGB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1gb) }
    $spoStorage | Add-Member -MemberType ScriptProperty -Name UsedTB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1tb) }
    $spoStorage = $spoStorage | Sort-Object ReportDate | Select-Object -Last 1

    $html += '<hr><table id="section"><tr><th class="data">SharePoint Site Count and Storage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Storage Used (TB)</th><td>' + ("{0:N2}" -f [decimal]$spoStorage.UsedTB) + '</td></tr>'
    $html += '<tr><th width="15%">Total Sites</th><td>' + ("{0:N0}" -f [int]$spoSites.Total) + '</td></tr>'
    $html += '<tr><th width="15%">Active Sites</th><td>' + ("{0:N0}" -f [int]$sposites.Active) + '</td></tr>'
    $html += '<tr><th width="15%">InActive Sites</th><td>' + ("{0:N0}" -f [int]$sposites.inactive) + '</td></tr>'
    $html += '</table>'

    if ($saveRawData) {
        # $spoSites | Export-Csv "$($reportFolder)\summary_sharepointOnlineSites.csv" -NoTypeInformation
        # $spoStorage | Export-Csv "$($reportFolder)\summary_sharepointOnlineStorage.csv" -NoTypeInformation
        $raw | Export-Csv "$($reportFolder)\raw_sharePointSiteUsageDetail.csv" -NoTypeInformation
    }
}

#==============================================
# OneDrive Report
#==============================================
if ($reportOneDrive) {
    Write-Verbose "$(Get-Date) : Processing OneDrive Report"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOneDriveUsageAccountDetail(period='D" + $($dPeriod) + "')"
    $getOneDriveUsageAccountDetail = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $getOneDriveUsageAccountDetail | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this."Last Activity Date" }
    $oneDriveSites = "" | Select-Object Total, Active, Inactive
    $oneDriveSites.Total = ($getOneDriveUsageAccountDetail).Count
    $oneDriveSites.Inactive = ($getOneDriveUsageAccountDetail | Where-Object { $_.LastActivityDate -lt ($today.AddDays(-$dPeriod)) }).Count
    $oneDriveSites.Active = ($getOneDriveUsageAccountDetail | Where-Object { $_.LastActivityDate -ge ($today.AddDays(-$dPeriod)) }).Count

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getOneDriveUsageStorage(period='D" + $($dPeriod) + "')"
    $oneDriveStorage = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name ReportDate -Value { [datetime]$this."Report Date" }
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name Inactive -Value { [int]$this.Total - [int]$this.Active }
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name UsedGB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1gb) }
    $oneDriveStorage | Add-Member -MemberType ScriptProperty -Name UsedTB -Value { "{0:N2}" -f ($this."Storage Used (Byte)" / 1tb) }
    $oneDriveStorage = $oneDriveStorage | Sort-Object ReportDate | Select-Object -Last 1

    $html += '<hr><table id="section"><tr><th class="data">OneDrive Site Count and Storage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Storage Used (TB)</th><td>' + ("{0:N2}" -f [decimal]$oneDriveStorage.UsedTB) + '</td></tr>'
    $html += '<tr><th width="15%">Total Sites</th><td>' + ("{0:N0}" -f [int]$oneDriveSites.Total) + '</td></tr>'
    $html += '<tr><th width="15%">Active Sites</th><td>' + ("{0:N0}" -f [int]$oneDriveSites.Active) + '</td></tr>'
    $html += '<tr><th width="15%">InActive Sites</th><td>' + ("{0:N0}" -f [int]$oneDriveSites.inactive) + '</td></tr>'
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
    Write-Verbose "$(Get-Date) : Processing Skype For Business Report"
    # Total minutes (audio/video)
    # Organized minutes
    $uri1 = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessOrganizerActivityMinuteCounts(period='D" + $($dPeriod) + "')"
    $sfb1 = (Invoke-RestMethod -Method Get -Uri $uri1 -Headers $headerParams) | ConvertFrom-Csv

    if ($saveRawData) {
        $sfb1 | Export-Csv "$($reportFolder)\raw_SkypeForBusinessOrganizerActivityMinuteCounts.csv" -NoTypeInformation
    }

    # Participant minutes
    $uri2 = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessParticipantActivityMinuteCounts(period='D" + $($dPeriod) + "')"
    $sfb2 = (Invoke-RestMethod -Method Get -Uri $uri2 -Headers $headerParams) | ConvertFrom-Csv

    if ($saveRawData) {
        $sfb2 | Export-Csv "$($reportFolder)\raw_SkypeForBusinessParticipantActivityMinuteCounts.csv" -NoTypeInformation
    }

    # Peer to peer minutes
    $uri3 = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessPeerToPeerActivityMinuteCounts(period='D" + $($dPeriod) + "')"
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
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessActivityUserDetail(period='D" + $($dPeriod) + "')"
    $activeUserCount = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $activeUserCount | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this."Last Activity Date" }
    if ($saveRawData) {
        $activeUserCount | Export-Csv "$($reportFolder)\raw_SkypeForBusinessActivityUserDetail.csv" -NoTypeInformation
    }
    $activeUserCount = ($activeUserCount | Where-Object { $_.LastActivityDate -ge $startDate -and $_.LastActivityDate -le $endDate }).Count

    # Conference count
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessOrganizerActivityCounts(period='D" + $($dPeriod) + "')"
    $conferenceCount = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    if ($saveRawData) {
        $conferenceCount | Export-Csv "$($reportFolder)\raw_SkypeForBusinessOrganizerActivityCounts.csv" -NoTypeInformation
    }

    $conferenceCount = $conferenceCount | Measure-Object -Property IM, "Audio/Video", "App Sharing", Web, "Dial-in/out 3rd Party", "Dial-in/out Microsoft" -sum
    $conferenceCount = ($conferenceCount | Measure-Object Sum -sum).Sumd

    # Peer to peer count
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessPeerToPeerActivityCounts(period='D" + $($dPeriod) + "')"
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

    $html += '<hr><table id="section"><tr><th class="data">Skype For Business Activities and Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Audio & Video Minutes</th><td>' + ("{0:N0}" -f $sfbMinutes.Total) + '</td></tr>'
    $html += '<tr><th width="15%">Active Users</th><td>' + ("{0:N0}" -f $sfbCount.ActiveUser) + '</td></tr>'
    $html += '<tr><th width="15%">Conferences</th><td>' + ("{0:N0}" -f $sfbCount.Conference) + '</td></tr>'
    $html += '<tr><th width="15%">Peer-To-Peer Sessions</th><td>' + ("{0:N0}" -f $sfbCount.PeerToPeer) + '</td></tr>'
    $html += '</table>'

    # Device usage distribution
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getSkypeForBusinessDeviceUsageDistributionUserCounts(period='D" + $($dPeriod) + "')"
    $sfbDevices = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    if ($saveRawData) {
        $sfbDevices | Export-Csv "$($reportFolder)\raw_SkypeForBusinessDeviceUsageDistributionUserCounts.csv" -NoTypeInformation
    }
    $html += '<hr><table id="section"><tr><th class="data">Skype For Business Device Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Windows</th><td>' + ("{0:N0}" -f [int]$sfbDevices.Windows) + '</td></tr>'
    $html += '<tr><th width="15%">Windows Phone</th><td>' + ("{0:N0}" -f [int]$sfbDevices."Windows Phone") + '</td></tr>'
    $html += '<tr><th width="15%">Android Phone</th><td>' + ("{0:N0}" -f [int]$sfbDevices."Android Phone") + '</td></tr>'
    $html += '<tr><th width="15%">iPhone</th><td>' + ("{0:N0}" -f [int]$sfbDevices.iPhone) + '</td></tr>'
    $html += '<tr><th width="15%">iPad</th><td>' + ("{0:N0}" -f [int]$sfbDevices.iPad) + '</td></tr>'
    $html += '</table>'
}

#==============================================
# Microsoft Teams Report
#==============================================
if ($reportTeams) {
    Write-Verbose "$(Get-Date) : Processing MS Teams Report"
    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getTeamsUserActivityUserDetail(period='D" + $($dPeriod) + "')"
    $TeamsUserActivityUserDetail = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv
    $TeamsUserActivityUserDetail | Add-Member -MemberType ScriptProperty -Name LastActivityDate -Value { [datetime]$this.'Last Activity Date' }

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getTeamsUserActivityCounts(period='D" + $($dPeriod) + "')"
    $TeamsUserActivityCounts = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    $uri = "https://graph.microsoft.com/$graphApiVersion/reports/getTeamsDeviceUsageDistributionUserCounts(period='D" + $($dPeriod) + "')"
    $TeamsDeviceUsageDistributionUserCounts = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headerParams) | ConvertFrom-Csv

    # Teams Users
    $html += '<hr><table id="section"><tr><th class="data">Teams Users</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Total Teams Users (Licensed)</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'Is Licensed' -eq 'Yes' }).count) + '</td></tr>'
    $html += '<tr><th width="15%">Active Teams Users (Licensed)</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'Is Licensed' -eq 'Yes' -and $_.LastActivityDate -ge $startDate }).count) + '</td></tr>'
    $html += '<tr><th width="15%">Inctive Teams Users (Licensed)</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'Is Licensed' -eq 'Yes' -and $_.LastActivityDate -lt $startDate }).count) + '</td></tr>'
    $html += '<tr><th width="15%">Total Guest Users</th><td>' + ("{0:N0}" -f ($TeamsUserActivityUserDetail | Where-Object { $_.'User Principal Name' -match '#EXT#' }).count) + '</td></tr>'
    $html += '</table>'

    # Teams User Activity
    $html += '<hr><table id="section"><tr><th class="data">Teams User Activity</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Total 1:1 Calls</th><td>' + ("{0:N0}" -f ($TeamsUserActivityCounts.Calls | Measure-Object -Sum).Sum) + '</td></tr>'
    $html += '<tr><th width="15%">Total Team Chat Messages</th><td>' + ("{0:N0}" -f ($TeamsUserActivityCounts.'Team Chat Messages' | Measure-Object -Sum).Sum) + '</td></tr>'
    $html += '<tr><th width="15%">Total Private Chat Messages</th><td>' + ("{0:N0}" -f ($TeamsUserActivityCounts.'Private Chat Messages' | Measure-Object -Sum).Sum) + '</td></tr>'
    $html += '</table>'

    # Teams Device Usage
    $html += '<hr><table id="section"><tr><th class="data">Teams Device Usage</th></tr></table><hr>'
    $html += '<table id="data">'
    $html += '<tr><th width="15%">Windows</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Windows)) + '</td></tr>'
    $html += '<tr><th width="15%">Mac</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Mac)) + '</td></tr>'
    $html += '<tr><th width="15%">Web</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Web)) + '</td></tr>'
    $html += '<tr><th width="15%">iOS</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.iOS)) + '</td></tr>'
    $html += '<tr><th width="15%">Android Phone</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.'Android Phone')) + '</td></tr>'
    $html += '<tr><th width="15%">Windows Phone</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.'Windows Phone')) + '</td></tr>'
    $html += '<tr><th width="15%">Chrome OS</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.'Chrome OS')) + '</td></tr>'
    $html += '<tr><th width="15%">Linux</th><td>' + ("{0:N0}" -f [int]($TeamsDeviceUsageDistributionUserCounts.Linux)) + '</td></tr>'
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
$html += 'Generated from Server: ' + ($env:COMPUTERNAME) + '<br />'
$html += 'Script File: ' + $MyInvocation.MyCommand.Definition + '<br />'
$html += 'Config File: ' + $Config + '<br />'
$html += '</p><p>'
$html += '<a href="' + ($scriptInfo.PROJECTURI) + '">Office 365 Usage Summary Report v.' + ($scriptInfo.Version) + '</a></p>'
$html += '</body></html>'
$html | Out-File "$($reportFolder)\report.html"

Write-Verbose "$(Get-Date) : Reports saved to $($reportFolder)"

$sendEmail = $options.mail.sendEmail
if ($sendEmail) {
    Write-Verbose "$(Get-Date) : Sending email report"
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
                        value = "Get-Office365UsageReport by june.castillote@gmail.com"
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
            Stop-TxnLogging
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
        Invoke-RestMethod -Method Post -Uri $mailApiUri -Body $mailbody -Headers $headerParams -ContentType application/json
        $null = $ServicePoint.CloseConnectionGroup("")
        Write-Verbose "$(Get-Date) : Sending Complete"
    }
    catch {
        Write-Verbose "$(Get-Date) : Sending failed"
        $_.Exception | Format-List
        Stop-TxnLogging
        [System.GC]::Collect()
        return $null
    }
}

Write-Verbose "$(Get-Date) : Script End"
Stop-TxnLogging
[System.GC]::Collect()