# Ms365UsageReport

## Overview

This PowerShell script exports the Office 365 usage report using the [*Microsoft Graph API v1.0*](https://docs.microsoft.com/en-us/graph/overview?view=graph-rest-1.0) and [*ExchangeOnlineManagement PowerShell Module*](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/2.0.3).

The reports that can be exported using this script are:

- Assigned Licenses Report Count
  - Exchange
  - SharePoint
  - OneDrive
  - Skype for Business
  - Teams
  - Yammer
- Exchange Reports
  - Mailbox Status (Active/Inactive)
  - Mailbox Provisioning (Created/Deleted)
  - Mailbox Quota
  - Email App Usage
  - Office 365 Groups Provisioning
  - Mail Traffic (Inbound/Outbound)
  - Spam and Malware Detection (Count)
  - ATP Mail Detection (ATP Safe Links, ATP Safe Attachments)
  - Top 10 Mail Senders
  - Top 10 Mail Recipients
  - Top 10 Spam Recipients
  - Top 10 Malware Recipients
  - Top 10 Malware Detected
- SharePoint Reports
  - Total, Active, Inactive Sites
  - Storage Used (Tenant)
- OneDrive Reports
  - Total, Active, Inactive Sites
  - Storage Used (Tenant)
- Skype for Business Reports
  - Organizer Minutes
  - Participant Minutes
  - Peer-to-Peer Minutes
  - Total Audio & Video Minutes
  - Number of Users
  - Total number of conferences
  - Total number of peer-to-peer sessions
  - Device Usage Distribution
- Microsoft Teams
  - Total Users
  - Active Users
  - Inactive Users
  - Guest Users
  - Total 1:1 Calls
  - Total Team Chat Messages
  - Total Private Chat Messages
  - Device Usage Distribution

## Requirements

- Registered Azure AD (OAuth) App
  - **API**: *Microsoft Graph*
  - **Permission Type**: *Application*
  - **Permission(s)**:
    - *Reports.Read.All* - For reading the usage reports.
    - *Directory.Read.All* - For getting the deleted Microsoft 365 Groups and users.
    - *Mail.Send* - For sending the report by email.
- Windows PowerShell 5.1 or PowerShell 7.
- *[ExchangeOnlineManagement PowerShell Module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/2.0.3)* must be installed on the computer where you will be running this script.
- A valid mailbox that will be used for sending the report. A shared mailbox (no license) is recommended.

## How to Get the Script
[Clone](https://github.com/junecastillote/Ms365UsageReport.git) or [download](https://github.com/junecastillote/Ms365UsageReport/archive/main.zip) the script from the [repository](https://github.com/junecastillote/Ms365UsageReport) and extract the files to your preferred location.

![Ms365UsageReport Files](/images/script_files.png)<br>Script files

### Files List

- *Get-Ms365UsageReport.ps1* - this is the main script file.
- *config_template.json* - this is the configuration file template.
- *LICENSE* - the license document for this repository.
- *README.md* - this document that you are reading right now.
- *.gitattributes* and *.gitignore* - ignore these files, they don't affect the script.

## Configuration

The script uses configuration details from a JSON document. You can have many configuration files with different settings if you want.

To create a new configuration, make a copy of the *config_template.json*.

> You can use any name for the new configuration file. The content is what's important, not the file name.

![Create a copy of the configuration file](/images/copy_config.png)<br>Create a copy of the configuration file
