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

- A registered Azure AD (OAuth) App with the following settings:
  
  > *Annex:* [Registering a New Azure AD App](#Registering-a-New-Azure-AD-App)
  
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

![Ms365UsageReport Files](images/script_files.png)<br>Script files

### Files List

- *Get-Ms365UsageReport.ps1* - this is the main script file.
- *config_template.json* - this is the configuration file template.
- *LICENSE* - the license document for this repository.
- *README.md* - this document that you are reading right now.
- *.gitattributes* and *.gitignore* - ignore these files, they don't affect the script.

## Configuration

### Make a New Configuration File

The script uses configuration details from a JSON document. You can have many configuration files with different settings if you want.

To create a new configuration, make a copy of the *config_template.json*. I recommend for you to use your Microsoft 365 tenant domain as the name of the configuration file for easier identification.

![Create a copy of the configuration file](images/copy_config.png)<br>Create a copy of the configuration file

> You can use any name for the new configuration file. The content is what's important, not the file name.

### Customize the Settings

Open your JSON file using any text editor. It would be better to use an editor that supports syntax/language support like [*Notepad++*](https://notepad-plus-plus.org/downloads/), [*Atom*](https://atom.io/), or [*Visual Studio Code*](https://code.visualstudio.com/).

The code below shows the default content of the configuration JSON file. The meaning of each setting is explained in the next section.

```JSON
{
    "parameters": {
        "transLog": "1",
        "saveRawData": "1",
        "period": "30"
    },
    "mail": {
        "sendEmail": "1",
        "fromAddress": "sender@domain.com",
        "toAddress": "recipient1@domain.com,recipient2@domain.com",
        "ccAddress": "",
        "bccAddress": ""
    },
    "reports": {
        "license": "1",
        "sharepoint": "1",
        "onedrive": "1",
        "SkypeForBusiness": "1",
        "teams": "1",
        "Office365Groups": "1",
        "exchangeMailbox": "1",
        "exchangeApp": "1",
        "exchangeTopMailTraffic": "1",
        "exchangeMailTraffic": "1",
        "exchangeATPDetections": "1"
    },
    "developer": {
        "graphApiVersion": "beta"
    }
}
```

#### Settings Explained

> Note: For settings that can be turned ON or OFF, used the value of `"1"` for *ON* and `""` (empty) for *OFF*.

| PARAMETERS      |                                                              |
| --------------- | ------------------------------------------------------------ |
| **transLog**    | Turn ON or OFF the transcript logging. When turned on, the transcript will be saved to the *<script_root>\transcript* folder. |
| **saveRawData** | Turn ON or OFF the saving of raw data. Raw data files are saved to the *<script_root>\reports\<organization name>* |
| period          | The period covered by the report in days. Valid values are: `7`,`30`,`90`,`180`. |

| MAIL            |                                                              |
| --------------- | ------------------------------------------------------------ |
| **sendEmail**   | Turn ON or OFF the sending of the HTML report by email. The HTML report file is saved to *<script_root>\reports\<organization name>\report.html*<br /><br />`"sendEmail": "1"` = ON<br />`"sendEmail": ""` = OFF |
| **fromAddress** | This is the email address used for sending the HTML report. This must be a valid mailbox and email address. Using a shared mailbox as the sender is recommend because it does not required an Exchange Online license.<br />Example:<br />`"fromAddress": "sender@domain.com"` |
| **toAddress**   | The recipient(s) email address that will appear in the TO address. Multiple entries are accepted, but must be delimited by a comma ",".<br />Example:<br />`"toAddress": "recipient1@domain.com,recipient2@domain.com"` |
| **ccAddress**   | The recipient(s) email address that will appear in the CC address. Multiple entries are accepted, but must be delimited by a comma ",".<br />Example: <br />"ccAddress": "recipient1@domain.com,recipient2@domain.com"` |
| **bccAddress**  | The recipient(s) email address that will appear in the BCC address. Multiple entries are accepted, but must be delimited by a comma ",".<br />Example:<br />`"bccAddress": "recipient1@domain.com,recipient2@domain.com"` |

| REPORTS                    |                                                              |
| -------------------------- | ------------------------------------------------------------ |
| **license**                | Turn ON or OFF the license assignment count report.<br />ON: `"license": "1"` <br />OFF: `"license": ""` |
| **sharepoint**             | Turn ON or OFF the SharePoint Online reports.<br />ON: `"sharepoint": "1"` <br />OFF: `"sharepoint": ""` |
| **onedrive**               | Turn ON or OFF the OneDrive for Business reports.<br />ON: `"onedrive": "1"` <br />OFF: `"onedrive": ""` |
| **SkypeForBusiness**       | Turn ON or OFF the Skype for Business reports.<br />ON: `"SkypeForBusiness": "1"` <br />OFF: `"SkypeForBusiness": ""` |
| **teams**                  | Turn ON or OFF the Microsoft Teams report.<br />ON: `"teams": "1"` <br />OFF: `"teams": ""` |
| **Office365Groups**        | Turn ON or OFF the Microsoft 365 Groups report.<br />ON: `"Office365Groups": "1"` <br />OFF: `"Office365Groups": ""` |
| **exchangeMailbox**        | Turn ON or OFF the Exchange Online Mailbox reports.<br />ON: `"exchangeMailbox": "1"` <br />OFF: `"exchangeMailbox": ""` |
| **exchangeApp**            | Turn ON or OFF the Exchange Online Email App report.<br />ON: `"exchangeApp": "1"` <br />OFF: `"exchangeApp": ""` |
| **exchangeMailTraffic**    | Turn ON or OFF the Exchange Online Mail Traffic reports.<br />ON: `"exchangeMailTraffic": "1"` <br />OFF: `"exchangeMailTraffic": ""` |
| **exchangeTopMailTraffic** | Turn ON or OFF the Exchange Online Mail Top Traffic reports.<br />ON: `"exchangeTopMailTraffic": "1"` <br />OFF: `"exchangeTopMailTraffic": ""` |
| **exchangeATPDetections**  | Turn ON or OFF the Exchange Online Mail ATP detection reports.<br />ON: `"exchangeATPDetections": "1"` <br />OFF: `"exchangeATPDetections": ""` |

| DEVELOPER       |                                                              |
| --------------- | ------------------------------------------------------------ |
| graphApiVersion | **DO NOT CHANGE!!! FOR DEVELOPMENT USE ONLY**.<br />This defines the Microsoft Graph API version used by the script. |

## How to Use the Script

### Syntax

The `*Get-Ms365UsageReport.ps1*` script accepts two (2) mandatory parameters.

- `-Config` - This parameter accepts the path of the [JSON configuration](#configuration) file.
- `-GraphApiAccessToken` - This parameter access the MS Graph API pre-authenticated token value.

```PowerShell
.\Get-Ms365UsageReport.ps1 -Config <PATH TO JSON FILE> -GraphApiAccessToken <ACCESS TOKEN> -Verbose
```

### Example 1



### Example 2



## ANNEX

### Registering a New Azure AD App

Go to the [Azure Active Directory admin center](https://aad.portal.azure.com/) and navigate to ***Azure Active Directory > App registrations***. Then, click on **New Registration**.

![azapp01](images/azapp01.png)

* In the **Name**, type in ***Microsoft 365 Usage Reports***.

* In the **Supported account types**, select ***Accounts in this organizational directory only***.

* In the **Redirect URI (optional)**, select ***Web*** and type in ***http://localhost***.

* Click on **Register**.

![azapp01](images/azapp02.png)

### Adding API Permissions

Go to **API Permissions** and click on the **Add a Permission** button.

- **API**: *Microsoft Graph*
- **Permission Type**: *Application*
- **Permission(s)**:
  - *Reports.Read.All* - For reading the usage reports.
  - *Directory.Read.All* - For getting the deleted Microsoft 365 Groups and users.
  - *Mail.Send* - For sending the report by email.


### Generating Access Tokens



