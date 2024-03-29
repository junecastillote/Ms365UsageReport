# Microsoft 365 Usage Reports PowerShell Script <!-- omit in toc -->

[![GitHub issues](https://img.shields.io/github/issues/junecastillote/Ms365UsageReport)](https://github.com/junecastillote/Ms365UsageReport/issues)
[![GitHub forks](https://img.shields.io/github/forks/junecastillote/Ms365UsageReport)](https://github.com/junecastillote/Ms365UsageReport/network)
[![GitHub license](https://img.shields.io/github/license/junecastillote/Ms365UsageReport)](https://github.com/junecastillote/Ms365UsageReport/blob/main/LICENSE)

> Important!!! - If you're coming from versions below 1.2.5 of this script, you'll need to update your config file from JSON to YAML format. The YAML template is included in this version.

- [Overview](#overview)
- [Requirements](#requirements)
- [How to Get the Script](#how-to-get-the-script)
- [Configuration](#configuration)
  - [Make a New Configuration File](#make-a-new-configuration-file)
  - [YAML Settings Explained](#yaml-settings-explained)
  - [Change the Header Logo](#change-the-header-logo)
- [How to Use the Script](#how-to-use-the-script)
  - [Syntax](#syntax)
  - [Running the Script](#running-the-script)
  - [Script Output](#script-output)
    - [Transcript File](#transcript-file)
    - [Raw Data and HTML Report Files](#raw-data-and-html-report-files)
    - [HTML Report](#html-report)
    - [Email Report](#email-report)
- [ANNEX](#annex)
  - [Register a New Azure AD App](#register-a-new-azure-ad-app)
    - [Adding the Required API Permissions](#adding-the-required-api-permissions)
    - [OPTION 1: Adding a Client Secret](#option-1-adding-a-client-secret)
    - [OPTION 2: Creating and Uploading a Self-Signed Certificate](#option-2-creating-and-uploading-a-self-signed-certificate)
    - [Granting Admin Consent](#granting-admin-consent)
  - [Creating an Encrypted Exchange Online Credentials File](#creating-an-encrypted-exchange-online-credentials-file)

## Overview

This PowerShell script exports the Microsoft 365 usage reports using the [*Microsoft Graph API v1.0*](https://docs.microsoft.com/en-us/graph/overview?view=graph-rest-1.0) and [*ExchangeOnlineManagement PowerShell Module*](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/2.0.3). The results are saved locally and can also be sent by email.

The reports that can be exported using this script are:

- Assigned Licenses Report Count
  - Exchange
  - SharePoint
  - OneDrive
  - Skype for Business
  - Teams
  - Yammer
- Microsoft 365
  - Users
  - Activations
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

  > Annex: [*Register a New Azure AD App*](#Register-a-New-Azure-AD-App)

  - **API**: *Microsoft Graph*
  - **Permission Type**: *Application*
  - **Permission(s)**:
    - `Reports.Read.All` - For reading the usage reports.
    - `Directory.Read.All` - For getting the deleted Microsoft 365 Groups and users.
    - `Mail.Send` - For sending the report by email.

- Windows PowerShell 5.1.

- The [*ExchangeOnlineManagement PowerShell Module*](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/)- must be installed on the computer where you will be running this script. The *minimum* version required is `2.0.3`.

- The [*MSAL.PS PowerShell Module*](https://www.powershellgallery.com/packages/MSAL.PS/) must be installed on the computer where you will be running this script. The *minimum* version required is `4.16.0.4`.

- The [*powershell-yaml module*](https://www.powershellgallery.com/packages/powershell-yaml) must be installed. The script will use this module to process the configuration file. The *minimum* version required is `0.4.2`.

- A valid mailbox that will be used for sending the report. A shared mailbox (no license) is recommended.

## How to Get the Script

[*Clone*](https://github.com/junecastillote/Ms365UsageReport.git) or [*download*](https://github.com/junecastillote/Ms365UsageReport/archive/main.zip) the script from the [*repository*](https://github.com/junecastillote/Ms365UsageReport). After downloading, extract the files to your preferred location.

## Configuration

### Make a New Configuration File

The script uses configuration details from a YAML document. You can have many configuration files with different settings if you want.

To create a new configuration, make a copy of the *config_template.yml*. I recommend for you to use your Microsoft 365 tenant domain as the name of the configuration file for easier identification.

![Create a copy of the configuration file](images/copy_config.png)

> You can use any name for the new configuration file. What's important is the contents, not the file name.

### YAML Settings Explained

Open your YAML file using any text editor. It would be best to use an editor that has syntax/language support like [*Notepad++*](https://notepad-plus-plus.org/downloads/), [*Atom*](https://atom.io/), or [*Visual Studio Code*](https://code.visualstudio.com/).

The code below shows the default content of the configuration YAML file. The meaning of each setting is explained in the next section.

```YAML
auth:
  tenantName: <tenant>.onmicrosoft.com
  msGraphAuthType: 2 # 1 = Certificate, 2 = Secret
  msGraphAppID:
  msGraphAppKey:
  msGraphAppCertificateThumbprint:
  exchangeAuthType: 2 # 1 = Certificate, 2 = Secret
  exchangeAppID:
  exchangeAppCertificateThumbprint:
  exchangeCredentialFile:
parameters:
  transLog: false
  saveRawData: false
  period: 30
  showLogo: true
mail:
  sendEmail: false
  fromAddress: sender@domain.com
  toAddress: recipient1@domain.com,recipient2@domain.com
  ccAddress:
  bccAddress:
reports:
  license: true
  ms365ActiveUsers: true
  ms365ActivationUsers: true
  exchangeMailbox: true
  exchangeApp: true
  exchangeTopMailTraffic: true
  exchangeMailTraffic: true
  exchangeATPDetections: true
  Office365Groups: true
  sharepoint: true
  onedrive: true
  SkypeForBusiness: true
  teams: true
developer:
  graphApiVersion: beta
  scriptCompatibleVersion: 1.2.5
```

<hr>

| AUTH                               |                                                              |
| ---------------------------------- | ------------------------------------------------------------ |
| `tenantName`                       | This is your Microsoft 365 tenant's organization name.<br>Example: `contoso.onmicrosoft.com` |
| `msGraphAuthType`                  | Determines the Graph API credential type.<br><br>`msGraphAuthType: 1` = Use Certificate<br>`msGraphAuthType: 2` = Use Client Secret |
| `msGraphAppID`                     | This is the registered app's Application ID. (refer to *[Register a New Azure AD App](#register-a-new-azure-ad-app)*). |
| `msGraphAppKey`                    | This is the registered app's Secret Key. (refer to *[Adding a Client Secret](#option-1-adding-a-client-secret)*). |
| `msGraphAppCertificateThumbprint`  | This is the registered app's Certificate Thumbprint. (refer to *[Creating and Uploading a Self-Signed Certificate*](#option-2-creating-and-uploading-a-self-signed-certificate)*). |
| `exchangeAuthType`                 | Determines the Exchange Credential authentication type.<br><br>`exchangeAuthType: 1` = Use Certificate<br>`exchangeAuthType: 2` = Use encrypted username and password |
| `exchangeAppID`                    | This is the registered app's Application ID.<br>Refer to *[Setting Up App-Only Authentication using PowerShell](https://adamtheautomator.com/exchange-online-powershell-mfa/#Setting_Up_AppOnly_Authentication_using_PowerShell)*.<br><br><br>This is only required if you're using If you're using `exchangeAuthType: 1`<br>If you're using `exchangeAuthType: 2`, you do not need to add a value to this. |
| `exchangeAppCertificateThumbprint` | This is the registered app's Certificate Thumbprint.<br>Refer to *[Setting Up App-Only Authentication using PowerShell](https://adamtheautomator.com/exchange-online-powershell-mfa/#Setting_Up_AppOnly_Authentication_using_PowerShell)*.<br><br><br>This is only required if you're using If you're using `exchangeAuthType: 1`<br><br>Make sure that the certificate is in the personal certificate store and uploaded to the registered Exchange app in Azure AD.<br><br>If you're using `exchangeAuthType: 2`, you do not need to add a value to this. |
| `exchangeCredentialFile`           | The file path to the encrypted credential file.<br><br>Refer to: *[Creating an Encrypted Exchange Online Credentials File](#creating-an-encrypted-exchange-online-credentials-file)*<br><br>Example:<br>`exchangeCredentialFile: C:\temp\cred.xml` |

<hr>

| PARAMETERS    |                                                              |
| ------------- | ------------------------------------------------------------ |
| `transLog`    | Turn ON or OFF the transcript logging. When turned on, the transcript will be saved to the `<script_root>\transcript` folder.<br />`true` or `false` |
| `saveRawData` | Turn ON or OFF the saving of raw data. Raw data files are saved to the `<script_root>\reports\<organization name>`.<br />`true` or `false` |
| `period`      | The period covered by the report in days. Valid values are: `7`,`30`,`90`,`180`. |
| `showLogo`    | Show or hide the banner logo on top of the report.<br />`true` or `false` |

<hr>

| MAIL          |                                                              |
| ------------- | ------------------------------------------------------------ |
| `sendEmail`   | Turn ON or OFF the sending of the HTML report by email. The HTML report file is saved to *<script_root>\reports\<organization name>\report.html*<br><br>`true` or `false` |
| `fromAddress` | This is the email address used for sending the HTML report. This must be a valid mailbox and email address. Using a shared mailbox as the sender is recommend because it does not required an Exchange Online license.<br>Example:<br>`fromAddress: sender@domain.com` |
| `toAddress`   | The recipient(s) email address that will appear in the TO address. Multiple entries are accepted, but must be delimited by a comma ",".<br>Example:<br>`toAddress: recipient1@domain.com,recipient2@domain.com` |
| `ccAddress`   | The recipient(s) email address that will appear in the CC address. Multiple entries are accepted, but must be delimited by a comma ",".<br>Example: <br>ccAddress: recipient1@domain.com,recipient2@domain.com` |
| `bccAddress`  | The recipient(s) email address that will appear in the BCC address. Multiple entries are accepted, but must be delimited by a comma ",".<br>Example:<br>`bccAddress: recipient1@domain.com,recipient2@domain.com` |

<hr>

| REPORTS                  |                                                              |
| ------------------------ | ------------------------------------------------------------ |
| `license`                | Turn ON or OFF the license assignment count report.<br>`true` or `false` |
| `sharepoint`             | Turn ON or OFF the SharePoint Online reports.<br>`true` or `false` |
| `onedrive`               | Turn ON or OFF the OneDrive for Business reports.<br>`true` or `false` |
| `SkypeForBusiness`       | Turn ON or OFF the Skype for Business reports.<br>`true` or `false` |
| `teams`                  | Turn ON or OFF the Microsoft Teams report.<br>`true` or `false` |
| `Office365Groups`        | Turn ON or OFF the Microsoft 365 Groups report.<br>`true` or `false` |
| `exchangeMailbox`        | Turn ON or OFF the Exchange Online Mailbox reports.<br>`true` or `false` |
| `exchangeApp`            | Turn ON or OFF the Exchange Online Email App report.<br>`true` or `false` |
| `exchangeMailTraffic`    | Turn ON or OFF the Exchange Online Mail Traffic reports.<br>`true` or `false` |
| `exchangeTopMailTraffic` | Turn ON or OFF the Exchange Online Mail Top Traffic reports.<br>`true` or `false` |
| `exchangeATPDetections`  | Turn ON or OFF the Exchange Online Mail ATP detection reports.<br>`true` or `false` |
| `ms365ActiveUsers`       | Turn ON or OFF the Office 365 Active Users report.<br>`true` or `false` |
| `ms365ActivationUsers`   | Turn ON or OFF the Office 365 Activations by Users report.<br>`true` or `false` |

<hr>

| DEVELOPER                 |                                                                                                                    |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `graphApiVersion`         | **DO NOT CHANGE!!! FOR DEVELOPMENT USE ONLY**.<br>This defines the Microsoft Graph API version used by the script. |
| `scriptCompatibleVersion` | **DO NOT CHANGE!!! FOR DEVELOPMENT USE ONLY**.                                                                     |

### Change the Header Logo

To change the report banner logo, replace the `<resource>\logo.png` file.

![Logo](resource/logo.png)

## How to Use the Script

### Syntax

The `Get-Ms365UsageReport.ps1` script accepts one (1) mandatory parameter.

- `-Config` - This parameter accepts the path of the [YAML configuration](#configuration) file.

```PowerShell
.\Get-Ms365UsageReport.ps1 -Config <PATH TO YAML FILE>
```

### Running the Script

1. Open PowerShell and change the working directory to where you saved the script.
2. Run the script. In the example below, the configuration file used is `poshlab.ml.yml` - which is in the same folder as the script.

   ```powershell
   .\Get-Ms365UsageReport.ps1 -config .\poshlab.ml.yml
   ```

You should see a screen output similar to the one below.

![Run Script](images/run_script.png)<br>Running the script

### Script Output

#### Transcript File

If transcript logging is enabled, the transcript is saved to the `transcript` folder.

![Transcript](images/transcript_file.png)

#### Raw Data and HTML Report Files

If raw data saving is enabled, the raw data files and HTML report is saved to the `reports\[ORGANIZATION]\[REPORT PERIOD]` folder.

![raw data and report files](images/report_files.png)

#### HTML Report

![HTML Report](images/html_report.png)

#### Email Report

![Email Report](images/email_report.png)

## ANNEX

### Register a New Azure AD App

Go to the [Azure Active Directory admin center](https://aad.portal.azure.com/) and navigate to ***Azure Active Directory > App registrations***. Then, click on **New Registration**.

![azapp01](images/azApp01.png)

- In the **Name**, type in ***Microsoft 365 Usage Reports***.

- In the **Supported account types**, select ***Accounts in this organizational directory only***.

- In the **Redirect URI (optional)**, select ***Web*** and type in `http://localhost`.

- Click on **Register**.

![azapp02](images/azApp02.png)

After the App has been registered, copy the **Application (client) ID** and **Directory (tenant) ID**.

![azApp09](images/azApp09.png)

#### Adding the Required API Permissions

Go to **API Permissions** and click on the **Add a Permission** button.

![azapp03](images/azApp03.png)

In the **Request API Permission**, select **Microsoft Graph API**.

![azapp04](images/azApp04.png)

In **What type of permissions does your application require?** click on **Application**.

![azapp05](images/azApp05.png)

From the list of permissions, search for and enable the following permissions.

- `Reports.Read.All`

- `Directory.Read.All`

- `Mail.Send`

Once you're done selecting the permissions, click on the **Add permissions** button.

Next, you have two options as to how your application can get authorization. Using a ***Client Secret*** or a ***Self-Signed Certificate***.

#### OPTION 1: Adding a Client Secret

> *Note: You can use either a Certificate or a Client Secret for API authentication. This example shows you only how to create a client secret.*

Go to **Certificates & secrets** and click on the **New client secret** button.

![azapp10](images/azApp10.png)

In the **Add a client secret** page:

- Type in the **Description*** box the description you want to use. In this example, the description used is ***secret key 1***.
- Select the validity period for this secret. In this example, the secret key **Expires** <u>***In 2 years***</u>. Choose which ever one is permitted for your organization.
- Click on **Add**.

Make sure to copy and save the key.

![azapp11](images/azApp11.png)

#### OPTION 2: Creating and Uploading a Self-Signed Certificate

> *Note: You can use either a Certificate or a Client Secret for API authentication. This example shows you only how to create a certificate.*

Copy the code below and run it in PowerShell. The script will create a new self-signed certificate in the personal certificate store with a validity of five (5) years. Then, exports the DER-encoded `ms365UsageReport.cer` file to the current working directory.

```PowerShell
# Create a self-sign certificate in the personal certificate store with 5-year validity.
$certSplat = @{
  Subject = 'CN=ms365UsageReport'
  NotBefore = ((Get-Date).AddDays(-1))
  NotAfter = ((Get-Date).AddYears(5))
  CertStoreLocation = "Cert:\CurrentUser\My"
}
$selfSignedCertificate = New-SelfSignedCertificate @certSplat

# Export the certificate (.CER)
Export-Certificate -Cert $selfSignedCertificate -FilePath .\ms365UsageReport.cer
```

After running the code, you can see the new certificate is in the personal certificate store.

![image-20210118230431617](images/image022.png)<br>
The new self-signed certificate in the personal certificate store

![image-20210118230618692](images/image023.png)<br>
The new self-signed certificate exported to a file

Now that you have the certificate file (.cer), the next step is to upload the certificate to the AzureAd App. Follow these next steps.

1. Go to **Certificates & secrets** and click on the **Upload certificate** button.
2. Click the browse button.
3. Locate the certificate file and click **Open**.
4. Click **Add**.

![click upload certificate](images/image024.png)<br>Uploading a certificate

You should now see the certificate was uploaded. Copy the **Thumbprint** value.

![new certificate uploaded](images/image025.png)<br>The certificate is uploaded

#### Granting Admin Consent

> *Note: Only a Global Admin can grant consent on behalf of the Microsoft 365 Tenant. If you do not have the proper rights, ask your Global Admin to grant the consent.*

You should see that the new API permissions are added, but the status if ***Not granted for [tenant]***. To finish granting the permissions, click on the **Grant admin consent for [tenant]** button.

![azapp06](images/azApp06.png)

When asked to confirm, click on **Yes**.

![azapp07](images/azApp07.png)

The API permissions status should change to **Granted for [tenant]**

![azapp08](images/azApp08.png)

Now you should have the following details available:

- Client ID
- Secret Key
- Tenant ID

### Creating an Encrypted Exchange Online Credentials File

If you're using a username + password to authenticate to Exchange Online, then you'll need to save your encrypted credentials to a file.

> Note: The Exchange Online credential you'll be using must at least have View-Only Organization Management permission

1. In PowerShell, enter this command - `Get-Credential | Export-CliXml <PATH\TO\FILE.xml>`.
2. When prompted, enter the username and password of the Exchange Online credential to use.

![Saving exchange credentials](images/image026.png)<br>Creating an Encrypted Exchange Online Credentials File
