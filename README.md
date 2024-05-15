<h1 align="center">RemoveExchangeEmailAttachments 📨</h1>
<p>
  <img alt="Version" src="https://img.shields.io/badge/version-1-blue.svg?cacheSeconds=2592000" />
</p>

This PowerShell code is designed to remove email attachments from a user's mailbox in Microsoft 365 (formerly Office 365) and log the details of the removed attachments in a Microsoft List (formerly SharePoint List). Here's a breakdown of what the code does, the technologies and tools it uses, and its key features:

## Purpose:

The primary purpose of this script is to remove email attachments from a user's mailbox that are older than a specified date range (in this case, 36 months or 3 years) and are larger than a specified minimum attachment size (in this case, 1 MB). The script also creates a Microsoft List to log the details of the removed attachments, including the email subject, attachment name, attachment size, email creation date, folder name, and email URL.

## Technologies and Tools:

- <b>PowerShell:</b> The script is written in PowerShell, a task automation and configuration management framework from Microsoft.
- <b>Microsoft Graph API:</b> The script utilizes the Microsoft Graph API to interact with Microsoft 365 services, such as retrieving mailbox folders, messages, and attachments, as well as creating and updating Microsoft Lists.
- <b>Azure Active Directory:</b> The script authenticates with Azure Active Directory (Azure AD) to obtain an access token for the Microsoft Graph API.
- <b>Azure Key Vault:</b> The script retrieves the client secret for the Azure AD application from an Azure Key Vault, which is a secure storage solution for application secrets.

## Key Features:

1. <b>Authentication:</b> The script authenticates with Azure AD using an Azure AD application and retrieves an access token for the Microsoft Graph API.
2. <b>Mailbox Iteration:</b> The script iterates through the user's mailbox folders and retrieves messages within the specified date range.
3. <b>Attachment Removal:</b> For each message with attachments, the script checks the attachment size and removes attachments larger than the specified minimum size.
4. <b>Microsoft List Management:<b> The script checks if a Microsoft List with the specified name exists and creates it if it doesn't. It then adds an item to the list for each removed attachment, logging the relevant details.
5. <b>Error Handling:<b> The script includes error handling mechanisms to catch and handle exceptions that may occur during the execution.
6. <b>Configurable Parameters:<b> The script allows configuring various parameters, such as the user's principal name (UPN), maximum number of emails to process per folder, date range, minimum attachment size, Azure AD application details, and Microsoft List name.

## Other Notable Features:

The script uses the <code>Connect-AzAccount</code> and <code>Connect-MgGraph</code> cmdlets to authenticate with Azure and Microsoft Graph, respectively.
It utilizes the Invoke-MgGraphRequest cmdlet to make REST API calls to the Microsoft Graph API.
The script employs various PowerShell cmdlets and techniques, such as creating custom objects, converting objects to JSON, and making REST API calls.

Overall, this PowerShell script is a powerful tool for managing email attachments in Microsoft 365 mailboxes. It leverages the power of the Microsoft Graph API and Azure Active Directory to automate the process of removing large attachments and logging the details in a centralized Microsoft List.

### 🏠 [Homepage](https://github.com/DylanWoodhead/RemoveExchangeEmailAttachments)

## Author

👤 **Dylan Woodhead**

* Github: [@DylanWoodhead](https://github.com/DylanWoodhead)

## Show your support

Give a ⭐️ if this project helped you!