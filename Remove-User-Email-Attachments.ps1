### Variables to change
# User UPN of mailbox to clear
$userPrincipalName = ""

# Amount of months to look back for emails
$months = 24

# Set the minimum attachment size to process (in MB)
$minAttachmentSize = 0

### ONLY change if needed
# Set the Microsoft List name
$listName = "IT-Email-Attachments-Removed-'$userPrincipalName'"

### Do not change anything below this line

# Set the date range to # months ago and before
$dateRange = (Get-Date).AddMonths(-$months).ToString("yyyy-MM-ddTHH:mm:ssZ")
$dateFilter = "createdDateTime le $dateRange"

# Set the maximum number of emails to process per folder
$maxEmailsPerFolder = 500

# Authenticate to Azure
try
{
    "Logging in to Azure..."
    Connect-AzAccount -Identity
}
catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
    break
}

# Set your Azure App registration details
$ApplicationID = ""
$ClientSecret = ""
$TenantID = ""

try{
    $graphtokenBody = @{   
    Grant_Type    = "client_credentials"   
    Scope         = "https://graph.microsoft.com/.default"   
    Client_Id     = $ApplicationID   
    Client_Secret = $ClientSecret.SecretValue | ConvertFrom-SecureString -AsPlainText
    }  

    $graphToken = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Method POST -Body $graphtokenBody | Select-Object -ExpandProperty Access_Token 
    $graphToken = $graphToken | ConvertTo-SecureString -AsPlainText -Force

    "Logging in to Graph..."
    Connect-MgGraph -AccessToken $graphToken

} catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

# Get the user's mailbox folders
$userMailboxFolders = Get-MgUserMailFolder -UserId $userPrincipalName

$displayName = $userPrincipalName.Replace("@email.co.uk","")
$displayName = $displayName.Replace("."," ")
try {
    $getUserSiteId = (Get-MgSite -All | Where {$_.DisplayName -like $displayName} | Select Id).Id
}
catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
    exit
}

$siteId = ($getUserSiteId -split ',')[1]

# Date of script run for output file
$todaysDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")

# Check if the Microsoft List exists, and create it if it doesn't
$existingLists = Get-MgSiteList -SiteId $siteId
if ($existingLists.displayName -notcontains $listName) {
    $newList = [PSCustomObject]@{
        displayName = $listName
        columns = @(
            [PSCustomObject]@{
                name = "Subject"
                text = [PSCustomObject]@{}
            },
            [PSCustomObject]@{
                name = "AttachmentName"
                text = [PSCustomObject]@{}
            },
            [PSCustomObject]@{
                name = "AttachmentSizeMB"
                number = [PSCustomObject]@{}
            },
            [PSCustomObject]@{
                name = "EmailCreatedDate"
                text = [PSCustomObject]@{}
            },
            [PSCustomObject]@{
                name = "FolderName"
                text = [PSCustomObject]@{}
            },
            [PSCustomObject]@{
                name = "EmailURL"
                text = [PSCustomObject]@{}
            },
            [PSCustomObject]@{
                name = "ScriptRunDate"
                text = [PSCustomObject]@{}
            }
        )
        list = [PSCustomObject]@{
            template = "genericList"
        }
    }
        $newListJson = $newList | ConvertTo-Json -Depth 3
        # Using Invoke as New-MgSiteList doesn't work
        $newListId = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists" -Body $newListJson -ContentType "application/json"
    } else {
        Write-Output "List already made"
        $newListId = $existingLists | Where-Object { $_.displayName -eq $listName }
    }

$listURL = Get-MgSiteList -SiteId $siteId -ListId $newListId.Id
$listURL.WebUrl

$ErrorActionPreference = "Stop"

# Iterate through the user's mailbox folders and remove attachments
foreach ($folder in $userMailboxFolders) {
    Write-Output "-- $($folder.DisplayName) --"

    $skipCount = 0
    $totalMessages = 0

    while ($true) {
        $messages = Get-MgUserMailFolderMessage -UserId $userPrincipalName -MailFolderId $folder.Id -Top $maxEmailsPerFolder -Skip $skipCount -Filter $dateFilter
        $totalMessages += $messages.Count

        foreach ($message in $messages) {
            if ($message.HasAttachments) {

                $attachments = Get-MgUserMessageAttachment -UserId $userPrincipalName -MessageId $message.Id

                foreach ($attachment in $attachments) {
                    $attachmentId = $attachment.Id
                    $attachmentSizeMB = [Math]::Round($attachment.Size / 1000000, 3) 
                    if($attachmentSizeMB -ge $minAttachmentSize){
                        try {
                            Write-Output "Deleting attachment in Folder - '$($folder.DisplayName)', Email Subject - '$($message.Subject)' from - '$($message.Sender.EmailAddress.Address)'"
                            Remove-MgUserMessageAttachment -UserId $userPrincipalName -MessageId $message.Id -AttachmentId $attachmentId
    
                            # Add the email details to the Microsoft List
                            $listItemBody = @{
                                fields = @{
                                    Title = $message.Sender.EmailAddress.Name
                                    Subject = $message.Subject
                                    AttachmentName = $attachment.Name 
                                    AttachmentSizeMB = $attachmentSizeMB
                                    EmailCreatedDate = $message.CreatedDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
                                    FolderName = $folder.DisplayName
                                    EmailURL = $message.WebLink
                                    ScriptRunDate = $todaysDate
                                    }
                                }
                            New-MgSiteListItem -SiteId $siteId -ListId $newListId.Id -BodyParameter $listItemBody | Out-Null
                            Start-Sleep -Seconds 1
                        } catch {
                            Write-Error -Message $_.Exception
                            throw $_.Exception
                        }
                    }                                       
                }
            }
        }

        if ($messages.Count -lt $maxEmailsPerFolder) {
            break
        }

        $skipCount += $maxEmailsPerFolder
    }

    Write-Output "Total messages processed in $($folder.DisplayName): $totalMessages"
}