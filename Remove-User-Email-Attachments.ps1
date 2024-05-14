# User UPN of mailbox to clear
$userPrincipalName = "rhys.lavender@propelfinance.co.uk"

# Set the maximum number of emails to process per folder
$maxEmailsPerFolder = 500

# Set the date range to # months ago and before
$dateRange = (Get-Date).AddMonths(-36).ToString("yyyy-MM-ddTHH:mm:ssZ")
$dateFilter = "createdDateTime le $dateRange"

# Set the minimum attachment size to process (in MB)
$minAttachmentSize = 3

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
$ApplicationID = "f7a0dfc4-17ec-4229-91c3-3994f0d9429f"
$ClientSecret = Get-AzKeyVaultSecret -VaultName "IT-Access" -Name "Graph-API-Removing-Email-Attachments" -AsPlainText
$TenantID = "3e833290-3714-42e2-a81f-a6606b883185"

try{
    $graphtokenBody = @{   
    Grant_Type    = "client_credentials"   
    Scope         = "https://graph.microsoft.com/.default"   
    Client_Id     = $ApplicationID   
    Client_Secret = $ClientSecret
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
$headers = @{
    "Authorization" = "Bearer $graphtoken"
}

$userMailboxFolders = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userPrincipalName/mailFolders" -Headers $headers

$getUserSiteId = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userPrincipalName/drive/SharePointIds" -Headers $headers

# Date of script run for output file
$todaysDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")

# Set the Microsoft List information
$listName = "IT-Email-Attachments-Removed-'$userPrincipalName'"
$siteId = $getUserSiteId.siteId

# Check if the Microsoft List exists, and create it if it doesn't
$existingLists = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists" -Headers $headers
if ($existingLists.Value.displayName -notcontains $listName) {
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
        }
    )
    list = [PSCustomObject]@{
        template = "genericList"
    }
}
    $newListJson = $newList | ConvertTo-Json -Depth 3
    $newListId = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists" -Body $newListJson -Headers $headers -ContentType "application/json"
} else {
    Write-Output "List already made"
    $newListId = $existingLists.Value | Where-Object { $_.displayName -eq $listName }
}

$listURL = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$($newListId.id)" -Body ($listItemBody | ConvertTo-Json) -Headers $headers -ContentType "application/json"
$listURL.webUrl

$ErrorActionPreference = "Stop"

# Iterate through the user's mailbox folders and remove attachments
foreach ($folder in $userMailboxFolders.Value) {
    Write-Output "-- $($folder.displayName) --"

    $skipCount = 0
    $totalMessages = 0

    while ($true) {
        $messages = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userPrincipalName/mailFolders/$($folder.Id)/messages`?`$top=$maxEmailsPerFolder&`$skip=$skipCount&`$filter=$dateFilter" -Headers $headers
        $totalMessages += $messages.Value.Count

        foreach ($message in $messages.Value) {
            if ($message.HasAttachments) {

                $attachments = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userPrincipalName/mailFolders/$($folder.Id)/messages/$($message.Id)/attachments" -Headers $headers

                foreach ($attachment in $attachments.Value) {
                    $attachmentId = $attachment.Id
                    $deleteAttachmentUri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/messages/$($message.Id)/attachments/$attachmentId"
                    $attachmentSizeMB = [Math]::Round($attachment.size / 1000000, 3) 
                    if($attachmentSizeMB -ge $minAttachmentSize){
                        try {
                            Write-Output "Deleting attachment in Folder - '$($folder.displayName)', Email Subject - '$($message.subject)' from - '$($message.sender.Values.address)'"
                            Invoke-MgGraphRequest -Method DELETE -Uri $deleteAttachmentUri -Headers $headers
    
                            # Add the email details to the Microsoft List
                            $listItemBody = @{
                                fields = @{
                                    Title = $message.sender.Values.name
                                    Subject = $message.subject
                                    AttachmentName = $attachment.Name 
                                    AttachmentSizeMB = $attachmentSizeMB
                                    EmailCreatedDate = $message.createdDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
                                    FolderName = $folder.displayName
                                    EmailURL = $message.webLink
                                    }
                                }
                                } catch {
                                    Write-Error -Message $_.Exception
                                    throw $_.Exception
                                }
                                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$($newListId.id)/items" -Body ($listItemBody | ConvertTo-Json) -Headers $headers -ContentType "application/json" | Out-Null
                                Start-Sleep -Seconds 1
                    }                                       
                }
            }
        }

        if ($messages.Value.Count -lt $maxEmailsPerFolder) {
            break
        }

        $skipCount += $maxEmailsPerFolder
    }

    Write-Output "Total messages processed in $($folder.displayName): $totalMessages"
}