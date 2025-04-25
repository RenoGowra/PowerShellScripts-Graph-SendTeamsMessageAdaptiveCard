# Required Modules
Import-Module ImportExcel
Import-Module Microsoft.Graph.Teams
Import-Module Microsoft.Graph.Authentication

# Auth
$scopes = @("Chat.ReadWrite", "User.Read")
Connect-MgGraph -Scopes $scopes
$myUserAccount = (Get-MgContext).Account
$tenantId = (Get-MgContext).TenantId

# Input Excel file
$excelFilePath = "C:\Users\regowra\Desktop\Users.xlsx"
$upns = Import-Excel -Path $excelFilePath | Select-Object -ExpandProperty UPN

# Output log file
$logFile = "C:\Users\regowra\Desktop\CopilotMessageLog.csv"
if (-not (Test-Path $logFile)) {
    "Timestamp,UPN,Status,Message" | Out-File -FilePath $logFile
}

# Loop through each user
foreach ($upn in $upns) {
    try {
        Write-Host "Sending to $upn..." -ForegroundColor Cyan

        # Create chat
        $params = @{
            chatType = "oneOnOne"
            members = @(
                @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    roles = @("owner")
                    "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$upn')"
                },
                @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    roles = @("owner")
                    "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$myUserAccount')"
                }
            )
        }

        $chat = New-MgChat -BodyParameter $params
        $chatId = $chat.Id
        $attachmentId = [guid]::NewGuid().ToString()

        # Adaptive Card JSON (Replace $tenantId with actual tenant ID)
        $adaptiveCardJson = @"
        {
            "`$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Discover the power of M365 Copilot!",
                            "wrap": true,
                            "size": "Medium",
                            "weight": "Bolder",
                            "color": "Accent"
                        },
                        {
                            "type": "TextBlock",
                            "text": "It looks like you haven't used it in a while, or maybe not at all. There are many ways you can use Copilot to make your workday a breeze, like:\n* Catching up on meetings\n* Writing an email\n* Getting a work question answered\n* Drafting a document\n* Transforming a document into a presentation\n* Finding a file",
                            "wrap": true,
                            "spacing": "Small"
                        }
                    ]
                },
                {
                    "type": "Container",
                    "separator": true,
                    "spacing": "Medium",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Here are some prompts you can try in M365 Copilot. Copy and use them in Copilot Business Chat by clicking on Try Now",
                            "wrap": true,
                            "size": "Medium",
                            "weight": "Bolder"
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Start Your Day:**",
                            "wrap": true,
                            "weight": "Bolder",
                            "spacing": "Medium"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "auto",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": "https://support.content.office.net/en-us/media/af6441ee-2e11-4ca8-9f1d-a1a3ae557dfc.png",
                                            "size": "Small"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "stretch",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "- Based on my emails and Teams chat channels of the last 36 hours, organize my tasks along the Eisenhower matrix for the coming 15 days.\n- What is the common perception of me based on my past 1-month emails (or chats) and what do you think I can improve in my email/chat writing style?",
                                            "wrap": true,
                                            "spacing": "Small"
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Catch Up on Emails:**",
                            "wrap": true,
                            "weight": "Bolder",
                            "spacing": "Medium"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "auto",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": "https://support.content.office.net/en-us/media/e6eecd30-cc50-477a-be9c-2ddea72012fe.png",
                                            "size": "Small"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "stretch",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "- Can you summarize the unread emails in my inbox?\n- Show me any urgent emails I need to respond to.",
                                            "wrap": true,
                                            "spacing": "Small"
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Gather Information:**",
                            "wrap": true,
                            "weight": "Bolder",
                            "spacing": "Medium"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "auto",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": "https://support.content.office.net/en-us/media/960dcbae-7561-44b8-b876-55045fbb9930.png",
                                            "size": "Small"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "stretch",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "- Can you find the latest sales report and summarize the key points?\n- Summarize the meeting /<Meeting Title here>",
                                            "wrap": true,
                                            "spacing": "Small"
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ],
            "actions": [
                {
                    "type": "Action.OpenUrl",
                    "title": "Try Now",
                    "url": "https://teams.microsoft.com/l/entity/b5abf2ae-c16b-4310-8f8a-d3bcdb52f162/conversations?tenantId=$tenantId",
                    "style": "positive"
                }
            ],
            "verticalContentAlignment": "Center"
        }
"@

        # Prepare message body with attachment marker
        $body = @{
            contentType = "html"
            content = "<attachment id='$attachmentId'></attachment>"
        }

        # Build the card attachment
        $attachment = @{
            id = $attachmentId
            contentType = "application/vnd.microsoft.card.adaptive"
            content = $adaptiveCardJson
        }

        # Final message object
        $message = @{
            body = $body
            attachments = @($attachment)
        }

        # Send message
        New-MgChatMessage -ChatId $chatId -Body $message.body -Attachments $message.attachments

        # Log success
        "$((Get-Date).ToString('u')),$upn,Success,Message sent" | Out-File -FilePath $logFile -Append
        Write-Host "Message sent to $upn" -ForegroundColor Green
    }
    catch {
        # Log failure
        "$((Get-Date).ToString('u')),$upn,Error,$($_.Exception.Message)" | Out-File -FilePath $logFile -Append
        Write-Host "Failed to send to $upn - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Cleanup
Disconnect-MgGraph
