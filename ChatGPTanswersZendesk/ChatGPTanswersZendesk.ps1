﻿<#
.SYNOPSIS
This script interacts with the OpenAI API, processes content, and integrates with Zendesk for ticket handling.

.DESCRIPTION
The script includes functions to load configuration, send requests to ChatGPT, sanitize content, 
reduce tokens, manage Zendesk ticket comments, and more. It requires a valid OpenAI API key and Zendesk credentials.

.PREREQUISITES
- PowerShell 5.1 or higher.
- Access to OpenAI API and Zendesk.
- Config.json and CSV files for replacements in the script directory.

.USAGE
Run the script with mandatory parameters like ticket number. Parameters can be passed via command line or hard-coded for testing.

.AUTHOR
Dan White

.LAST MODIFIED
13:04 23/11/2023
#>

param (
    [Parameter(Mandatory=$true)][string]$ticketNumber
)
$TicketNumberGPT = $ticketNumber -replace 'chatgptprotocol:', ''

# Global variable to store processed comments
$Global:ProcessedCommentsCache = $null

function Get-ConfigData {
    $Path = "$PSScriptRoot\config.json"
    return Get-Content -Path $Path -Raw | ConvertFrom-Json
}

function Write-Log {
    param([string]$Message)
    Add-Content -Path "C:\Users\Public\LogFile.log" -Value "$(Get-Date) - $Message"
}

function Write-ConditionalHostMessage {
    param(
        [Parameter(Mandatory = $true)]
        $VariableToCheck,

        [Parameter(Mandatory = $true)]
        [string]$MessageIfNotEmpty,

        [Parameter(Mandatory = $true)]
        [string]$MessageIfEmpty,

        [Parameter(Mandatory = $true)]
        [System.ConsoleColor]$ForegroundColor
    )

    if (![string]::IsNullOrEmpty($VariableToCheck)) {
        Write-Host $MessageIfNotEmpty -ForegroundColor $ForegroundColor
    } else {
        Write-Host $MessageIfEmpty -ForegroundColor $ForegroundColor
    }
}


function Send-ToChatGPT {
    <#
    This function is designed to send a prompt to the ChatGPT API and output the response. 
    It includes error handling and requires an API key from a configuration file.
    #>

    # Define parameters for the function.
    # 'prompt' is a mandatory string parameter with an alias 'p'.
    Param(
        [Parameter(Mandatory=$true, HelpMessage='prompt:')]
        [Alias('p')]
        [string]$prompt,
        [string]$model
    )

    # Configuration for the API request.
    $temperature = 0.5
    $maxTokens = 4096

    # Retrieve the API key from the configuration data.
    $config = Get-ConfigData
    $apiKey = $config.OpenAI.ApiKey

    # Prepare headers for the API request including Content-Type and Authorization.
    $headers = @{
        "Content-Type" = "application/json"
        "Authorization" = "Bearer $apiKey"
    }
    
    # Define the body of the message to send to the ChatGPT API.
    # Includes a system message setting the context and a user message containing the prompt.
    $messages = @(
        @{
            "role" = "system"
            "content" = "You are Jarvis, an extremely skilled and adept IT support expert for SMEs. With your comprehensive IT knowledge, problem-solving, certifications, and communication skills, tackle the tickets in the directed manner. If you are referred to directly as Jarvis in an internal ticket body, please answer the query as if you were asked directly as a primary question."
        },
        @{
            "role" = "user"
            "content" = $prompt
        }
    )

    # Convert the message and other parameters into a JSON-formatted body for the API request.
    $body = @{
        "model" = $model
        "messages" = $messages
        "temperature" = $temperature
        "max_tokens" = $maxTokens
    } | ConvertTo-Json

    # API endpoint URL for ChatGPT completions.
    $url = "https://api.openai.com/v1/chat/completions"

    try {
        # Make the API request and handle the response.
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body
        # Extract and output the message from the assistant.
        $assistantMessage = $response.choices[0].message.content
        Write-Output $assistantMessage.Trim()
    } catch {
        # Error handling: log the error and output a failure message.
        Write-Log "Error sending request: $_"
        Write-Output "Failed to send request to ChatGPT."
    }
}



function Process-TextForChatGPT {
    param(
        [string]$InputText
    )

    # Import simple replacement rules from a CSV file
    $Replacements = Import-Csv (Join-Path $PSScriptRoot "replacements.csv")
    #Write-Host $Replacements -ForegroundColor Cyan

    # Remove all data between {JARVIS START and JARVIS END}
    $InputText = [regex]::Replace($InputText, '\{JARVIS START.*?JARVIS END\}', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
    # Remove instances of the date in the format "Date: YYYY-MM-DDTHH:MM:SSZ"
    $InputText = [regex]::Replace($InputText, 'Date:\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z', '')

    # Replace hyperlinks with just the domain name
    #$InputText = [regex]::Replace($InputText, '(http|https)://(www\.)?([^/\s]+)', '$2$3')
    
    # Remove UK phone numbers (with and without the optional "(0)")
    $InputText = [regex]::Replace($InputText, '\+44\s*\(\s*0\s*\)\s*\d{4}\s*\d{6}', '')
    $InputText = [regex]::Replace($InputText, '\+44\s*\d{4}\s*\d{6}', '')

    # Remove email addresses
    $InputText = [regex]::Replace($InputText, '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', '')

    # Remove "Attachments:" line with filenames
    $InputText = [regex]::Replace($InputText, 'Attachments:([^\r\n]+)', '')

    # Remove any strings of repeating text that are more than 30 characters
    $InputText = [regex]::Replace($InputText, '(\b\w{30,}\b)(?:\s+\1\b)+', '$1')

    # Apply string replacements to the input text
    foreach ($replacement in $Replacements) {
        $InputText = $InputText.Replace($replacement.old, $replacement.new)
    }

    # Clean the input text (adjust regex as needed)
    $InputText = $InputText -replace '\s+', ' ' -replace '[^a-zA-Z0-9\s.,!?]', ''

    # Apply string replacements to the input text - second pass
    foreach ($replacement in $Replacements) {
     $InputText = $InputText.Replace($replacement.old, $replacement.new)
    }
    $CleanText = $InputText
    Write-Host "Input sanitised." -ForegroundColor White
    # Return the processed text
    return $CleanText.Trim()
}


function Add-ZendeskInternalCommentWithAttachment {
    param (
        [Parameter(Mandatory=$true)][int]$TicketNumber,
        [Parameter()][string]$Content,
        [Parameter()][string]$AttachmentFilePath,
        [Parameter()][bool]$AsAttachment = $false
    )

    $config = Get-ConfigData
    $baseUrl = "https://$($config.Zendesk.Subdomain).zendesk.com/api/v2"
    $authInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($config.Zendesk.Email)/token`:$($config.Zendesk.ApiToken)"))

    if ($AsAttachment -and $AttachmentFilePath) {
        $attachmentHeaders = @{
            "Authorization" = "Basic $authInfo"
            "Content-Type" = "application/binary"
        }
        $attachmentResponse = Invoke-WebRequest -Uri "$baseUrl/uploads.json?filename=$([System.IO.Path]::GetFileName($AttachmentFilePath))" -Method Post -InFile $AttachmentFilePath -Headers $attachmentHeaders

        if ($attachmentResponse.StatusCode -eq 201) {
            $attachmentContent = $attachmentResponse.Content | ConvertFrom-Json
            $attachmentToken = $attachmentContent.upload.token
        } else {
            Write-Error "Failed to upload the attachment. Status code: $($attachmentResponse.StatusCode)"
            return
        }
    }

    # Construct the comment body.
    $comment = @{
        ticket = @{
            comment = @{
                body = $Content
                public = $false
            }
        }
    }

    # Add attachment token if there's an attachment.
    if ($AsAttachment -and $AttachmentFilePath) {
        $comment.ticket.comment.uploads = @($attachmentToken)
    }

    # Convert the comment to JSON.
    $commentJson = ConvertTo-Json -InputObject $comment -Depth 5

    $commentHeaders = @{
        "Authorization" = "Basic $authInfo"
        "Content-Type" = "application/json"
    }

    try {
        $commentResponse = Invoke-WebRequest -Uri "$baseUrl/tickets/$TicketNumber.json" -Method Put -Body $commentJson -Headers $commentHeaders
        if ($commentResponse.StatusCode -eq 200) {
            Write-Output "Internal comment successfully added to ticket #$TicketNumber"
        } else {
            Write-Error "Failed to add internal comment. Status code: $($commentResponse.StatusCode)"
        }
     } catch {
        Write-Output "An error occurred while adding internal comment."
     }
}


function Export-ZendeskTicketComments {
    <#
    This function exports comments from a specific Zendesk ticket to a text file. 
    It uses the Zendesk API to fetch ticket details and comments, processes each comment, and formats them for export. 
    The nested Process-Comment function handles the individual comment processing, including formatting and attachment handling. 
    The main function sets up the necessary API request details, retrieves ticket and user information, and writes the processed comments to an output file.
    #>

    # Define a mandatory parameter for the Zendesk ticket number.
    param (
        [Parameter(Mandatory=$true)][int]$TicketNumber
    )

    # Retrieve configuration data for Zendesk API.
    $config = Get-ConfigData
    $baseUrl = "https://$($config.Zendesk.Subdomain).zendesk.com/api/v2"
    $authInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($config.Zendesk.Email)/token`:$($config.Zendesk.ApiToken)"))

    # Nested function to process individual comments.
    function Process-Comment {
        param (
            $Comment,
            $Users,
            $Signatures,
            $Attachments
        )

        # Extract and format comment details.
        $from = $Users[$Comment.author_id].name
        $date = $Comment.created_at
        $commentType = if ($Comment.public) { "Public" } else { "Internal" }
        $body = $Comment.body

        # Create a list of attachments for the comment.
        $attachmentList = @()
        foreach ($attachment in $Comment.attachments) {
            if ($Attachments.ContainsKey($attachment.id)) {
                $attachmentList += $Attachments[$attachment.id].file_name
            } else {
                $Attachments.Add($attachment.id, $attachment)
                $attachmentList += $attachment.file_name
            }
        }
        $attachmentText = if ($attachmentList) { $attachmentList -join ', ' } else { "" }

        # Format and return the processed comment.
        @"
From: $from
Date: $date
Type: $commentType
Body: $body
Attachments: $attachmentText
"@
    }

    # Set up headers for the API request.
    $headers = @{
        "Authorization" = "Basic $authInfo"
        "Content-Type" = "application/json"
    }

    # Retrieve ticket and related data from Zendesk.
    $ticketResponse = Invoke-WebRequest -Uri "$baseUrl/tickets/$TicketNumber.json?include=users,organizations" -Headers $headers
    $ticket = ($ticketResponse.Content | ConvertFrom-Json).ticket
    $users = @{}

    # Map user data for easy reference.
    foreach ($user in ($ticketResponse.Content | ConvertFrom-Json).users) {
        $users.Add($user.id, $user)
    }

    # Retrieve comments for the ticket.
    $commentsResponse = Invoke-WebRequest -Uri "$baseUrl/tickets/$TicketNumber/comments.json?include=users" -Headers $headers
    $comments = ($commentsResponse.Content | ConvertFrom-Json).comments

    # Initialize an empty hashtable for attachments.
    $attachments = @{}

    # Initialize an array to hold the processed comments.
    $processedComments = @()

    # Add the header to the processed comments.
    $header = @"
Requester: $($users[$ticket.requester_id].name)
Assignee: $(if ($ticket.assignee_id -ne $null) { $users[$ticket.assignee_id].name } else { "Not assigned" })
Ticket Subject: $($ticket.subject)
"@
    $processedComments += $header

    # Process and add each comment to the processed comments array.
    foreach ($comment in $comments) {
        $processedComment = Process-Comment -Comment $comment -Users $users -Signatures $signatures -Attachments $attachments
        $processedComments += $processedComment
        $processedComments += "----"
    }

    # Return the array of processed comments.
    return $processedComments
}

function Process-AndDisplayContent {
    param (
        [Parameter(Mandatory=$true)][array]$Comments,
        [Parameter(Mandatory=$true)][string]$AdditionalText,
        [string]$Model
    )

    $joinedContent = $Comments -join "`n"
    #Write-Host $joinedContent -ForegroundColor Red
    
    Write-Host "GPT pre-processing started for $Model" -ForegroundColor Yellow
    $processedContent = Process-TextForChatGPT -InputText $joinedContent
    #Write-Host $processedContent -ForegroundColor Yellow
    Write-Host "GPT pre-processing finished for $Model" -ForegroundColor Yellow

    $ticketObject = @{ content = $processedContent }
    $jsonTicketContent = $ticketObject | ConvertTo-Json -Compress
    #Write-Host $jsonTicketContent -ForegroundColor White


    $gptPrompt = $AdditionalText + "`n" + $jsonTicketContent
    return Send-ToChatGPT -prompt $gptPrompt -model $Model
}

function ChatGPTanswersZendesk {
    <#
    This function integrates ChatGPT with Zendesk ticketing. 
    It exports Zendesk ticket comments, sanitizes the content, and reduces it to a manageable size for ChatGPT. 
    The function then creates a prompt for ChatGPT using the sanitized content and a user-provided question. 
    The response from ChatGPT is formatted based on the specified mode and outputted to a file. 
    The function can also optionally add the response as an internal comment in Zendesk.
    #>

    # Define mandatory parameters: the question to ask, the Zendesk ticket number, and the mode (technician/customer).
    param (
        [Parameter(Mandatory=$true)][int]$TicketNumber,
        [Parameter(Mandatory=$true)][ValidateSet("technician", "customer")][string]$Mode
    )

    if (-not $Global:ProcessedCommentsCache) {
                # Retrieve and process the ticket comments if not already done
        $sanitiseq = @'
Process the dataset of email communications by extracting specific details from each email, including the requester and body of the message. It is critically essential and an absolute mandate that all email signatures which includes surnames, position, telephone numbers, addresses, usernames, passwords and disclaimers are rigorously excluded from the dataset. This exclusion is a cornerstone requirement for strict compliance with privacy laws and regulations. Under no circumstances should these elements be included. This directive is of the highest priority and is to be adhered to with utmost diligence. Any oversight or deviation in this regard will be a direct violation of privacy and compliance protocols and is entirely unacceptable. Do not add any explanations or notes about the process or reasons for data exclusion. Remove any references to empty emails. Present the extracted information in a straightforward format without any introductory or concluding remarks. Simply list the details for each email, labeled as 'Email 1', 'Email 2', etc., with the relevant information under each label. 
'@
        $processedComments = Export-ZendeskTicketComments -TicketNumber $TicketNumber
        Write-Host "Ticket Exported." -ForegroundColor White
        Write-Host "Preparing for GPT3.5 sanitisation." -ForegroundColor White
        $Global:ProcessedCommentsCache = Process-AndDisplayContent -Comments $processedComments -AdditionalText $sanitiseq -Model "gpt-3.5-turbo-16k"
    }

    # Define the question based on the mode
    $question = if ($Mode -eq "technician") { @'
Provide concise internal guidance in a mentor role for the IT support technicians addressing this issue. Offer clear, actionable advice and troubleshooting steps, format in Markdown (do not answer within a codeblock) for quick reference. Assume technician competence and that they have knowledge of the customer’s IT infrastructure and our standard procedures, so only detail complicated or unusual tasks. Short lists of best practice steps or highlighting potential problems during any procedure is encouraged. Include relevant insights or considerations based on the customer’s known environment. Refer to our RMM tool, Atera where necessary to get tasks done, and you can recommend practical PowerShell scripts (within a markdown codeblock but do not specify Powershell after the backticks) when appropriate to diagnose or resolve issues. Refer to ITGlue for documentation, and gently remind technicians to document any changes to infrastructure where appropriate. Conclude with a brief 'Harvest Notes field:' entry that encapsulates the main goal or action of the ticket without referring to the tools used or the customer company for efficient time tracking and management. Commence with ticket support direction.
'@} else { @'
Respond to customers ticket -who is the requester- in clear, non-technical language. You are well-acquainted with the customers technical setup and preferences. Provide solutions that are thorough yet explained in an accessible manner, avoiding technical jargon. Offer responses that are concise, informative, and tailored to the customer’s understanding, ensuring all communication is easy to grasp while still capturing all necessary details for the task. Use all the information but only respond to the latest comments - do not write responses to previous emails. Format in Markdown if appropriate.Do not ever claim to be a person and do not use a signature. Proceed with customer-focused ticket resolution.
'@
    }
    Write-Host "Submitting $Mode question to Jarvis." -ForegroundColor Red
    $gptResponse = Process-AndDisplayContent -Comments $Global:ProcessedCommentsCache -AdditionalText $Question -Model "gpt-4-1106-preview"
    $responseTemplate = if ($Mode -eq "customer") { '## Customer answer template (formatted in markdown)' } else { '## Technician notes' }
    $fullResponse = "{JARVIS START`n" + $responseTemplate + "`n`n" + $gptResponse + "`nJARVIS END}"
    
    #Write-Host $fullResponse -ForegroundColor Green
    Write-Host "Jarvis has an answer." -ForegroundColor Green
    Add-ZendeskInternalCommentWithAttachment -TicketNumber $TicketNumber -Content $fullResponse -AsAttachment $false
    Write-Host "Jarvis has put in a ticket." -ForegroundColor Green
}

cls
Write-Host "Processing ticket $TicketNumberGPT" -ForegroundColor Cyan
# Technician answer: Process as technician answer.
ChatGPTanswersZendesk -TicketNumber $TicketNumberGPT -mode "technician"

#Uncomment the line below if we start hitting rate limits.
#Start-Sleep -Seconds 60

#Customer answer: Process as customer answer.
ChatGPTanswersZendesk -TicketNumber $TicketNumberGPT -mode "customer"

#10 second delay to allow window to be seen. Remove if you want.
Start-Sleep -Seconds 10