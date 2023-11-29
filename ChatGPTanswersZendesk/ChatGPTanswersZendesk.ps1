<#
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
13 23/11/2023
#>

param (
    [Parameter(Mandatory=$true)][string]$ticketNumber
)
$TicketNumberGPT = $ticketNumber -replace 'chatgptprotocol:', ''

# Global variable to store processed comments
$Global:ProcessedCommentsCache = $null

<#
.SYNOPSIS
Retrieves configuration data from a JSON file.
.DESCRIPTION
This function reads a JSON formatted configuration file located in the same directory as the script and converts it into a PowerShell object. 
It's typically used to manage settings and credentials in a centralized and easily accessible manner.
.EXAMPLE
$config = Get-ConfigData
.NOTES
The configuration file is named 'config.json' and should be located in the same directory as the script. Ensure that the file is properly formatted as valid JSON.
#>
function Get-ConfigData {
    # Define the path to the configuration file relative to the script location
    $Path = "$PSScriptRoot\config.json"

    # Read the configuration file and convert its content from JSON to a PowerShell object
    return Get-Content -Path $Path -Raw | ConvertFrom-Json
}

function Get-Prompt {
    param (
        [string]$Tag
    )
    
    $allPrompts = Get-Content -Path "$PSScriptRoot\prompts.txt" -Raw
    # Updated regex pattern
    $pattern = "\[$Tag\](.*?)(?=\[Prompt_|\[System_|\z)"
    $matches = [regex]::Matches($allPrompts, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    if ($matches.Count -gt 0) {
        return $matches[0].Groups[1].Value.Trim()
    } else {
        Write-Error "Prompt not found for tag: $Tag"
        return $null
    }
}

<#
.SYNOPSIS
Writes a log message to a log file and displays an error message.
.DESCRIPTION
This function appends a provided message to a log file with a timestamp. It also outputs the message as an error in the console. 
The log file is stored at a specified path on the system.
.PARAMETER Message
The message to be logged.
.EXAMPLE
Write-Log "An error occurred while processing data."
.NOTES
The log file is located at "C:\Users\Public\LogFile.log". Ensure that the script has write permissions to this path.
#>
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    # Define the path for the log file.
    $logFilePath = "C:\Users\Public\LogFile.log"

    # Append the message with a timestamp to the log file.
    Add-Content -Path $logFilePath -Value "$(Get-Date) - $Message"

    # Output the message as an error to the console.
    Write-Error $Message
}

<#
.SYNOPSIS
Sends a prompt to the ChatGPT API and retrieves the response.
.DESCRIPTION
This function sends a user-defined prompt to the ChatGPT API and outputs the response from the assistant. 
It requires an API key obtained from a configuration file and handles the construction and execution of the API request. 
Error handling is included to manage potential request failures.
.PARAMETER prompt
The text prompt to be sent to the ChatGPT API.
.PARAMETER model
Specifies the ChatGPT model to be used. If not provided, a default value is used.
.EXAMPLE
$response = Send-ToChatGPT -prompt "How do I reset a router?" -model "gpt-3.5-turbo"
.NOTES
Ensure that the OpenAI API key is correctly configured in the configuration file.
#>
function Send-ToChatGPT {
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
    $systemMessageContent = Get-Prompt -Tag "System_Jarvis"
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
    <#
    $messages = @(
        @{
            "role" = "system"
            "content" = $systemMessageContent
        },
        @{
            "role" = "user"
            "content" = $prompt
        }
    )
    #>
    # Convert the message and other parameters into a JSON-formatted body.
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

<#
.SYNOPSIS
Processes text to prepare it for input to ChatGPT.
.DESCRIPTION
This function performs a series of text processing steps including regex replacements and string substitutions to clean and format the input text. 
It is designed to remove unwanted elements like specific date formats, email addresses, phone numbers, and more, making the text more suitable for processing by ChatGPT.
.PARAMETER InputText
The text to be processed.
.EXAMPLE
$processedText = Process-TextForChatGPT -InputText $rawText
.NOTES
Ensure that the replacements.csv file is properly formatted and located in the same directory as the script.
#>
function Process-TextForChatGPT {
    param(
        [string]$InputText
    )

    # Import simple replacement rules from a CSV file located in the same directory as the script
    $Replacements = Import-Csv (Join-Path $PSScriptRoot "replacements.csv")

    # Regular expressions for specific text patterns removal
    # Remove all data between {JARVIS START and JARVIS END}
    $InputText = [regex]::Replace($InputText, '\{JARVIS START.*?JARVIS END\}', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
    # Remove instances of the date in the format "Date: YYYY-MM-DDTHH:MM:SSZ"
    $InputText = [regex]::Replace($InputText, 'Date:\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z', '')

    # Remove UK phone numbers (with and without the optional "(0)")
    $InputText = [regex]::Replace($InputText, '\+44\s*\(\s*0\s*\)\s*\d{4}\s*\d{6}', '')
    $InputText = [regex]::Replace($InputText, '\+44\s*\d{4}\s*\d{6}', '')

    # Remove email addresses
    $InputText = [regex]::Replace($InputText, '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', '')

    # Remove "Attachments:" line with filenames
    $InputText = [regex]::Replace($InputText, 'Attachments:([^\r\n]+)', '')

    # Remove any strings of repeating text that are more than 30 characters
    $InputText = [regex]::Replace($InputText, '(\b\w{30,}\b)(?:\s+\1\b)+', '$1')

    # Apply string replacements from the CSV file to the input text
    foreach ($replacement in $Replacements) {
        $InputText = $InputText.Replace($replacement.old, $replacement.new)
    }

    # Clean the input text by replacing multiple spaces with a single space and removing non-alphanumeric characters
    $InputText = $InputText -replace '\s+', ' ' -replace '[^a-zA-Z0-9\s.,!?]', ''

    # Apply string replacements again in case any new matches are found after cleaning
    foreach ($replacement in $Replacements) {
     $InputText = $InputText.Replace($replacement.old, $replacement.new)
    }
    $CleanText = $InputText

    Write-Host "Input sanitized." -ForegroundColor White

    # Return the processed text
    return $CleanText.Trim()
}

<#
.SYNOPSIS
Adds an internal comment to a Zendesk ticket.
.DESCRIPTION
This function adds an internal comment to a specified Zendesk ticket. It constructs a comment body and sends it to the Zendesk API. 
The function handles the authentication and submission process.
.PARAMETER TicketNumber
The number of the Zendesk ticket to which the comment will be added.
.PARAMETER Content
The content of the comment to be added to the ticket.
.EXAMPLE
Add-ZendeskInternalComment -TicketNumber 12345 -Content "This is an internal comment."
.NOTES
Ensure that the Zendesk API credentials are properly configured before using this function.
#>
function Add-ZendeskInternalComment {
    param (
        [Parameter(Mandatory=$true)][int]$TicketNumber,
        [Parameter(Mandatory=$true)][string]$Content
    )

    # Retrieve configuration data for Zendesk API.
    $config = Get-ConfigData
    $baseUrl = "https://$($config.Zendesk.Subdomain).zendesk.com/api/v2"
    $authInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($config.Zendesk.Email)/token`:$($config.Zendesk.ApiToken)"))

    # Construct the JSON body for the comment.
    $comment = @{
        ticket = @{
            comment = @{
                body = $Content
                public = $false  # Indicates an internal comment
            }
        }
    }

    # Convert the comment to JSON string.
    $commentJson = ConvertTo-Json -InputObject $comment -Depth 5

    # Set up headers for the API request.
    $commentHeaders = @{
        "Authorization" = "Basic $authInfo"  # Authentication header
        "Content-Type" = "application/json"  # Content type header
    }

    try {
        # Send the comment to Zendesk API.
        $commentResponse = Invoke-WebRequest -Uri "$baseUrl/tickets/$TicketNumber.json" -Method Put -Body $commentJson -Headers $commentHeaders
        if ($commentResponse.StatusCode -eq 200) {
            Write-Host "Internal comment successfully added to ticket #$TicketNumber"
        } else {
            Write-Log "Failed to add internal comment. Status code: $($commentResponse.StatusCode)"
        }
    } catch {
        Write-Output "An error occurred while adding internal comment: $_"
    }
}

<#
.SYNOPSIS
Fetches data from a specified Zendesk API endpoint.
.DESCRIPTION
This function sends a request to a Zendesk API endpoint and returns the response data. It handles errors and returns null if the request fails.
.PARAMETER Uri
The API endpoint URI to send the request to.
.PARAMETER Headers
Headers to be used for the API request, typically including authorization and content type.
.EXAMPLE
$data = Get-ZendeskData -Uri "https://example.zendesk.com/api/v2/tickets/123.json" -Headers $headers
#>
function Get-ZendeskData {
    param (
        [string]$Uri,
        [hashtable]$Headers
    )
    try {
        $response = Invoke-WebRequest -Uri $Uri -Headers $Headers
        return $response.Content | ConvertFrom-Json
    } catch {
        Write-Log "Error fetching data from Zendesk: $_"
        return $null
    }
}

<#
.SYNOPSIS
Maps user IDs to user objects.
.DESCRIPTION
Given an array of user objects, this function creates a hashtable mapping user IDs to their corresponding user objects for quick lookup.
.PARAMETER Users
An array of user objects, typically obtained from a Zendesk API response.
.EXAMPLE
$users = Map-Users -Users $userArray
#>
function Map-Users {
    param (
        [object[]]$Users
    )
    $userMap = @{}
    foreach ($user in $Users) {
        $userMap[$user.id] = $user
    }
    return $userMap
}

<#
.SYNOPSIS
Exports comments from a specific Zendesk ticket.
.DESCRIPTION
This function retrieves comments from a specified Zendesk ticket using the Zendesk API. It processes each comment and formats them for export. 
The function handles the retrieval of ticket details, user information, and comment processing.
.PARAMETER TicketNumber
The number of the Zendesk ticket from which to export comments.
.EXAMPLE
$comments = Export-ZendeskTicketComments -TicketNumber 12345
#>
function Export-ZendeskTicketComments {
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

        $ticketData = Get-ZendeskData -Uri "$baseUrl/tickets/$TicketNumber.json?include=users,organizations" -Headers $headers
    $ticket = $ticketData.ticket
    $users = Map-Users -Users $ticketData.users

    $commentsData = Get-ZendeskData -Uri "$baseUrl/tickets/$TicketNumber/comments.json?include=users" -Headers $headers
    $comments = $commentsData.comments
    $attachments = @{}

    # Initialize an array to hold the processed comments.
   $processedComments = @("Requester: $($users[$ticket.requester_id].name)", "Assignee: $(if ($ticket.assignee_id) { $users[$ticket.assignee_id].name } else { 'Not assigned' })", "Ticket Subject: $($ticket.subject)")

    foreach ($comment in $comments) {
        $processedComment = Process-Comment -Comment $comment -Users $users -Attachments $attachments
        $processedComments += $processedComment
        $processedComments += "----"
    }

    return $processedComments -join "`n"
}

<#
.SYNOPSIS
Processes and displays ticket comments for input to ChatGPT.
.DESCRIPTION
This function takes ticket comments, processes them for formatting and cleaning, and then constructs a prompt for ChatGPT. 
It displays the processing status and returns the response from ChatGPT.
.PARAMETER Comments
An array of comments to be processed and sent to ChatGPT.
.PARAMETER AdditionalText
Additional text to prepend to the processed comments when constructing the ChatGPT prompt.
.PARAMETER Model
Specifies the ChatGPT model to be used. If not provided, defaults to the model used in Send-ToChatGPT function.
.EXAMPLE
$response = Process-AndDisplayContent -Comments $commentsArray -AdditionalText "Please analyze these comments:" -Model "gpt-3.5-turbo"
#>
function Process-AndDisplayContent {
    param (
        [Parameter(Mandatory=$true)][array]$Comments,
        [Parameter(Mandatory=$true)][string]$AdditionalText,
        [string]$Model
    )

    # Join comments into a single string
    $joinedContent = $Comments -join "`n"
    
    Write-Host "GPT pre-processing started for $Model" -ForegroundColor Yellow
    
    # Process the content for ChatGPT
    $processedContent = Process-TextForChatGPT -InputText $joinedContent
    
    Write-Host "GPT pre-processing finished for $Model" -ForegroundColor Yellow

    # Create a JSON object from the processed content
    $ticketObject = @{ content = $processedContent }
    $jsonTicketContent = $ticketObject | ConvertTo-Json -Compress

    # Construct the ChatGPT prompt
    $gptPrompt = $AdditionalText + "`n" + $jsonTicketContent
    Write-Host $gptPrompt
    # Send the prompt to ChatGPT and return the response
    return Send-ToChatGPT -prompt $gptPrompt -model $Model
}

<#
    This function integrates ChatGPT with Zendesk ticketing. 
    It exports Zendesk ticket comments, sanitizes the content, and reduces it to a manageable size for ChatGPT. 
    The function then creates a prompt for ChatGPT using the sanitized content and a user-provided question. 
    The response from ChatGPT is formatted based on the specified mode and outputted to a file. 
    The function can also optionally add the response as an internal comment in Zendesk.
    #>
function ChatGPTanswersZendesk {
    # Define mandatory parameters: the question to ask, the Zendesk ticket number, and the mode (technician/customer).
    param (
        [Parameter(Mandatory=$true)][int]$TicketNumber,
        [Parameter(Mandatory=$true)][ValidateSet("technician", "customer")][string]$Mode
    )
    if (-not $Global:ProcessedCommentsCache) {
                # Retrieve and process the ticket comments if not already done
        #$sanitiseq = Get-Prompt -Tag "Prompt_Sanitise"
        $sanitiseq = @'
Process the dataset of email communications by extracting specific details from each email, including the requester and body of the message. It is critically essential and an absolute mandate that all email signatures which includes surnames, position, telephone numbers, addresses, usernames, passwords and disclaimers are rigorously excluded from the dataset. This exclusion is a cornerstone requirement for strict compliance with privacy laws and regulations. Under no circumstances should these elements be included. This directive is of the highest priority and is to be adhered to with utmost diligence. Any oversight or deviation in this regard will be a direct violation of privacy and compliance protocols and is entirely unacceptable. Do not add any explanations or notes about the process or reasons for data exclusion. Remove any references to empty emails. Present the extracted information in a straightforward format without any introductory or concluding remarks. Simply list the details for each email, labeled as 'Email 1', 'Email 2', etc., with the relevant information under each label. 
'@
        $processedComments = Export-ZendeskTicketComments -TicketNumber $TicketNumber
        Write-Host "Ticket Exported." -ForegroundColor White
        Write-Host "Preparing for GPT3.5 sanitisation." -ForegroundColor White
        $Global:ProcessedCommentsCache = Process-AndDisplayContent -Comments $processedComments -AdditionalText $sanitiseq -Model "gpt-3.5-turbo-16k"
    }

    # Define the question based on the mode
    <#
    $question = if ($Mode -eq "technician") { 
        Get-Prompt -Tag "Prompt_Technician"
    } else { 
        Get-Prompt -Tag "Prompt_Customer"
    }
    #>
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
    #Add-ZendeskInternalComment -TicketNumber $TicketNumber -Content $fullResponse
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
Write-Output "beta version"
Start-Sleep -Seconds 10