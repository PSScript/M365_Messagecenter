#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Digest Email Sender - Modular HTML Email System with OAuth/Basic Auth
.DESCRIPTION
    Professional email sending system with:
    - Separate HTML template processing
    - Inline image embedding with AlternateView
    - OAuth2 and Basic Auth support
    - Batched sending with rate limiting
    - Checkpoint-based resume capability
    - Comprehensive error handling and retry logic
.AUTHOR
    Jan Hübener
.VERSION
    1.0.0
#>

# ============================================================================
# MODULE: Email Authentication
# ============================================================================

function Get-EmailAuthenticationCredential {
    <#
    .SYNOPSIS
        Creates SMTP credential object based on authentication method
    .PARAMETER AuthMethod
        Authentication method: 'Basic' or 'OAuth'
    .PARAMETER Username
        SMTP username (email address)
    .PARAMETER Password
        Password for Basic auth
    .PARAMETER TenantId
        Azure AD Tenant ID for OAuth
    .PARAMETER ClientId
        Azure AD Application (Client) ID for OAuth
    .PARAMETER ClientSecret
        Azure AD Application Secret for OAuth
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Basic', 'OAuth')]
        [string]$AuthMethod,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(ParameterSetName = 'Basic', Mandatory = $true)]
        [string]$Password,

        [Parameter(ParameterSetName = 'OAuth', Mandatory = $true)]
        [string]$TenantId,

        [Parameter(ParameterSetName = 'OAuth', Mandatory = $true)]
        [string]$ClientId,

        [Parameter(ParameterSetName = 'OAuth', Mandatory = $true)]
        [string]$ClientSecret
    )

    switch ($AuthMethod) {
        'Basic' {
            Write-Verbose "Creating Basic Authentication credential for $Username"
            $securePass = ConvertTo-SecureString $Password -AsPlainText -Force
            return New-Object System.Management.Automation.PSCredential($Username, $securePass)
        }
        
        'OAuth' {
            Write-Verbose "Acquiring OAuth2 token for $Username"
            try {
                $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
                
                $body = @{
                    client_id     = $ClientId
                    client_secret = $ClientSecret
                    scope         = "https://outlook.office365.com/.default"
                    grant_type    = "client_credentials"
                }

                $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
                
                # For OAuth, we use the access token as password with XOAUTH2
                $oauthToken = $response.access_token
                $secureToken = ConvertTo-SecureString $oauthToken -AsPlainText -Force
                
                Write-Verbose "OAuth token acquired successfully"
                return New-Object System.Management.Automation.PSCredential($Username, $secureToken)
            }
            catch {
                throw "Failed to acquire OAuth token: $($_.Exception.Message)"
            }
        }
    }
}

# ============================================================================
# MODULE: HTML Template Processing
# ============================================================================

function Get-ProcessedHtmlTemplate {
    <#
    .SYNOPSIS
        Loads and processes HTML template with placeholder replacement
    .PARAMETER TemplatePath
        Path to HTML template file
    .PARAMETER Replacements
        Hashtable of placeholder-value pairs for replacement
    .PARAMETER Encoding
        File encoding (default: UTF8)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ })]
        [string]$TemplatePath,

        [Parameter(Mandatory = $false)]
        [hashtable]$Replacements = @{},

        [Parameter(Mandatory = $false)]
        [string]$Encoding = 'UTF8'
    )

    try {
        Write-Verbose "Loading HTML template from: $TemplatePath"
        $htmlContent = Get-Content -LiteralPath $TemplatePath -Raw -Encoding $Encoding

        # HTML-encode all replacement values to prevent injection
        Add-Type -AssemblyName System.Web
        foreach ($key in $Replacements.Keys) {
            $value = $Replacements[$key]
            $encodedValue = [System.Web.HttpUtility]::HtmlEncode($value)
            $htmlContent = $htmlContent.Replace($key, $encodedValue)
        }

        Write-Verbose "Template processed with $($Replacements.Count) replacements"
        return $htmlContent
    }
    catch {
        throw "Failed to process HTML template: $($_.Exception.Message)"
    }
}

# ============================================================================
# MODULE: Inline Image Handling
# ============================================================================

function New-EmailAlternateViewWithImages {
    <#
    .SYNOPSIS
        Creates AlternateView with embedded inline images
    .PARAMETER HtmlBody
        HTML content as string
    .PARAMETER InlineImages
        Array of hashtables with ContentId and FilePath
        Example: @{ ContentId = 'logo1'; FilePath = 'C:\temp\logo.png' }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter(Mandatory = $false)]
        [hashtable[]]$InlineImages = @()
    )

    try {
        # Create AlternateView for HTML body
        $altView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString(
            $HtmlBody,
            [System.Text.Encoding]::UTF8,
            "text/html"
        )

        # Add inline images as LinkedResources
        foreach ($image in $InlineImages) {
            if (-not (Test-Path $image.FilePath)) {
                Write-Warning "Inline image not found: $($image.FilePath)"
                continue
            }

            $linkedResource = New-Object System.Net.Mail.LinkedResource($image.FilePath)
            
            # Determine MIME type from extension
            $extension = [System.IO.Path]::GetExtension($image.FilePath).TrimStart('.').ToLower()
            if ($extension -eq 'jpg') { $extension = 'jpeg' }
            
            $linkedResource.ContentType = New-Object System.Net.Mime.ContentType("image/$extension")
            $linkedResource.ContentId = $image.ContentId
            $linkedResource.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64

            [void]$altView.LinkedResources.Add($linkedResource)
            Write-Verbose "Added inline image: $($image.ContentId) from $($image.FilePath)"
        }

        return $altView
    }
    catch {
        throw "Failed to create AlternateView with images: $($_.Exception.Message)"
    }
}

# ============================================================================
# MODULE: SMTP Sending
# ============================================================================

function Send-HtmlEmail {
    <#
    .SYNOPSIS
        Sends HTML email with inline images and attachments
    .PARAMETER To
        Recipient email address
    .PARAMETER From
        Sender email address
    .PARAMETER Subject
        Email subject
    .PARAMETER HtmlBody
        HTML body content
    .PARAMETER InlineImages
        Array of hashtables with inline image definitions
    .PARAMETER Attachments
        Array of file paths to attach
    .PARAMETER Bcc
        BCC recipient(s)
    .PARAMETER SmtpServer
        SMTP server address
    .PARAMETER SmtpPort
        SMTP server port
    .PARAMETER Credential
        PSCredential object for authentication
    .PARAMETER EnableSsl
        Enable SSL/TLS (default: $true)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$To,

        [Parameter(Mandatory = $true)]
        [string]$From,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter(Mandatory = $false)]
        [hashtable[]]$InlineImages = @(),

        [Parameter(Mandatory = $false)]
        [string[]]$Attachments = @(),

        [Parameter(Mandatory = $false)]
        [string]$Bcc,

        [Parameter(Mandatory = $true)]
        [string]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [int]$SmtpPort,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory = $false)]
        [bool]$EnableSsl = $true
    )

    $mailMessage = $null
    $smtpClient = $null

    try {
        # Create mail message
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $From
        $mailMessage.To.Add($To)
        if ($Bcc) { $mailMessage.Bcc.Add($Bcc) }
        $mailMessage.Subject = $Subject
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.IsBodyHtml = $true

        # Create and add AlternateView with inline images
        $altView = New-EmailAlternateViewWithImages -HtmlBody $HtmlBody -InlineImages $InlineImages
        [void]$mailMessage.AlternateViews.Add($altView)

        # Add attachments
        foreach ($attachmentPath in $Attachments) {
            if (Test-Path $attachmentPath) {
                $attachment = New-Object System.Net.Mail.Attachment($attachmentPath)
                [void]$mailMessage.Attachments.Add($attachment)
                Write-Verbose "Added attachment: $attachmentPath"
            }
            else {
                Write-Warning "Attachment not found, skipping: $attachmentPath"
            }
        }

        # Create SMTP client
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtpClient.EnableSsl = $EnableSsl
        $smtpClient.Credentials = $Credential

        # Send email
        Write-Verbose "Sending email to: $To"
        $smtpClient.Send($mailMessage)
        Write-Verbose "Email sent successfully to: $To"

        return $true
    }
    catch {
        Write-Warning "Failed to send email to ${To}: $($_.Exception.Message)"
        return $false
    }
    finally {
        # Cleanup
        if ($mailMessage) { $mailMessage.Dispose() }
        if ($smtpClient) { $smtpClient.Dispose() }
    }
}

# ============================================================================
# MODULE: Batch Sending with Checkpointing
# ============================================================================

function Send-BulkHtmlEmail {
    <#
    .SYNOPSIS
        Sends bulk HTML emails with batching, rate limiting, and checkpointing
    .PARAMETER Recipients
        Array of recipient objects with email and replacement data
    .PARAMETER TemplateConfig
        Hashtable with template configuration
    .PARAMETER SmtpConfig
        Hashtable with SMTP configuration
    .PARAMETER BatchSize
        Number of emails per batch window
    .PARAMETER WindowMinutes
        Minutes between batch windows
    .PARAMETER MaxRetries
        Maximum retry attempts per email (default: 3)
    .PARAMETER CheckpointPath
        Path to checkpoint file for resume capability
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Recipients,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateConfig,

        [Parameter(Mandatory = $true)]
        [hashtable]$SmtpConfig,

        [Parameter(Mandatory = $false)]
        [int]$BatchSize = 20,

        [Parameter(Mandatory = $false)]
        [double]$WindowMinutes = 3.0,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [string]$CheckpointPath = "C:\temp\email_checkpoint.txt"
    )

    # Load checkpoint
    $sentEmails = New-Object System.Collections.Generic.HashSet[string]
    if (Test-Path $CheckpointPath) {
        Get-Content $CheckpointPath | ForEach-Object {
            [void]$sentEmails.Add($_.Trim())
        }
        Write-Host "Loaded checkpoint: $($sentEmails.Count) emails already sent"
    }

    # Filter recipients
    $pendingRecipients = $Recipients | Where-Object {
        $email = $_.Email.Trim()
        -not [string]::IsNullOrWhiteSpace($email) -and -not $sentEmails.Contains($email)
    }

    if ($pendingRecipients.Count -eq 0) {
        Write-Host "No pending emails to send (all recipients already processed or list empty)"
        return
    }

    Write-Host "`n=== BULK EMAIL CAMPAIGN ===" -ForegroundColor Cyan
    Write-Host "Total pending: $($pendingRecipients.Count)"
    Write-Host "Batch size: $BatchSize"
    Write-Host "Window interval: $WindowMinutes minutes"
    Write-Host "Max retries: $MaxRetries"
    Write-Host "===========================`n" -ForegroundColor Cyan

    # Load base HTML template
    $baseHtml = Get-Content -LiteralPath $TemplateConfig.TemplatePath -Raw -Encoding $TemplateConfig.Encoding

    # Stopwatch for timing
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Process in batches
    for ($offset = 0; $offset -lt $pendingRecipients.Count; $offset += $BatchSize) {
        $windowStart = $stopwatch.Elapsed
        $endIndex = [Math]::Min($offset + $BatchSize - 1, $pendingRecipients.Count - 1)
        $batch = $pendingRecipients[$offset..$endIndex]

        Write-Host ("[{0:HH:mm:ss}] === Batch {1}-{2} of {3} ===" -f (Get-Date), ($offset + 1), ($endIndex + 1), $pendingRecipients.Count) -ForegroundColor Yellow

        foreach ($recipient in $batch) {
            $email = $recipient.Email.Trim()
            
            # Process template with recipient-specific replacements
            $htmlBody = $baseHtml
            foreach ($key in $recipient.Replacements.Keys) {
                $value = [System.Web.HttpUtility]::HtmlEncode($recipient.Replacements[$key])
                $htmlBody = $htmlBody.Replace($key, $value)
            }

            # Retry logic
            $attempt = 0
            $success = $false
            $retryDelay = 2

            while ($attempt -lt $MaxRetries -and -not $success) {
                $attempt++
                
                $sendParams = @{
                    To          = $email
                    From        = $SmtpConfig.From
                    Subject     = $SmtpConfig.Subject
                    HtmlBody    = $htmlBody
                    InlineImages = $TemplateConfig.InlineImages
                    Attachments = $TemplateConfig.Attachments
                    Bcc         = $SmtpConfig.Bcc
                    SmtpServer  = $SmtpConfig.Server
                    SmtpPort    = $SmtpConfig.Port
                    Credential  = $SmtpConfig.Credential
                    EnableSsl   = $SmtpConfig.EnableSsl
                }

                $success = Send-HtmlEmail @sendParams

                if ($success) {
                    # Checkpoint immediately
                    Add-Content -LiteralPath $CheckpointPath -Value $email
                    Write-Host "  ✓ Sent: $email" -ForegroundColor Green
                }
                else {
                    if ($attempt -lt $MaxRetries) {
                        Write-Warning "  ⚠ Retry $attempt/$MaxRetries for $email (waiting ${retryDelay}s)"
                        Start-Sleep -Seconds $retryDelay
                        $retryDelay = [Math]::Min($retryDelay * 2, 30)
                    }
                    else {
                        Write-Warning "  ✗ Permanent failure after $MaxRetries attempts: $email"
                    }
                }
            }

            # Small jitter between emails
            Start-Sleep -Milliseconds (Get-Random -Minimum 150 -Maximum 500)
        }

        # Enforce window spacing
        $elapsed = $stopwatch.Elapsed - $windowStart
        $targetWindow = [TimeSpan]::FromMinutes($WindowMinutes)
        
        if ($elapsed -lt $targetWindow) {
            $sleepSeconds = [int](($targetWindow - $elapsed).TotalSeconds)
            if ($sleepSeconds -gt 0) {
                Write-Host "`n⏱ Waiting ${sleepSeconds}s to honor batch window spacing..." -ForegroundColor Cyan
                Start-Sleep -Seconds $sleepSeconds
            }
        }
        Write-Host ""
    }

    $stopwatch.Stop()
    Write-Host "`n=== CAMPAIGN COMPLETE ===" -ForegroundColor Green
    Write-Host "Total time: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
    Write-Host "Emails sent: $($pendingRecipients.Count)"
    Write-Host "========================`n" -ForegroundColor Green
}

# ============================================================================
# EXPORTS
# ============================================================================

Export-ModuleMember -Function @(
    'Get-EmailAuthenticationCredential',
    'Get-ProcessedHtmlTemplate',
    'New-EmailAlternateViewWithImages',
    'Send-HtmlEmail',
    'Send-BulkHtmlEmail'
)
