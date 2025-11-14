#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Digest Email Campaign - Basic Authentication Example
.DESCRIPTION
    Example script demonstrating bulk email sending with:
    - Basic SMTP authentication
    - CSV data import
    - Template-based HTML emails
    - Inline images (logo + 3 product icons)
    - PDF attachments
    - Batched sending with rate limiting
.NOTES
    Requires: M365DigestEmailModule.psm1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigMode = 'Production'  # 'Test' or 'Production'
)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Import module
$modulePath = Join-Path $PSScriptRoot "M365DigestEmailModule.psm1"
Import-Module $modulePath -Force -Verbose

# Paths
$csvPath = "C:\Temp\master_users_all_merged_with_wave4.csv"
$htmlTemplate = "C:\Temp\M365_Digest_Template.htm"
$checkpointFile = "C:\Temp\smtp_send_checkpoint_m365digest.txt"

# Inline Images (CID must match template references)
$inlineImages = @(
    @{
        ContentId = 'datagroup_logo'
        FilePath  = 'C:\temp\datagroup_logo.png'
    },
    @{
        ContentId = 'm365_icon'
        FilePath  = 'C:\temp\m365_icon.png'
    },
    @{
        ContentId = 'exchange_icon'
        FilePath  = 'C:\temp\exchange_icon.png'
    },
    @{
        ContentId = 'sharepoint_icon'
        FilePath  = 'C:\temp\sharepoint_icon.png'
    }
)

# Attachments
$attachments = @(
    "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Smartphone.pdf",
    "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Telefon.pdf",
    "C:\temp\Anleitung_Postfach_und_Dateiablage_auf_dem_Computer_einrichten.pdf",
    "C:\temp\Anleitung_Einrichtung-Thunderbird.pdf"
)

# SMTP Configuration
$smtpConfig = @{
    Server    = "smtp.office365.com"
    Port      = 587
    EnableSsl = $true
    From      = "OKR.noreply@ELK-WUE.DE"
    Bcc       = "jan.huebener@elkw.de"
    Subject   = "Microsoft 365 Monthly Digest – What's new?"
}

# Authentication (Basic)
$smtpUsername = "adm_huebener@elkw.de"
$smtpPassword = "YourSecurePassword123!"  # TODO: Replace with secure storage

# Batch Configuration
$batchConfig = @{
    BatchSize     = 20
    WindowMinutes = 3.0
    MaxRetries    = 3
}

# Test Mode Configuration
if ($ConfigMode -eq 'Test') {
    Write-Host "`n⚠ RUNNING IN TEST MODE" -ForegroundColor Yellow
    $batchConfig.BatchSize = 2
    $batchConfig.WindowMinutes = 0.1
    $smtpConfig.Bcc = "jan.huebener@elkw.de"
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

try {
    Write-Host "`n╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  M365 Monthly Digest Email Campaign (Basic Auth)         ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

    # Validate files exist
    Write-Host "Validating configuration..." -ForegroundColor Cyan
    
    if (-not (Test-Path $csvPath)) {
        throw "CSV file not found: $csvPath"
    }
    if (-not (Test-Path $htmlTemplate)) {
        throw "HTML template not found: $htmlTemplate"
    }
    
    foreach ($img in $inlineImages) {
        if (-not (Test-Path $img.FilePath)) {
            Write-Warning "Inline image not found: $($img.FilePath) (CID: $($img.ContentId))"
        }
    }
    
    foreach ($attachment in $attachments) {
        if (-not (Test-Path $attachment)) {
            Write-Warning "Attachment not found: $attachment"
        }
    }

    Write-Host "✓ Configuration validated`n" -ForegroundColor Green

    # Get authentication credential
    Write-Host "Authenticating..." -ForegroundColor Cyan
    $credential = Get-EmailAuthenticationCredential `
        -AuthMethod 'Basic' `
        -Username $smtpUsername `
        -Password $smtpPassword
    
    $smtpConfig.Credential = $credential
    Write-Host "✓ Authentication configured`n" -ForegroundColor Green

    # Load and process CSV
    Write-Host "Loading recipient data..." -ForegroundColor Cyan
    $csvData = Import-Csv -LiteralPath $csvPath -Delimiter ';' -Encoding UTF8

    # Build recipient objects
    $recipients = @()
    foreach ($row in $csvData) {
        $email = ($row.email).Trim()
        if ([string]::IsNullOrWhiteSpace($email)) { continue }

        # Build replacement hashtable for this recipient
        $replacements = @{
            'CARD1_TITLE'   = "New Teams Features"
            'CARD1_CONTENT' = "Microsoft Teams introduces new collaboration features including enhanced meeting recordings and AI-powered meeting summaries."
            'CARD1_LINK'    = "https://admin.microsoft.com/?ref=MessageCenter/:/messages/MC1069560"
            
            'CARD2_TITLE'   = "Exchange Online Updates"
            'CARD2_CONTENT' = "Enhanced security features now available for Exchange Online mailboxes, including improved phishing protection."
            'CARD2_LINK'    = "https://admin.microsoft.com/?ref=MessageCenter/:/messages/MC1134178"
            
            'CARD3_TITLE'   = "SharePoint Improvements"
            'CARD3_CONTENT' = "New document management capabilities in SharePoint Online with AI-powered search and classification."
            'CARD3_LINK'    = "https://admin.microsoft.com/?ref=MessageCenter/:/messages/MC1069560"
            
            'UNSUBSCRIBE_LINK' = "https://www.datagroup.de/unsubscribe?email=$email"
        }

        # You can personalize per user if CSV has columns like DisplayName
        if ($row.PSObject.Properties.Name -contains 'DisplayName_email' -and $row.DisplayName_email) {
            $replacements['CARD1_CONTENT'] = "Hello $($row.DisplayName_email), " + $replacements['CARD1_CONTENT']
        }

        $recipients += [PSCustomObject]@{
            Email        = $email
            Replacements = $replacements
        }
    }

    Write-Host "✓ Loaded $($recipients.Count) recipients`n" -ForegroundColor Green

    # Template Configuration
    $templateConfig = @{
        TemplatePath = $htmlTemplate
        Encoding     = 'UTF8'
        InlineImages = $inlineImages
        Attachments  = $attachments
    }

    # Send bulk emails
    $bulkParams = @{
        Recipients      = $recipients
        TemplateConfig  = $templateConfig
        SmtpConfig      = $smtpConfig
        BatchSize       = $batchConfig.BatchSize
        WindowMinutes   = $batchConfig.WindowMinutes
        MaxRetries      = $batchConfig.MaxRetries
        CheckpointPath  = $checkpointFile
    }

    Send-BulkHtmlEmail @bulkParams

    Write-Host "`n✓ Campaign completed successfully!" -ForegroundColor Green

}
catch {
    Write-Host "`n✗ ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
