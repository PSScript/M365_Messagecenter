# M365 Digest Email System

Professional, modular PowerShell email system for sending HTML emails with inline images, attachments, and flexible authentication (Basic Auth / OAuth2).

## 🎯 Features

- **Modular Architecture**: Separate modules for authentication, template processing, and sending
- **Dual Authentication**: Support for both Basic Auth and OAuth2 (Client Credentials Flow)
- **Inline Images**: Embedded images using AlternateView and LinkedResources
- **Batch Processing**: Rate-limited sending with configurable batch windows
- **Resume Capability**: Checkpoint-based system to resume interrupted campaigns
- **Retry Logic**: Automatic retry with exponential backoff
- **Template System**: HTML templates with placeholder replacement
- **Security**: HTML encoding to prevent injection attacks
- **Attachments**: Support for multiple PDF/file attachments

## 📁 File Structure

```
M365DigestEmailModule.psm1         # Core PowerShell module
M365_Digest_Template.htm           # HTML email template
Send-M365Digest-BasicAuth.ps1     # Example: Basic Authentication
Send-M365Digest-OAuth.ps1          # Example: OAuth2 Authentication
README.md                          # This file
```

## 🚀 Quick Start

### Prerequisites

- PowerShell 5.1 or higher
- .NET Framework 4.5+
- SMTP access to Office 365 (smtp.office365.com:587)
- For OAuth: Azure AD App Registration with Mail.Send permissions

### Basic Authentication Example

```powershell
# 1. Import the module
Import-Module .\M365DigestEmailModule.psm1

# 2. Configure and run
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Test
```

### OAuth2 Authentication Example

```powershell
# 1. Set up Azure AD App Registration (see OAuth Setup below)
# 2. Update OAuth configuration in Send-M365Digest-OAuth.ps1
# 3. Run the script
.\Send-M365Digest-OAuth.ps1 -ConfigMode Production
```

## 🔐 OAuth Setup (Azure AD)

### Step 1: Register Application

1. Navigate to **Azure Portal** → **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Name: `M365-Digest-Email-Sender`
4. Supported account types: **Accounts in this organizational directory only**
5. Click **Register**

### Step 2: Configure API Permissions

1. Go to **API permissions** blade
2. Click **Add a permission**
3. Select **Microsoft Graph** → **Application permissions**
4. Add: `Mail.Send`
5. Click **Grant admin consent for [Your Organization]**

### Step 3: Create Client Secret

1. Go to **Certificates & secrets** blade
2. Click **New client secret**
3. Description: `M365 Digest SMTP`
4. Expires: Select appropriate expiration (e.g., 24 months)
5. Click **Add**
6. **IMPORTANT**: Copy the secret value immediately (can't retrieve later)

### Step 4: Gather Configuration Values

From the **Overview** blade, note:
- **Application (client) ID**: `12345678-1234-1234-1234-123456789abc`
- **Directory (tenant) ID**: `87654321-4321-4321-4321-cba987654321`
- **Client Secret**: From Step 3

### Step 5: Update Script

Edit `Send-M365Digest-OAuth.ps1`:

```powershell
$oauthConfig = @{
    TenantId     = "87654321-4321-4321-4321-cba987654321"
    ClientId     = "12345678-1234-1234-1234-123456789abc"
    ClientSecret = "your_client_secret_here"
    Username     = "sender@yourdomain.com"
}
```

## 📧 HTML Template Customization

The template uses simple placeholder replacement:

```html
<!-- In M365_Digest_Template.htm -->
<div>CARD1_TITLE</div>
<div>CARD1_CONTENT</div>
<a href="CARD1_LINK">Link</a>

<!-- Inline images use CID references -->
<img src="cid:m365_icon" alt="M365" width="100">
```

### Placeholders

Configure in your script:

```powershell
$replacements = @{
    'CARD1_TITLE'   = "New Feature Announcement"
    'CARD1_CONTENT' = "Detailed description here..."
    'CARD1_LINK'    = "https://admin.microsoft.com/..."
    'CARD2_TITLE'   = "Another Update"
    # ... etc
}
```

## 🖼️ Inline Images Setup

Define images in your script:

```powershell
$inlineImages = @(
    @{
        ContentId = 'datagroup_logo'    # Must match cid: in template
        FilePath  = 'C:\temp\logo.png'
    },
    @{
        ContentId = 'm365_icon'
        FilePath  = 'C:\temp\m365.png'
    }
)
```

In HTML template:

```html
<img src="cid:datagroup_logo" alt="Logo" width="130">
<img src="cid:m365_icon" alt="M365" width="100">
```

## 📊 CSV Data Format

Expected CSV structure (semicolon-separated):

```csv
email;DisplayName_email;password;secret_link
user1@example.com;John Doe;Pass123;https://link1
user2@example.com;Jane Smith;Pass456;https://link2
```

Customize column mapping in the script:

```powershell
$csvData = Import-Csv -LiteralPath $csvPath -Delimiter ';' -Encoding UTF8

foreach ($row in $csvData) {
    $email = $row.email
    $displayName = $row.DisplayName_email
    # Build replacements based on your CSV columns
}
```

## ⚙️ Configuration Options

### Batch Control

```powershell
$batchConfig = @{
    BatchSize     = 20      # Emails per batch window
    WindowMinutes = 3.0     # Minutes between batches
    MaxRetries    = 3       # Retry attempts per email
}
```

### SMTP Settings

```powershell
$smtpConfig = @{
    Server    = "smtp.office365.com"
    Port      = 587
    EnableSsl = $true
    From      = "sender@example.com"
    Bcc       = "admin@example.com"    # Optional
    Subject   = "Your Email Subject"
}
```

## 🔄 Resume Capability

The system uses checkpoint files to track sent emails:

```powershell
$checkpointFile = "C:\Temp\email_checkpoint.txt"
```

If sending is interrupted:
1. The checkpoint file contains all successfully sent emails
2. Re-run the script - it will skip already-sent addresses
3. Sending resumes from where it left off

To restart from scratch: Delete the checkpoint file

## 🧪 Testing Mode

Use Test mode for initial testing:

```powershell
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Test
```

Test mode changes:
- BatchSize = 2 (sends only 2 emails per batch)
- WindowMinutes = 0.1 (6 seconds between batches)
- Faster execution for validation

## 🛡️ Security Best Practices

### Credential Storage

**Never hardcode passwords in production scripts!**

#### Option 1: Secure String Export

```powershell
# Save credential once
$cred = Get-Credential
$cred.Password | ConvertFrom-SecureString | Set-Content "C:\secure\smtp.txt"

# Load in script
$securePassword = Get-Content "C:\secure\smtp.txt" | ConvertTo-SecureString
```

#### Option 2: Azure Key Vault

```powershell
# Requires Az.KeyVault module
$secret = Get-AzKeyVaultSecret -VaultName "MyVault" -Name "SMTPPassword"
$password = $secret.SecretValueText
```

#### Option 3: Windows Credential Manager

```powershell
# Requires CredentialManager module
$cred = Get-StoredCredential -Target "M365-SMTP"
```

### OAuth Token Security

- Store Client Secrets in Azure Key Vault
- Use Managed Identities when running from Azure
- Rotate secrets every 6-12 months
- Apply principle of least privilege (Mail.Send only)

## 📈 Monitoring & Logging

The module provides colored console output:

```
✓ Green  = Success
⚠ Yellow = Warning  
✗ Red    = Error
```

### Enhanced Logging

Add file logging:

```powershell
# Add to your script
$logFile = "C:\Logs\email-campaign-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
Start-Transcript -Path $logFile
# ... run campaign ...
Stop-Transcript
```

## 🔧 Troubleshooting

### Common Issues

#### Authentication Failures (Basic Auth)

```
Error: 5.7.57 SMTP; Client was not authenticated
```

**Solution**: 
- Verify username/password
- Enable "SMTP AUTH" in Exchange admin center
- Check if Modern Auth is required

#### OAuth Token Errors

```
Error: Failed to acquire OAuth token
```

**Solution**:
- Verify TenantId, ClientId, ClientSecret
- Ensure Mail.Send permission is granted
- Confirm admin consent is completed
- Check token endpoint: `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`

#### Inline Images Not Displaying

**Solution**:
- Verify CID in template matches ContentId in config
- Use exact CID format: `<img src="cid:image_name">`
- Check file paths exist
- Ensure image files are valid (PNG, JPG)

#### Rate Limiting / Throttling

```
Error: 4.3.2 Service too busy
```

**Solution**:
- Increase WindowMinutes (try 5-10 minutes)
- Reduce BatchSize (try 10-15 emails)
- Add random jitter between sends

## 📖 Module Function Reference

### Get-EmailAuthenticationCredential

Creates credential object for SMTP authentication.

```powershell
# Basic Auth
$cred = Get-EmailAuthenticationCredential `
    -AuthMethod 'Basic' `
    -Username 'user@example.com' `
    -Password 'SecurePass123'

# OAuth2
$cred = Get-EmailAuthenticationCredential `
    -AuthMethod 'OAuth' `
    -Username 'user@example.com' `
    -TenantId '...' `
    -ClientId '...' `
    -ClientSecret '...'
```

### Get-ProcessedHtmlTemplate

Loads HTML template and replaces placeholders.

```powershell
$html = Get-ProcessedHtmlTemplate `
    -TemplatePath 'C:\temp\template.htm' `
    -Replacements @{
        'TITLE' = 'Hello World'
        'CONTENT' = 'Email body text'
    }
```

### New-EmailAlternateViewWithImages

Creates AlternateView with embedded inline images.

```powershell
$altView = New-EmailAlternateViewWithImages `
    -HtmlBody $htmlContent `
    -InlineImages @(
        @{ ContentId = 'logo'; FilePath = 'C:\logo.png' }
    )
```

### Send-HtmlEmail

Sends single HTML email with inline images and attachments.

```powershell
Send-HtmlEmail `
    -To 'recipient@example.com' `
    -From 'sender@example.com' `
    -Subject 'Test Email' `
    -HtmlBody $html `
    -InlineImages $images `
    -Attachments @('C:\file.pdf') `
    -SmtpServer 'smtp.office365.com' `
    -SmtpPort 587 `
    -Credential $cred
```

### Send-BulkHtmlEmail

Sends bulk emails with batching, rate limiting, and checkpointing.

```powershell
Send-BulkHtmlEmail `
    -Recipients $recipients `
    -TemplateConfig $templateConfig `
    -SmtpConfig $smtpConfig `
    -BatchSize 20 `
    -WindowMinutes 3.0 `
    -CheckpointPath 'C:\checkpoint.txt'
```

## 🎨 Customization Examples

### Dynamic Content Per Recipient

```powershell
foreach ($row in $csvData) {
    # Personalize based on user data
    $replacements = @{
        'GREETING' = "Hello $($row.FirstName),"
        'CONTENT' = "Your account expires on $($row.ExpiryDate)"
    }
    
    $recipients += [PSCustomObject]@{
        Email = $row.email
        Replacements = $replacements
    }
}
```

### Conditional Card Content

```powershell
# Show different content based on user type
if ($row.UserType -eq 'Premium') {
    $replacements['CARD3_TITLE'] = 'Premium Features'
    $replacements['CARD3_CONTENT'] = 'Exclusive content...'
} else {
    $replacements['CARD3_TITLE'] = 'Upgrade Today'
    $replacements['CARD3_CONTENT'] = 'Get premium features...'
}
```

## 📝 License

This code is provided as-is for use within your organization. Adapt and modify as needed.

## 🤝 Support

For issues or questions:
1. Check troubleshooting section
2. Review Office 365 SMTP documentation
3. Verify Azure AD app permissions (for OAuth)
4. Test with a single recipient first

## 🔄 Version History

- **v1.0.0** (2025-11-07): Initial release
  - Basic and OAuth authentication
  - Inline images with AlternateView
  - Batch processing with checkpoints
  - Modular architecture

---

**Built with precision for Microsoft 365 enterprise email campaigns.**
