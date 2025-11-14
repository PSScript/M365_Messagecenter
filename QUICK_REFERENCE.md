# M365 Digest Email System - Quick Reference

## 📋 Quick Start Checklist

```
□ Import module: Import-Module .\M365DigestEmailModule.psm1
□ Update HTML template with your content
□ Prepare inline images (PNG/JPG) matching CID references
□ Create CSV file with recipient data
□ Configure authentication (Basic or OAuth)
□ Test configuration: .\Test-M365DigestConfig.ps1
□ Run test campaign: .\Send-M365Digest-BasicAuth.ps1 -ConfigMode Test
□ Run production: .\Send-M365Digest-BasicAuth.ps1 -ConfigMode Production
```

## 🎯 Common Commands

### Test Configuration
```powershell
# Full validation
.\Test-M365DigestConfig.ps1

# Skip auth test
.\Test-M365DigestConfig.ps1 -SkipAuthTest
```

### Send Emails (Basic Auth)
```powershell
# Test mode (2 emails, fast)
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Test

# Production mode
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Production
```

### Send Emails (OAuth)
```powershell
# Test mode
.\Send-M365Digest-OAuth.ps1 -ConfigMode Test

# Production mode
.\Send-M365Digest-OAuth.ps1 -ConfigMode Production
```

### Import Module Manually
```powershell
Import-Module .\M365DigestEmailModule.psm1 -Force -Verbose
```

## 🔑 Template Placeholders

### In HTML Template
```html
<!-- Card placeholders -->
CARD1_TITLE
CARD1_CONTENT
CARD1_LINK

CARD2_TITLE
CARD2_CONTENT
CARD2_LINK

CARD3_TITLE
CARD3_CONTENT
CARD3_LINK

<!-- Other placeholders -->
UNSUBSCRIBE_LINK
DISPLAYNAME
```

### In Script Replacements
```powershell
$replacements = @{
    'CARD1_TITLE'   = 'Your Title Here'
    'CARD1_CONTENT' = 'Your content here'
    'CARD1_LINK'    = 'https://example.com'
    # ... repeat for CARD2, CARD3
}
```

## 🖼️ Inline Images Configuration

### Define in Script
```powershell
$inlineImages = @(
    @{
        ContentId = 'logo1'                    # Unique CID
        FilePath  = 'C:\temp\logo.png'         # Actual file path
    },
    @{
        ContentId = 'm365_icon'
        FilePath  = 'C:\temp\m365.png'
    }
)
```

### Reference in HTML
```html
<img src="cid:logo1" alt="Logo" width="130">
<img src="cid:m365_icon" alt="M365" width="100">
```

## 📊 CSV File Format

### Required Columns
```csv
email;DisplayName_email;password;secret_link
john@example.com;John Doe;Pass123;https://link1
jane@example.com;Jane Smith;Pass456;https://link2
```

### Import in Script
```powershell
$csvData = Import-Csv -Path $csvPath -Delimiter ';' -Encoding UTF8

foreach ($row in $csvData) {
    $email = $row.email
    $name = $row.DisplayName_email
    # Process row...
}
```

## ⚙️ Batch Configuration

### Standard (Production)
```powershell
$batchConfig = @{
    BatchSize     = 20      # 20 emails per window
    WindowMinutes = 3.0     # 3 minute windows
    MaxRetries    = 3       # Retry 3 times
}
```

### Fast (Testing)
```powershell
$batchConfig = @{
    BatchSize     = 2       # 2 emails per window
    WindowMinutes = 0.1     # 6 second windows
    MaxRetries    = 1       # Retry once
}
```

### Slow (Conservative)
```powershell
$batchConfig = @{
    BatchSize     = 10      # 10 emails per window
    WindowMinutes = 5.0     # 5 minute windows
    MaxRetries    = 5       # Retry 5 times
}
```

## 🔐 Authentication Quick Setup

### Basic Auth
```powershell
$credential = Get-EmailAuthenticationCredential `
    -AuthMethod 'Basic' `
    -Username 'user@example.com' `
    -Password 'YourPassword'
```

### OAuth (Azure AD App)
```powershell
$credential = Get-EmailAuthenticationCredential `
    -AuthMethod 'OAuth' `
    -Username 'user@example.com' `
    -TenantId 'your-tenant-id' `
    -ClientId 'your-client-id' `
    -ClientSecret 'your-secret'
```

## 🔄 Resume Interrupted Campaign

If sending is interrupted:

1. **Don't panic** - checkpoint file tracks sent emails
2. **Rerun the script** - it automatically skips sent addresses
3. **Check checkpoint file** to see what was sent:
   ```powershell
   Get-Content C:\temp\smtp_send_checkpoint.txt
   ```
4. **Start fresh** (optional) - delete checkpoint to resend all:
   ```powershell
   Remove-Item C:\temp\smtp_send_checkpoint.txt
   ```

## 🛠️ Quick Fixes

### Images Not Showing
```powershell
# Check CIDs match
Select-String -Path .\M365_Digest_Template.htm -Pattern "cid:"

# Expected: cid:m365_icon, cid:exchange_icon, etc.
# Must match ContentId in $inlineImages array
```

### Authentication Failed (Basic)
```powershell
# Test manually
$cred = Get-Credential
Send-MailMessage -To 'test@example.com' `
    -From 'sender@example.com' `
    -Subject 'Test' -Body 'Test' `
    -SmtpServer 'smtp.office365.com' -Port 587 `
    -UseSsl -Credential $cred
```

### OAuth Token Fails
```powershell
# Test token acquisition manually
$tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://outlook.office365.com/.default"
    grant_type    = "client_credentials"
}
Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body
```

### Check Sent Count
```powershell
# Count successfully sent emails
$sent = Get-Content C:\temp\smtp_send_checkpoint.txt
Write-Host "Sent: $($sent.Count) emails"
```

## 📈 Monitoring During Send

### Watch Progress
```powershell
# In another PowerShell window while campaign runs
Get-Content C:\temp\smtp_send_checkpoint.txt -Wait
```

### Real-time Stats
```powershell
# Count sent every 10 seconds
while ($true) {
    $count = (Get-Content C:\temp\smtp_send_checkpoint.txt).Count
    Write-Host "Sent: $count emails @ $(Get-Date -Format 'HH:mm:ss')"
    Start-Sleep -Seconds 10
}
```

## 🔍 Troubleshooting Commands

### Verify Module Functions
```powershell
Get-Command -Module M365DigestEmailModule
```

### Check Image Files
```powershell
$images = @('logo.png', 'm365.png', 'exchange.png')
$images | ForEach-Object {
    $path = "C:\temp\$_"
    if (Test-Path $path) {
        $size = (Get-Item $path).Length / 1KB
        Write-Host "✓ $_ exists (${size} KB)"
    } else {
        Write-Host "✗ $_ NOT FOUND"
    }
}
```

### Test HTML Template
```powershell
$html = Get-Content .\M365_Digest_Template.htm -Raw
$html -match 'cid:' | Out-Null
$Matches
```

### Verify CSV Format
```powershell
$csv = Import-Csv .\recipients.csv -Delimiter ';' -Encoding UTF8
$csv | Select-Object -First 3 | Format-Table
$csv | Measure-Object | Select-Object Count
```

## 🎨 Customization Snippets

### Add Custom Placeholder
```powershell
# In HTML
<div>CUSTOM_FIELD</div>

# In script
$replacements['CUSTOM_FIELD'] = $row.CustomColumn
```

### Conditional Content
```powershell
if ($row.UserType -eq 'Premium') {
    $replacements['CARD3_TITLE'] = 'Premium Features'
} else {
    $replacements['CARD3_TITLE'] = 'Upgrade Available'
}
```

### Dynamic Subject Line
```powershell
$smtpConfig.Subject = "Digest for $($row.DisplayName_email)"
```

## 🎯 Production Checklist

Before running production campaign:

```
□ Test configuration validated (green)
□ Test email sent successfully
□ HTML renders correctly in Outlook/Gmail
□ Images display correctly
□ All links functional
□ Unsubscribe link works
□ Batch size appropriate for volume
□ Authentication credentials secured
□ Checkpoint file path configured
□ BCC address set for monitoring
□ CSV file has valid email addresses
□ Backup checkpoint file: Copy-Item checkpoint.txt checkpoint.bak
```

## 📞 Emergency Stop

If you need to stop a running campaign:

1. **Press Ctrl+C** in PowerShell window
2. **Checkpoint is safe** - already-sent emails are tracked
3. **Resume later** by rerunning the same script

## 💡 Pro Tips

- **Test first**: Always run with `-ConfigMode Test` initially
- **Monitor logs**: Keep PowerShell window visible during send
- **Check inbox**: Send test email to yourself first
- **Verify images**: Open test email in multiple clients (Outlook, Gmail, Apple Mail)
- **Backup checkpoint**: Copy checkpoint file before production runs
- **Stagger sending**: Use WindowMinutes=5 for large campaigns (>1000 emails)

---

**Need more help?** See full [README.md](README.md) for detailed documentation.
