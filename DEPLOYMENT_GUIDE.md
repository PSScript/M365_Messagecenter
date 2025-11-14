# M365 Digest Email System - Deployment Guide

## 🚀 Production Deployment Checklist

Follow this guide for first-time deployment and production campaigns.

---

## Phase 1: Initial Setup (30 minutes)

### Step 1: Prepare Your Environment

```powershell
# Create working directory
New-Item -Path "C:\M365Digest" -ItemType Directory -Force
Set-Location "C:\M365Digest"

# Create subdirectories
New-Item -Path ".\Images" -ItemType Directory -Force
New-Item -Path ".\Attachments" -ItemType Directory -Force
New-Item -Path ".\Data" -ItemType Directory -Force
New-Item -Path ".\Logs" -ItemType Directory -Force
```

### Step 2: Copy System Files

Copy these files to `C:\M365Digest`:

```
✓ M365DigestEmailModule.psm1       (Required)
✓ M365_Digest_Template.htm         (Required)
✓ Send-M365Digest-BasicAuth.ps1    (Choose one)
✓ Send-M365Digest-OAuth.ps1        (Choose one)
✓ Test-M365DigestConfig.ps1        (Recommended)
✓ README.md                         (Reference)
✓ QUICK_REFERENCE.md                (Reference)
```

### Step 3: Prepare Your Assets

#### Inline Images (to `.\Images\`)
```
Required CID references (match template):
✓ datagroup_logo.png     (Logo, ~130px wide)
✓ m365_icon.png          (Icon, ~100px square)
✓ exchange_icon.png      (Icon, ~100px square)
✓ sharepoint_icon.png    (Icon, ~100px square)
```

**Image Requirements:**
- Format: PNG or JPG
- Max size: 500KB per image (smaller is better for email clients)
- Dimensions: As specified in template
- Optimization: Use tools like TinyPNG to compress

#### PDF Attachments (to `.\Attachments\`)
```
Optional attachments:
□ Manual1.pdf
□ Manual2.pdf
□ Manual3.pdf
```

**Attachment Guidelines:**
- Max total size: 10MB per email
- Keep PDFs optimized (<2MB each)
- Test with attachments first

---

## Phase 2: Configuration (20 minutes)

### Step 1: Choose Authentication Method

#### Option A: Basic Authentication (Simpler)

**Requirements:**
- SMTP enabled account in Office 365
- Modern Authentication or SMTP AUTH enabled

**Configuration:**
Edit `Send-M365Digest-BasicAuth.ps1`:

```powershell
# Lines 40-43
$smtpUsername = "your-smtp-account@yourdomain.com"
$smtpPassword = "YourSecurePassword"  # Use secure storage in production!
```

**Security Note:** Never commit passwords to source control!

Consider using:
```powershell
# Secure string from file
$securePassword = Get-Content "C:\Secure\smtp.key" | ConvertTo-SecureString
$smtpPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
)
```

#### Option B: OAuth2 Authentication (Recommended for Enterprise)

**Requirements:**
- Azure AD App Registration
- Mail.Send permission
- Admin consent granted

**Setup Steps:**

1. **Register Azure AD Application**
   ```
   Portal: https://portal.azure.com
   Path: Azure Active Directory > App registrations > New registration
   Name: M365-Digest-Email-Sender
   Type: Single tenant
   ```

2. **Configure API Permissions**
   ```
   Path: API permissions > Add permission
   API: Microsoft Graph
   Type: Application permissions
   Permission: Mail.Send
   Action: Grant admin consent
   ```

3. **Create Client Secret**
   ```
   Path: Certificates & secrets > New client secret
   Description: M365 Digest SMTP
   Expires: 24 months
   ⚠ Copy the secret value immediately!
   ```

4. **Update Script Configuration**

   Edit `Send-M365Digest-OAuth.ps1`:

   ```powershell
   # Lines 60-65
   $oauthConfig = @{
       TenantId     = "your-tenant-id-here"
       ClientId     = "your-client-id-here"
       ClientSecret = "your-client-secret-here"
       Username     = "sending-account@yourdomain.com"
   }
   ```

### Step 2: Update File Paths

Edit your chosen script (Basic or OAuth):

```powershell
# Lines 23-26
$csvPath = "C:\M365Digest\Data\recipients.csv"
$htmlTemplate = "C:\M365Digest\M365_Digest_Template.htm"
$checkpointFile = "C:\M365Digest\Logs\checkpoint.txt"

# Lines 29-40 - Update image paths
$inlineImages = @(
    @{
        ContentId = 'datagroup_logo'
        FilePath  = 'C:\M365Digest\Images\datagroup_logo.png'
    },
    # ... update remaining paths
)

# Lines 44-49 - Update attachment paths
$attachments = @(
    "C:\M365Digest\Attachments\Manual1.pdf"
    # ... add your attachments
)

# Lines 52-58 - Update SMTP configuration
$smtpConfig = @{
    Server    = "smtp.office365.com"
    Port      = 587
    EnableSsl = $true
    From      = "sender@yourdomain.com"
    Bcc       = "admin@yourdomain.com"  # Optional monitoring
    Subject   = "Microsoft 365 Monthly Digest"
}
```

### Step 3: Customize HTML Template

Edit `M365_Digest_Template.htm`:

1. **Update Header** (Line ~25-30)
   ```html
   <div style="color:#ffffff; font-size:40px;">Microsoft</div>
   <div style="color:#ffffff; font-size:56px;">365</div>
   <!-- Change to your branding -->
   ```

2. **Update Footer** (Line ~180-190)
   ```html
   <strong>DATAGROUP Stuttgart GmbH</strong><br>
   <!-- Update with your company information -->
   ```

3. **Verify CID References**
   ```html
   <img src="cid:datagroup_logo" ...>
   <img src="cid:m365_icon" ...>
   <img src="cid:exchange_icon" ...>
   <img src="cid:sharepoint_icon" ...>
   ```

   Must match `ContentId` values in your script's `$inlineImages` array!

### Step 4: Prepare Recipient Data

Create `C:\M365Digest\Data\recipients.csv`:

```csv
email;DisplayName_email;password;secret_link
john.doe@example.com;John Doe;TempPass123;https://portal.example.com/setup/abc
jane.smith@example.com;Jane Smith;SecurePass456;https://portal.example.com/setup/def
```

**Important:**
- Delimiter: semicolon (`;`)
- Encoding: UTF-8
- No BOM (Byte Order Mark)
- Headers required: email, DisplayName_email (or customize in script)

---

## Phase 3: Testing (15 minutes)

### Test 1: Validate Configuration

```powershell
cd C:\M365Digest
.\Test-M365DigestConfig.ps1
```

Expected output:
```
✓ Module imported successfully
✓ HTML Template exists
✓ Template processing successful
✓ Image valid: CID='datagroup_logo' (45 KB)
✓ Image valid: CID='m365_icon' (12 KB)
✓ CID reference found in template: cid:datagroup_logo
✓ Authentication configured

Results: 7 / 7 tests passed
```

### Test 2: Send Test Email

Create test CSV with YOUR email only:

```csv
email;DisplayName_email;password;secret_link
your.email@yourdomain.com;Test User;TestPass123;https://example.com/test
```

Run test campaign:

```powershell
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Test
```

### Test 3: Verify Email

Check your inbox:

```
✓ Email received
✓ Subject line correct
✓ All inline images display
✓ All 3 cards display correctly
✓ Links are clickable
✓ Attachments present (if configured)
✓ Unsubscribe link works
✓ Renders well in multiple clients:
  □ Outlook Desktop
  □ Outlook Web
  □ Gmail
  □ Mobile (iOS/Android)
```

---

## Phase 4: Production Deployment (10 minutes)

### Pre-Flight Checklist

```
Configuration:
✓ All file paths updated and verified
✓ Images exist and display correctly
✓ Authentication tested successfully
✓ HTML template customized
✓ Recipient CSV prepared and validated
✓ Test email sent and verified

Security:
✓ Passwords stored securely (not hardcoded)
✓ OAuth secrets protected (Key Vault recommended)
✓ BCC configured for monitoring
✓ Checkpoint file path writable

Performance:
✓ Batch size appropriate (10-20 for production)
✓ Window interval sufficient (3-5 minutes)
✓ Retry count reasonable (3 recommended)

Monitoring:
✓ Log directory exists and writable
✓ Checkpoint directory accessible
✓ Backup strategy in place
```

### Launch Production Campaign

1. **Backup checkpoint** (if resuming):
   ```powershell
   Copy-Item C:\M365Digest\Logs\checkpoint.txt `
             C:\M365Digest\Logs\checkpoint_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt
   ```

2. **Start transcription logging**:
   ```powershell
   $logFile = "C:\M365Digest\Logs\campaign_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
   Start-Transcript -Path $logFile
   ```

3. **Launch campaign**:
   ```powershell
   .\Send-M365Digest-BasicAuth.ps1 -ConfigMode Production
   # or
   .\Send-M365Digest-OAuth.ps1 -ConfigMode Production
   ```

4. **Monitor progress**:
   - Watch PowerShell output
   - Monitor checkpoint file growth
   - Check BCC inbox for sample emails

5. **Stop transcript** (after completion):
   ```powershell
   Stop-Transcript
   ```

---

## Phase 5: Monitoring & Troubleshooting

### Real-Time Monitoring

**Terminal Window 1: Campaign**
```powershell
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Production
```

**Terminal Window 2: Progress Monitor**
```powershell
# Watch checkpoint file
Get-Content C:\M365Digest\Logs\checkpoint.txt -Wait

# Or count sent emails every 10 seconds
while ($true) {
    $count = (Get-Content C:\M365Digest\Logs\checkpoint.txt).Count
    Write-Host "Sent: $count emails @ $(Get-Date -Format 'HH:mm:ss')"
    Start-Sleep -Seconds 10
}
```

### Common Issues & Solutions

#### Issue: "Authentication failed"

**Basic Auth:**
```powershell
# Test credentials manually
$cred = Get-Credential
Test-Connection -ComputerName smtp.office365.com -Count 1
Send-MailMessage -To "test@example.com" -From $smtpConfig.From `
    -Subject "Test" -Body "Test" -SmtpServer "smtp.office365.com" `
    -Port 587 -UseSsl -Credential $cred
```

**OAuth:**
```powershell
# Test token acquisition
$tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$body = @{
    client_id = $ClientId
    client_secret = $ClientSecret
    scope = "https://outlook.office365.com/.default"
    grant_type = "client_credentials"
}
$response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body
$response.access_token  # Should return token
```

#### Issue: "Inline images not displaying"

```powershell
# Verify CIDs in template
Select-String -Path .\M365_Digest_Template.htm -Pattern "cid:"

# Verify file paths
Get-ChildItem C:\M365Digest\Images

# Check CID matches
# Template: <img src="cid:m365_icon">
# Script: ContentId = 'm365_icon'  (must match exactly!)
```

#### Issue: "Rate limiting / throttling"

```powershell
# Slow down the campaign
$batchConfig = @{
    BatchSize     = 10       # Reduce from 20
    WindowMinutes = 5.0      # Increase from 3.0
    MaxRetries    = 3
}
```

#### Issue: "Campaign interrupted"

```powershell
# Check what was sent
Get-Content C:\M365Digest\Logs\checkpoint.txt

# Resume by rerunning script (automatic resume)
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Production

# Script will skip emails already in checkpoint file
```

---

## Phase 6: Post-Campaign

### Verify Completion

```powershell
# Count sent emails
$sent = (Get-Content C:\M365Digest\Logs\checkpoint.txt).Count
$total = (Import-Csv C:\M365Digest\Data\recipients.csv -Delimiter ';').Count

Write-Host "Campaign Statistics:"
Write-Host "  Total recipients: $total"
Write-Host "  Successfully sent: $sent"
Write-Host "  Success rate: $([math]::Round($sent/$total*100, 2))%"
```

### Archive Campaign Data

```powershell
$archiveDate = Get-Date -Format 'yyyyMMdd_HHmmss'
$archivePath = "C:\M365Digest\Archives\Campaign_$archiveDate"

New-Item -Path $archivePath -ItemType Directory -Force

# Copy campaign artifacts
Copy-Item C:\M365Digest\Logs\checkpoint.txt "$archivePath\"
Copy-Item C:\M365Digest\Logs\*.log "$archivePath\"
Copy-Item C:\M365Digest\Data\recipients.csv "$archivePath\"

# Compress archive
Compress-Archive -Path $archivePath -DestinationPath "$archivePath.zip"

# Clean up checkpoint for next campaign
Remove-Item C:\M365Digest\Logs\checkpoint.txt
```

### Generate Report

```powershell
$report = @"
M365 Digest Email Campaign Report
==================================
Date: $(Get-Date -Format 'yyyy-MM-DD HH:mm:ss')
Campaign: Microsoft 365 Monthly Digest

Statistics:
-----------
Total Recipients: $total
Successfully Sent: $sent
Success Rate: $([math]::Round($sent/$total*100, 2))%

Configuration:
--------------
Batch Size: 20 emails
Window Interval: 3 minutes
Max Retries: 3

Duration: [Calculate from logs]
"@

$report | Out-File "C:\M365Digest\Archives\Campaign_${archiveDate}\report.txt"
Write-Host $report
```

---

## Security Best Practices

### Credential Storage

**DON'T:**
```powershell
$smtpPassword = "MyPassword123"  # ✗ Hardcoded
```

**DO:**
```powershell
# Option 1: Secure string file
$securePassword = Get-Content "C:\Secure\smtp.key" | ConvertTo-SecureString

# Option 2: Windows Credential Manager
$cred = Get-StoredCredential -Target "M365-SMTP"

# Option 3: Azure Key Vault
$secret = Get-AzKeyVaultSecret -VaultName "MyVault" -Name "SMTPPassword"
```

### OAuth Token Management

```powershell
# Store in Azure Key Vault
Set-AzKeyVaultSecret -VaultName "MyVault" -Name "M365-TenantId" -SecretValue (ConvertTo-SecureString "..." -AsPlainText -Force)
Set-AzKeyVaultSecret -VaultName "MyVault" -Name "M365-ClientId" -SecretValue (ConvertTo-SecureString "..." -AsPlainText -Force)
Set-AzKeyVaultSecret -VaultName "MyVault" -Name "M365-ClientSecret" -SecretValue (ConvertTo-SecureString "..." -AsPlainText -Force)

# Retrieve in script
$tenantId = (Get-AzKeyVaultSecret -VaultName "MyVault" -Name "M365-TenantId").SecretValueText
```

### File Permissions

```powershell
# Restrict access to sensitive files
$acl = Get-Acl "C:\M365Digest\Secure"
$acl.SetAccessRuleProtection($true, $false)
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $env:USERNAME, "FullControl", "Allow"
)
$acl.SetAccessRule($rule)
Set-Acl "C:\M365Digest\Secure" $acl
```

---

## Maintenance & Updates

### Monthly Tasks

```
□ Rotate OAuth client secrets (if expiring)
□ Review and update email template
□ Update inline images if branding changes
□ Verify SMTP configuration still valid
□ Archive old campaign logs
□ Update recipient CSV
```

### Quarterly Tasks

```
□ Review security practices
□ Update PowerShell module if new version available
□ Test email rendering in latest client versions
□ Review and optimize batch settings
□ Backup entire system configuration
```

---

## Emergency Procedures

### Stop Running Campaign

1. **Press Ctrl+C** in PowerShell window
2. Checkpoint is automatically saved
3. Review what was sent: `Get-Content checkpoint.txt`

### Resume Interrupted Campaign

```powershell
# Simply rerun the script - it automatically resumes
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Production
```

### Rollback / Resend

```powershell
# Start fresh (WARNING: Will resend to all recipients)
Remove-Item C:\M365Digest\Logs\checkpoint.txt
.\Send-M365Digest-BasicAuth.ps1 -ConfigMode Production
```

---

## Support & Resources

- **Module Documentation**: README.md
- **Quick Reference**: QUICK_REFERENCE.md
- **Architecture**: ARCHITECTURE.md
- **Configuration Test**: `.\Test-M365DigestConfig.ps1`

---

**Deployment Guide v1.0.0** | M365 Digest Email System | Production Ready
