#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Digest Email System - Configuration Validator
.DESCRIPTION
    Tests and validates your email system configuration without sending emails:
    - Checks file paths
    - Validates image files
    - Tests HTML template processing
    - Verifies authentication configuration
    - Validates inline image CID references
.NOTES
    Run this before your first campaign to catch configuration issues
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$SkipAuthTest
)

# ============================================================================
# CONFIGURATION (Update these to match your setup)
# ============================================================================

$config = @{
    # Paths
    ModulePath    = ".\M365DigestEmailModule.psm1"
    HtmlTemplate  = ".\M365_Digest_Template.htm"
    TestCsvPath   = "C:\Temp\test_recipients.csv"
    
    # Images
    InlineImages  = @(
        @{ ContentId = 'datagroup_logo'; FilePath = 'C:\temp\datagroup_logo.png' }
        @{ ContentId = 'm365_icon'; FilePath = 'C:\temp\m365_icon.png' }
        @{ ContentId = 'exchange_icon'; FilePath = 'C:\temp\exchange_icon.png' }
        @{ ContentId = 'sharepoint_icon'; FilePath = 'C:\temp\sharepoint_icon.png' }
    )
    
    # Attachments
    Attachments   = @(
        "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Smartphone.pdf"
        "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Telefon.pdf"
    )
    
    # Auth Test (Basic)
    AuthType      = 'Basic'  # 'Basic' or 'OAuth'
    SmtpUser      = 'user@example.com'
    SmtpPassword  = 'password'
    
    # Auth Test (OAuth) - only if AuthType = 'OAuth'
    TenantId      = 'your-tenant-id'
    ClientId      = 'your-client-id'
    ClientSecret  = 'your-secret'
}

# ============================================================================
# TEST FUNCTIONS
# ============================================================================

function Test-FileExists {
    param([string]$Path, [string]$Description)
    
    if (Test-Path $Path) {
        Write-Host "✓ " -NoNewline -ForegroundColor Green
        Write-Host "$Description exists: " -NoNewline
        Write-Host $Path -ForegroundColor Cyan
        return $true
    }
    else {
        Write-Host "✗ " -NoNewline -ForegroundColor Red
        Write-Host "$Description NOT FOUND: " -NoNewline
        Write-Host $Path -ForegroundColor Red
        return $false
    }
}

function Test-ImageFile {
    param([hashtable]$Image)
    
    $exists = Test-Path $Image.FilePath
    if ($exists) {
        # Check if it's a valid image
        try {
            $file = Get-Item $Image.FilePath
            $validExtensions = @('.png', '.jpg', '.jpeg', '.gif')
            $ext = $file.Extension.ToLower()
            
            if ($ext -in $validExtensions) {
                Write-Host "✓ " -NoNewline -ForegroundColor Green
                Write-Host "Image valid: " -NoNewline
                Write-Host "CID='$($Image.ContentId)' " -NoNewline -ForegroundColor Yellow
                Write-Host "($($file.Length / 1KB) KB) " -NoNewline -ForegroundColor Gray
                Write-Host $Image.FilePath -ForegroundColor Cyan
                return $true
            }
            else {
                Write-Host "⚠ " -NoNewline -ForegroundColor Yellow
                Write-Host "Image has unexpected extension: $ext - $($Image.FilePath)" -ForegroundColor Yellow
                return $false
            }
        }
        catch {
            Write-Host "✗ " -NoNewline -ForegroundColor Red
            Write-Host "Cannot read image: $($Image.FilePath)" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "✗ " -NoNewline -ForegroundColor Red
        Write-Host "Image NOT FOUND: " -NoNewline
        Write-Host "CID='$($Image.ContentId)' " -NoNewline -ForegroundColor Yellow
        Write-Host $Image.FilePath -ForegroundColor Red
        return $false
    }
}

function Test-HtmlTemplateCids {
    param([string]$TemplatePath, [hashtable[]]$InlineImages)
    
    Write-Host "`nValidating CID references in HTML template..." -ForegroundColor Cyan
    $html = Get-Content $TemplatePath -Raw
    $allValid = $true
    
    foreach ($img in $InlineImages) {
        $cidPattern = "cid:$($img.ContentId)"
        if ($html -match [regex]::Escape($cidPattern)) {
            Write-Host "✓ " -NoNewline -ForegroundColor Green
            Write-Host "CID reference found in template: " -NoNewline
            Write-Host $cidPattern -ForegroundColor Yellow
        }
        else {
            Write-Host "⚠ " -NoNewline -ForegroundColor Yellow
            Write-Host "CID reference NOT found in template: " -NoNewline
            Write-Host $cidPattern -ForegroundColor Yellow
            Write-Host "  Make sure template contains: <img src=""$cidPattern"" ...>" -ForegroundColor Gray
            $allValid = $false
        }
    }
    
    return $allValid
}

function Test-ModuleImport {
    param([string]$ModulePath)
    
    try {
        Import-Module $ModulePath -Force -ErrorAction Stop
        Write-Host "✓ Module imported successfully" -ForegroundColor Green
        
        # Test exported functions
        $functions = @(
            'Get-EmailAuthenticationCredential',
            'Get-ProcessedHtmlTemplate',
            'New-EmailAlternateViewWithImages',
            'Send-HtmlEmail',
            'Send-BulkHtmlEmail'
        )
        
        $allFound = $true
        foreach ($func in $functions) {
            if (Get-Command $func -ErrorAction SilentlyContinue) {
                Write-Host "  ✓ Function available: " -NoNewline -ForegroundColor Green
                Write-Host $func -ForegroundColor Cyan
            }
            else {
                Write-Host "  ✗ Function NOT available: " -NoNewline -ForegroundColor Red
                Write-Host $func -ForegroundColor Red
                $allFound = $false
            }
        }
        
        return $allFound
    }
    catch {
        Write-Host "✗ Failed to import module: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Test-TemplateProcessing {
    param([string]$TemplatePath)
    
    try {
        $replacements = @{
            'CARD1_TITLE'   = 'Test Title 1'
            'CARD1_CONTENT' = 'Test Content 1'
            'CARD1_LINK'    = 'https://example.com/1'
            'CARD2_TITLE'   = 'Test Title 2'
            'CARD2_CONTENT' = 'Test Content 2'
            'CARD2_LINK'    = 'https://example.com/2'
        }
        
        $html = Get-ProcessedHtmlTemplate -TemplatePath $TemplatePath -Replacements $replacements -Encoding UTF8
        
        # Verify replacements occurred
        $allReplaced = $true
        foreach ($key in $replacements.Keys) {
            if ($html -notmatch [regex]::Escape($replacements[$key])) {
                Write-Host "⚠ Placeholder not replaced: $key" -ForegroundColor Yellow
                $allReplaced = $false
            }
        }
        
        if ($allReplaced) {
            Write-Host "✓ Template processing successful (all placeholders replaced)" -ForegroundColor Green
            Write-Host "  HTML length: $($html.Length) characters" -ForegroundColor Gray
        }
        else {
            Write-Host "⚠ Some placeholders were not replaced" -ForegroundColor Yellow
        }
        
        return $allReplaced
    }
    catch {
        Write-Host "✗ Template processing failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Test-Authentication {
    param([hashtable]$Config)
    
    if ($Config.AuthType -eq 'Basic') {
        Write-Host "Testing Basic Authentication configuration..." -ForegroundColor Cyan
        
        if ($Config.SmtpUser -eq 'user@example.com' -or $Config.SmtpPassword -eq 'password') {
            Write-Host "⚠ Using placeholder credentials - update before production use!" -ForegroundColor Yellow
            return $false
        }
        
        try {
            $cred = Get-EmailAuthenticationCredential `
                -AuthMethod 'Basic' `
                -Username $Config.SmtpUser `
                -Password $Config.SmtpPassword
            
            Write-Host "✓ Basic auth credential created successfully" -ForegroundColor Green
            Write-Host "  Username: $($Config.SmtpUser)" -ForegroundColor Cyan
            return $true
        }
        catch {
            Write-Host "✗ Failed to create Basic auth credential: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    elseif ($Config.AuthType -eq 'OAuth') {
        Write-Host "Testing OAuth2 configuration..." -ForegroundColor Cyan
        
        if ($Config.TenantId -eq 'your-tenant-id' -or 
            $Config.ClientId -eq 'your-client-id' -or 
            $Config.ClientSecret -eq 'your-secret') {
            Write-Host "⚠ Using placeholder OAuth values - update before production use!" -ForegroundColor Yellow
            return $false
        }
        
        try {
            Write-Host "  Attempting to acquire OAuth token (this may take a few seconds)..." -ForegroundColor Gray
            $cred = Get-EmailAuthenticationCredential `
                -AuthMethod 'OAuth' `
                -Username $Config.SmtpUser `
                -TenantId $Config.TenantId `
                -ClientId $Config.ClientId `
                -ClientSecret $Config.ClientSecret
            
            Write-Host "✓ OAuth token acquired successfully" -ForegroundColor Green
            Write-Host "  Tenant: $($Config.TenantId)" -ForegroundColor Cyan
            Write-Host "  Client: $($Config.ClientId)" -ForegroundColor Cyan
            return $true
        }
        catch {
            Write-Host "✗ Failed to acquire OAuth token: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  Common issues:" -ForegroundColor Gray
            Write-Host "    - Verify TenantId, ClientId, ClientSecret are correct" -ForegroundColor Gray
            Write-Host "    - Ensure Mail.Send permission is granted and admin consent given" -ForegroundColor Gray
            Write-Host "    - Check Azure AD app is enabled" -ForegroundColor Gray
            return $false
        }
    }
}

# ============================================================================
# MAIN VALIDATION
# ============================================================================

Write-Host "`n╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  M365 Digest Email System - Configuration Validator      ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

$results = @{
    ModuleImport        = $false
    TemplateExists      = $false
    TemplateProcessing  = $false
    ImagesValid         = $true
    CidsValid           = $false
    AttachmentsValid    = $true
    Authentication      = $false
}

# Test 1: Module Import
Write-Host "`n[1/7] Testing PowerShell Module..." -ForegroundColor Magenta
$results.ModuleImport = Test-ModuleImport -ModulePath $config.ModulePath

if (-not $results.ModuleImport) {
    Write-Host "`n✗ Cannot proceed without valid module. Fix module import first." -ForegroundColor Red
    exit 1
}

# Test 2: HTML Template
Write-Host "`n[2/7] Checking HTML Template..." -ForegroundColor Magenta
$results.TemplateExists = Test-FileExists -Path $config.HtmlTemplate -Description "HTML Template"

# Test 3: Template Processing
if ($results.TemplateExists) {
    Write-Host "`n[3/7] Testing Template Processing..." -ForegroundColor Magenta
    $results.TemplateProcessing = Test-TemplateProcessing -TemplatePath $config.HtmlTemplate
}
else {
    Write-Host "`n[3/7] Skipping template processing (template not found)" -ForegroundColor Yellow
}

# Test 4: Inline Images
Write-Host "`n[4/7] Validating Inline Images..." -ForegroundColor Magenta
foreach ($img in $config.InlineImages) {
    $valid = Test-ImageFile -Image $img
    if (-not $valid) { $results.ImagesValid = $false }
}

# Test 5: CID References
if ($results.TemplateExists) {
    Write-Host "`n[5/7] Checking CID References..." -ForegroundColor Magenta
    $results.CidsValid = Test-HtmlTemplateCids -TemplatePath $config.HtmlTemplate -InlineImages $config.InlineImages
}
else {
    Write-Host "`n[5/7] Skipping CID validation (template not found)" -ForegroundColor Yellow
}

# Test 6: Attachments
Write-Host "`n[6/7] Checking Attachments..." -ForegroundColor Magenta
foreach ($attachment in $config.Attachments) {
    $valid = Test-FileExists -Path $attachment -Description "Attachment"
    if (-not $valid) { $results.AttachmentsValid = $false }
}

# Test 7: Authentication
if (-not $SkipAuthTest) {
    Write-Host "`n[7/7] Testing Authentication..." -ForegroundColor Magenta
    $results.Authentication = Test-Authentication -Config $config
}
else {
    Write-Host "`n[7/7] Skipping authentication test (-SkipAuthTest specified)" -ForegroundColor Yellow
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  VALIDATION SUMMARY                                       ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

$totalTests = 0
$passedTests = 0

foreach ($key in $results.Keys) {
    $totalTests++
    $status = $results[$key]
    
    if ($status) {
        $passedTests++
        Write-Host "✓ " -NoNewline -ForegroundColor Green
    }
    else {
        Write-Host "✗ " -NoNewline -ForegroundColor Red
    }
    
    Write-Host "$key" -ForegroundColor Cyan
}

Write-Host "`nResults: $passedTests / $totalTests tests passed" -ForegroundColor $(if ($passedTests -eq $totalTests) { 'Green' } else { 'Yellow' })

if ($passedTests -eq $totalTests) {
    Write-Host "`n✓ All validation tests passed! System is ready for use." -ForegroundColor Green
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "  1. Update CSV file with recipient data" -ForegroundColor Gray
    Write-Host "  2. Run Send-M365Digest-BasicAuth.ps1 or Send-M365Digest-OAuth.ps1" -ForegroundColor Gray
    Write-Host "  3. Start with -ConfigMode Test for initial testing" -ForegroundColor Gray
    exit 0
}
else {
    Write-Host "`n⚠ Some validation tests failed. Review errors above and fix before proceeding." -ForegroundColor Yellow
    Write-Host "`nCommon fixes:" -ForegroundColor Cyan
    Write-Host "  - Update file paths in this script to match your environment" -ForegroundColor Gray
    Write-Host "  - Ensure inline images exist and match CID references" -ForegroundColor Gray
    Write-Host "  - Verify authentication credentials are correct" -ForegroundColor Gray
    exit 1
}
