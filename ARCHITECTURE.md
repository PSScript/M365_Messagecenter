# M365 Digest Email System - Architecture

## System Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                    M365 Digest Email System                     │
│                                                                 │
│  Professional HTML email sending with inline images & OAuth2   │
└─────────────────────────────────────────────────────────────────┘
```

## Component Architecture

```
┌───────────────────┐      ┌──────────────────────┐      ┌──────────────────┐
│                   │      │                      │      │                  │
│  CSV Data File    │─────▶│   PowerShell         │─────▶│  Microsoft 365   │
│  (Recipients)     │      │   Control Script     │      │  Exchange Online │
│                   │      │                      │      │                  │
└───────────────────┘      └──────────────────────┘      └──────────────────┘
                                      │
                                      │ Uses
                                      ▼
                           ┌──────────────────────┐
                           │                      │
                           │  M365DigestEmail     │
                           │  PowerShell Module   │
                           │                      │
                           └──────────────────────┘
                                      │
                    ┌─────────────────┼─────────────────┐
                    │                 │                 │
                    ▼                 ▼                 ▼
         ┌─────────────────┐ ┌──────────────┐ ┌──────────────────┐
         │  Authentication │ │   Template   │ │  SMTP Sending    │
         │     Module      │ │   Processor  │ │  with Batching   │
         └─────────────────┘ └──────────────┘ └──────────────────┘
                    │                 │                 │
         ┌──────────┴────────┐        │        ┌────────┴──────────┐
         │                   │        │        │                   │
         ▼                   ▼        ▼        ▼                   ▼
    ┌────────┐         ┌────────┐ ┌──────┐ ┌────────┐      ┌──────────┐
    │ Basic  │         │ OAuth2 │ │ HTML │ │ Inline │      │Checkpoint│
    │  Auth  │         │  (AAD) │ │ Body │ │ Images │      │   File   │
    └────────┘         └────────┘ └──────┘ └────────┘      └──────────┘
```

## Data Flow Diagram

```
START
  │
  ▼
┌─────────────────────────────────┐
│ Load Configuration              │
│ - CSV path                      │
│ - HTML template                 │
│ - Inline images                 │
│ - SMTP settings                 │
│ - Auth method                   │
└─────────────────┬───────────────┘
                  │
                  ▼
┌─────────────────────────────────┐
│ Validate Configuration          │
│ - Check files exist             │
│ - Verify image CIDs             │
│ - Test authentication           │
└─────────────────┬───────────────┘
                  │
                  ▼
┌─────────────────────────────────┐
│ Load Checkpoint File            │
│ (Resume from previous run)      │
└─────────────────┬───────────────┘
                  │
                  ▼
┌─────────────────────────────────┐
│ Import & Parse CSV              │
│ - Filter already-sent emails    │
│ - Build recipient objects       │
└─────────────────┬───────────────┘
                  │
                  ▼
┌─────────────────────────────────┐
│ For Each Batch (20 emails):     │
│                                 │
│  ┌─────────────────────────┐   │
│  │ For Each Recipient:      │   │
│  │                          │   │
│  │  1. Process HTML         │   │
│  │     template             │   │
│  │  2. Replace placeholders │   │
│  │  3. Create AlternateView │   │
│  │  4. Embed inline images  │   │
│  │  5. Add attachments      │   │
│  │  6. Send via SMTP        │   │
│  │  7. Retry if failed (3x) │   │
│  │  8. Write to checkpoint  │   │
│  │                          │   │
│  └─────────────────────────┘   │
│                                 │
│  Wait for batch window (3 min) │
└─────────────────┬───────────────┘
                  │
                  ▼
┌─────────────────────────────────┐
│ Campaign Complete               │
│ - Display statistics            │
│ - Close connections             │
└─────────────────────────────────┘
  │
  ▼
END
```

## Module Function Hierarchy

```
M365DigestEmailModule.psm1
│
├── Get-EmailAuthenticationCredential
│   │
│   ├── Basic Authentication
│   │   └── PSCredential (username/password)
│   │
│   └── OAuth2 Authentication
│       ├── Token Endpoint Request
│       ├── Client Credentials Flow
│       └── PSCredential (username/token)
│
├── Get-ProcessedHtmlTemplate
│   │
│   ├── Load HTML file
│   ├── HTML-encode values (XSS protection)
│   └── Replace placeholders
│
├── New-EmailAlternateViewWithImages
│   │
│   ├── Create AlternateView object
│   ├── For each inline image:
│   │   ├── Create LinkedResource
│   │   ├── Set ContentType (image/png, image/jpeg)
│   │   ├── Set ContentId (CID reference)
│   │   └── Add to AlternateView
│   └── Return AlternateView
│
├── Send-HtmlEmail
│   │
│   ├── Create MailMessage object
│   ├── Set To, From, Subject, BCC
│   ├── Create AlternateView with images
│   ├── Add attachments
│   ├── Create SmtpClient
│   ├── Set credentials & SSL
│   ├── Send email
│   └── Dispose objects (cleanup)
│
└── Send-BulkHtmlEmail
    │
    ├── Load checkpoint file
    ├── Filter pending recipients
    ├── For each batch:
    │   ├── Process template per recipient
    │   ├── Call Send-HtmlEmail
    │   ├── Retry logic (exponential backoff)
    │   ├── Update checkpoint file
    │   └── Wait for window interval
    └── Return statistics
```

## Authentication Flow

### Basic Authentication
```
User Script
    │
    ▼
Get-EmailAuthenticationCredential
    │
    ├── Convert password to SecureString
    │
    ├── Create PSCredential object
    │   (username + secure password)
    │
    └── Return credential
          │
          ▼
SmtpClient.Credentials = PSCredential
          │
          ▼
    SMTP AUTH LOGIN
    (Base64 encoded credentials)
          │
          ▼
    Office 365 validates
          │
          ▼
    Authentication SUCCESS/FAIL
```

### OAuth2 Authentication
```
User Script
    │
    ▼
Get-EmailAuthenticationCredential
    │
    ├── Build token request
    │   - TenantId
    │   - ClientId  
    │   - ClientSecret
    │   - Scope: outlook.office365.com/.default
    │
    ├── POST to token endpoint
    │   https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
    │
    ├── Receive access token
    │
    ├── Create PSCredential
    │   (username + token as password)
    │
    └── Return credential
          │
          ▼
SmtpClient.Credentials = PSCredential
          │
          ▼
    SMTP AUTH XOAUTH2
    (OAuth bearer token)
          │
          ▼
    Azure AD validates token
          │
          ▼
    Authentication SUCCESS/FAIL
```

## Email Construction

```
┌────────────────────────────────────────────────────┐
│                   MailMessage                      │
├────────────────────────────────────────────────────┤
│                                                    │
│  Headers:                                          │
│    From: sender@example.com                        │
│    To: recipient@example.com                       │
│    Subject: Microsoft 365 Monthly Digest           │
│    BCC: admin@example.com                          │
│                                                    │
│  ┌──────────────────────────────────────────────┐ │
│  │          AlternateView (text/html)           │ │
│  ├──────────────────────────────────────────────┤ │
│  │                                              │ │
│  │  <html>                                      │ │
│  │    <body>                                    │ │
│  │      ...HTML content...                      │ │
│  │      <img src="cid:logo1">                   │ │
│  │      <img src="cid:icon1">                   │ │
│  │    </body>                                   │ │
│  │  </html>                                     │ │
│  │                                              │ │
│  │  ┌────────────────────────────────────────┐ │ │
│  │  │    LinkedResource (logo1)              │ │ │
│  │  │    - ContentId: logo1                  │ │ │
│  │  │    - ContentType: image/png            │ │ │
│  │  │    - TransferEncoding: Base64          │ │ │
│  │  │    - Data: [PNG bytes]                 │ │ │
│  │  └────────────────────────────────────────┘ │ │
│  │                                              │ │
│  │  ┌────────────────────────────────────────┐ │ │
│  │  │    LinkedResource (icon1)              │ │ │
│  │  │    - ContentId: icon1                  │ │ │
│  │  │    - ContentType: image/png            │ │ │
│  │  │    - TransferEncoding: Base64          │ │ │
│  │  │    - Data: [PNG bytes]                 │ │ │
│  │  └────────────────────────────────────────┘ │ │
│  │                                              │ │
│  └──────────────────────────────────────────────┘ │
│                                                    │
│  Attachments:                                      │
│    ├── Manual.pdf (application/pdf)                │
│    └── Guide.pdf (application/pdf)                 │
│                                                    │
└────────────────────────────────────────────────────┘
```

## Batch Processing Flow

```
Batch 1 (emails 1-20)
├── Email 1  ────────┐
├── Email 2          │
├── Email 3          │
├── ...              │  Sending Phase
├── Email 19         │  (~30-60 seconds)
├── Email 20  ───────┘
└── WAIT 3 minutes ◄─── Rate Limiting Window

Batch 2 (emails 21-40)
├── Email 21 ────────┐
├── Email 22         │
├── ...              │  Sending Phase
├── Email 39         │
├── Email 40  ───────┘
└── WAIT 3 minutes

... Continue for all batches
```

## Error Handling & Retry Logic

```
Send Email
    │
    ▼
┌─────────────────────┐
│   Attempt 1         │
│   (immediate)       │
└──────┬──────────────┘
       │
   ┌───┴───┐
   │ OK?   │
   └───┬───┘
       │
   ┌───┴───────────────────────┐
   │ Success → Write checkpoint│
   │ Fail → Wait 2s             │
   └───┬───────────────────────┘
       │
       ▼
┌─────────────────────┐
│   Attempt 2         │
│   (after 2s)        │
└──────┬──────────────┘
       │
   ┌───┴───┐
   │ OK?   │
   └───┬───┘
       │
   ┌───┴───────────────────────┐
   │ Success → Write checkpoint│
   │ Fail → Wait 4s             │
   └───┬───────────────────────┘
       │
       ▼
┌─────────────────────┐
│   Attempt 3         │
│   (after 4s)        │
└──────┬──────────────┘
       │
   ┌───┴───┐
   │ OK?   │
   └───┬───┘
       │
   ┌───┴───────────────────────┐
   │ Success → Write checkpoint│
   │ Fail → Log permanent error │
   └───────────────────────────┘
```

## Checkpoint System

```
┌───────────────────────────────────────────────────┐
│          Email Campaign Execution                 │
├───────────────────────────────────────────────────┤
│                                                   │
│  Email 1 → Success ──┬─→ checkpoint.txt          │
│  Email 2 → Success ──┤     user1@example.com     │
│  Email 3 → Success ──┤     user2@example.com     │
│  Email 4 → FAIL      │     user3@example.com     │
│  Email 5 → Success ──┤     user5@example.com     │
│  Email 6 → Success ──┘                            │
│                                                   │
│  ╔════════════════════════════════════╗           │
│  ║  SCRIPT INTERRUPTED (Ctrl+C)      ║           │
│  ╚════════════════════════════════════╝           │
│                                                   │
│  ┌─────────────────────────────────┐             │
│  │  Restart script                 │             │
│  │  ↓                               │             │
│  │  Load checkpoint.txt             │             │
│  │  ↓                               │             │
│  │  Skip: user1, user2, user3, user5│             │
│  │  ↓                               │             │
│  │  Resume: Email 4, 7, 8, 9...    │             │
│  └─────────────────────────────────┘             │
│                                                   │
└───────────────────────────────────────────────────┘
```

## File Dependencies

```
Project Root
│
├── M365DigestEmailModule.psm1         ◄─── Core module (REQUIRED)
│
├── M365_Digest_Template.htm           ◄─── HTML template (REQUIRED)
│
├── Send-M365Digest-BasicAuth.ps1     ◄─── Example script (Basic)
├── Send-M365Digest-OAuth.ps1         ◄─── Example script (OAuth)
├── Test-M365DigestConfig.ps1         ◄─── Configuration validator
│
├── recipients.csv                     ◄─── Recipient data (USER PROVIDED)
│
├── Images/                            ◄─── Inline images (USER PROVIDED)
│   ├── datagroup_logo.png
│   ├── m365_icon.png
│   ├── exchange_icon.png
│   └── sharepoint_icon.png
│
├── Attachments/                       ◄─── PDF files (USER PROVIDED)
│   ├── Manual1.pdf
│   ├── Manual2.pdf
│   └── Manual3.pdf
│
└── Logs/                              ◄─── Runtime files (AUTO GENERATED)
    ├── smtp_send_checkpoint.txt
    └── campaign_YYYYMMDD_HHmmss.log
```

## Performance Characteristics

```
┌─────────────────────────────────────────────────────┐
│  Campaign Size: 1000 recipients                     │
│  Batch Size: 20 emails                              │
│  Window Interval: 3 minutes                         │
├─────────────────────────────────────────────────────┤
│                                                     │
│  Total Batches: 50                                  │
│  Total Time: ~150 minutes (2.5 hours)               │
│                                                     │
│  Time Breakdown:                                    │
│    - Sending: ~50 minutes (1 min/batch)             │
│    - Waiting: ~147 minutes (3 min × 49 intervals)   │
│    - Retries: ~3 minutes (if 1% failure rate)       │
│                                                     │
│  Throughput: ~6.7 emails/minute (sustainable)       │
│                                                     │
└─────────────────────────────────────────────────────┘
```

---

## Security Layers

```
┌─────────────────────────────────────────────────────┐
│                  Security Measures                  │
├─────────────────────────────────────────────────────┤
│                                                     │
│  1. HTML Encoding                                   │
│     └─→ Prevents XSS injection attacks              │
│                                                     │
│  2. Secure Credentials                              │
│     ├─→ PSCredential with SecureString              │
│     └─→ OAuth tokens (short-lived)                  │
│                                                     │
│  3. TLS/SSL Encryption                              │
│     └─→ SMTP over TLS (port 587)                    │
│                                                     │
│  4. Application Permissions                         │
│     └─→ OAuth: Mail.Send only (least privilege)     │
│                                                     │
│  5. Input Validation                                │
│     ├─→ File path validation                        │
│     ├─→ Email address format check                  │
│     └─→ Content size limits                         │
│                                                     │
└─────────────────────────────────────────────────────┘
```

---

**Architecture Documentation v1.0.0** | Built for Microsoft 365 Enterprise Email Campaigns
