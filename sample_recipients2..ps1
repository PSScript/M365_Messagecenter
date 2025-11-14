# ================== CONFIG ==================
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Paths
#$csvPath       = "C:\Temp\2\Wave_W3_usersecrets_neu.csv"     # ; separated
#$csvPath       = "C:\Temp\master_users_all_merged_with_wave4.csv"     # ; separated
#$csvPath       = "C:\Temp\Ditzingen_WithOTP.csv"     # ; separated
#$csvPath       = "C:\Temp\Migrationsuser_Links_TEST2.csv"     # ; separated
#$csvPath       = "C:\Temp\master_users_all_merged_with_waveSTGVaihingen.csv"     # ; separated
#$csvPath       = "C:\Temp\Export_Welle_Mail_5_Final.csv"     # ; separated
$csvPath       = "C:\Temp\sample_recipients2.csv"     # ; separated

#$htmlTemplate  = "C:\temp\kopano_woche4_w2.htm"  # Template8.htm"
#$htmlTemplate  = "C:\Temp\kopano_woche4_w6.htm"
#$htmlTemplate  = "C:\temp\E-Mail3_Plain-Kennwort_DGM-AnwenderInnen-ohne-PCiP_Umstellungswoche-Montag1.htm"

# Abschluss_E-Mail_alle_DGM_AnwenderInnen.htm
# WillkommensE-Mail_OHNE_PIP_Umstellungstag.htm
# Willkommens-E-Mail_AnwenderInnen_PIP_Umstellungstag.htm

$htmlTemplate  = "C:\temp\M365_Digest_Template.htm"  # Template8.htm"

#$htmlTemplate  = "C:\temp\WillkommensE-Mail_OHNE_PIP_Umstellungstag.htm"  # Template8.htm"

#$htmlTemplate  = "C:\temp\2025-10-02_E-Mail 2_DGM-AnwenderInnen Ditzingen_PASSWORDSTRING.htm"  # Template8.htm"

$pdf1          = "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Smartphone.pdf"
$pdf2          = "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Telefon.pdf"
$pdf3          = "C:\temp\Anleitung_Postfach_und_Dateiablage_auf_dem_Computer_einrichten.pdf"
$pdf4          = "C:\temp\Anleitung_Einrichtung-Thunderbird.pdf"

# Resume checkpoint
$checkpoint    = "C:\temp\smtp_send_checkpoint_abschluss.txt"  # stores emails already sent

$Attachments  = @(
  # @($pdf1,$pdf2,$pdf3,$pdf4)
)

$inlineImages  = @(
  # 'C:\temp\logo.png|logo1'   # format: path|cid ; optional
)

# SMTP
$smtpServer    = "smtp.office365.com"
$smtpPort      = 587
$smtpUser      = "adm_huebener@elkw.de"
$smtpPassword  = 'fdxoF!U$EIv!23U9'  # use a secure app password/secret
$from          = "OKR.noreply@ELK-WUE.DE"
$Bcc           = "jan.huebener@elkw.de"
$subject       = "Willkommen im DGM 2.0"

# Throughput control
$batchSize     = 20           # messages per window
$windowMinutes = 0.04            # minutes per window


# ================ PREPARE ====================
$securePass = ConvertTo-SecureString $smtpPassword -AsPlainText -Force
$cred       = New-Object System.Management.Automation.PSCredential($smtpUser, $securePass)

# Templates & CSV (use 1252 for German Excel CSVs)
$htmlRaw = Get-Content -LiteralPath $htmlTemplate -Raw -Encoding UTF7
$data    = Import-Csv -LiteralPath $csvPath -Delimiter ';' -Encoding UTF8 # -Encoding Windows-1252

# checkpoint list
$sent = New-Object System.Collections.Generic.HashSet[string]
if(Test-Path $checkpoint){
  Get-Content $checkpoint | ForEach-Object { [void]$sent.Add($_.Trim()) }
}

# Utilities
Add-Type -AssemblyName System.Web   # for HtmlEncode

function New-AlternateViewWithCid {
  param(
    [Parameter(Mandatory=$true)] [string]$Html,
    [Parameter()] [string[]]$InlineSpecs # 'path|cid'
  )
  $htmlOut = $Html
  $altView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($htmlOut, $null, "text/html")

  foreach($spec in ($InlineSpecs | Where-Object { $_ })){
    $bits = $spec -split '\|', 2
    $path = $bits[0]
    $cid  = if($bits.Count -gt 1 -and $bits[1]){ $bits[1] } else { [IO.Path]::GetFileName($path) }
    if(-not (Test-Path $path)){ throw "Inline image not found: $path" }

    # replace {{cid:NAME}} tokens if present
    $token = "{{cid:$cid}}"
    if($htmlOut.Contains($token)){
      $htmlOut = $htmlOut -replace [regex]::Escape($token), "<img src=""cid:$cid"">"
      $altView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($htmlOut, $null, "text/html")
    }

    $lr = New-Object System.Net.Mail.LinkedResource($path)
    $ext = [IO.Path]::GetExtension($path).TrimStart('.').ToLower()
    if($ext -eq 'jpg'){ $ext = 'jpeg' }
    $lr.ContentType = New-Object System.Net.Mime.ContentType("image/$ext")
    $lr.ContentId = $cid
    $lr.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
    [void]$altView.LinkedResources.Add($lr)
  }

  return $altView
}

function Send-OneSmtp {
  param(
    [Parameter(Mandatory=$true)][string]$To,
    [Parameter(Mandatory=$true)][string]$Subject,
    [Parameter(Mandatory=$true)][string]$HtmlBody,
    [string[]]$Attachments,
    [string[]]$InlineSpecs
  )

  $mailMessage = $null
  $smtpClient  = $null
  try {
    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.BodyEncoding    = [System.Text.Encoding]::UTF8
    $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
    $mailMessage.From            = $from
    $mailMessage.To.Add($To)
    if($Bcc){ $mailMessage.Bcc.Add($Bcc) }
    $mailMessage.Subject         = $Subject
    # Use AlternateView to bind LinkedResources; Body can remain empty
    $alt = New-AlternateViewWithCid -Html $HtmlBody -InlineSpecs $InlineImages
    $mailMessage.AlternateViews.Add($alt) | Out-Null
    $mailMessage.IsBodyHtml = $true

    foreach($a in ($Attachments | Where-Object { $_ })){
      if(Test-Path $a){ $mailMessage.Attachments.Add([System.Net.Mail.Attachment]::new($a)) | Out-Null }
      else { Write-Warning "Attachment missing, skipping: $a" }
    }

    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtpClient.EnableSsl   = $true
    $smtpClient.Credentials = $cred

    $smtpClient.Send($mailMessage)
    return $true
  }
  catch {
    Write-Warning "Send failed → $To : $($_.Exception.Message)"
    return $false
  }
  finally {
    if($mailMessage){ $mailMessage.Dispose() }
    if($smtpClient){ $smtpClient.Dispose() }
  }
}

# ================ SEND LOOP (BATCHED) =================

$rows = @()
foreach($entry in $data){
  $to = ($entry.email).Trim()
  if([string]::IsNullOrWhiteSpace($to)){ continue }
  if($sent.Contains($to)){ continue }           # skip already done (resume)
  $displayNameRaw = $entry.DisplayName_email
  $passwordRaw    = $entry.password
  $secretlinkRaw    = $entry.secret_link


  # HTML-safe insertion (text nodes)
         $displayName = [System.Web.HttpUtility]::HtmlEncode($displayNameRaw)
         $password    = [System.Web.HttpUtility]::HtmlEncode($passwordRaw)
  #$MigrationsDatum    = [System.Web.HttpUtility]::HtmlEncode($MigrationsDatum)
         $secretlink    = [System.Web.HttpUtility]::HtmlEncode($secretlinkRaw)

  $htmlBody = $htmlRaw.Replace("DISPLAYNAME", $displayName)
  $htmlBody = $htmlBody.Replace("PASSWORDSTRING", $password)
  $htmlBody = $htmlBody.Replace("SECRETLINK", $secretlink)
  $rows += [pscustomobject]@{
    To   = $to
    Html = $htmlBody
  }
}

if($rows.Count -eq 0){
  Write-Host "Nothing to send (all recipients already in checkpoint or CSV empty)."
  return
}

Write-Host ("Will send {0} messages in batches of {1} every {2} minutes..." -f $rows.Count, $batchSize, $windowMinutes)

# simple 3x retry wrapper with small backoff
function Try-SendWithRetry {
  param([pscustomobject]$row)
  $attempt=0
  $delay=2
  while($attempt -lt 3){
    $attempt++
    if(Send-OneSmtp -To $row.To -Subject $subject -HtmlBody $row.Html -Attachments $Attachments -InlineSpecs $inlineImages){
   # if(Send-OneSmtp -To $row.To -Subject $subject -HtmlBody $row.Html -InlineSpecs $inlineImages){
      # checkpoint immediately on success
      Add-Content -LiteralPath $checkpoint -Value $row.To
      return
    } else {
      Start-Sleep -Seconds $delay
      $delay = [Math]::Min($delay * 2, 30)
    }
  }
  Write-Warning "❌ Permanent failure after 3 attempts → $($row.To)"
}

# Partition into windows of size $batchSize
$sw = [System.Diagnostics.Stopwatch]::StartNew()
for($offset=0; $offset -lt $rows.Count; $offset += $batchSize){
  $windowStart = $sw.Elapsed
  $batch = $rows[$offset..([Math]::Min($offset + $batchSize - 1, $rows.Count - 1))]

  Write-Host ("=== Window starting {0:HH:mm:ss}, sending {1} messages ===" -f (Get-Date), $batch.Count)

  foreach($row in $batch){
    Try-SendWithRetry -row $row
    Start-Sleep -Milliseconds (150 + (Get-Random -Minimum 0 -Maximum 300))  # tiny jitter
  }

  # Enforce 3-minute window spacing
  $elapsed = $sw.Elapsed - $windowStart
  $target  = [TimeSpan]::FromMinutes($windowMinutes)
  if($elapsed -lt $target){
    $sleep = [int](($target - $elapsed).TotalMilliseconds)
    if($sleep -gt 0){
      Write-Host ("Waiting {0}s to honor window spacing..." -f $sleep)
      Start-Sleep -Milliseconds $sleep
    }
  }
}

Write-Host "✅ All done. Sent: $($rows.Count) (plus any checkpointed earlier)."
