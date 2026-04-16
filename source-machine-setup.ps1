# ============================================================
# File Sync — Source Machine Setup
# รันครั้งเดียว ทำทุกอย่างอัตโนมัติ
#
# Usage:
#   .\source-machine-setup.ps1
#   .\source-machine-setup.ps1 -PersonName "สมชาย"
#   .\source-machine-setup.ps1 -PersonName "สมชาย" -SyncFolder "C:\Users\User\Documents\Reports"
#   .\source-machine-setup.ps1 -PersonName "ClientA" -DepartmentName "Finance" -SyncFolder "C:\Finance"
#   .\source-machine-setup.ps1 -PersonName "ClientA" -DepartmentName "Marketing" -SyncFolder "C:\Marketing"
# ============================================================

param(
    [string]$PersonName      = "User",
    [string]$SyncFolder      = "C:\sync-out",
    [string]$DepartmentName  = ""
)

$kayDeviceId  = "TIC5RDP-LZZ7SJL-4ZILCW3-22Z3BVV-C5WOPBH-DXJ73IM-RRFZE6W-CZZQGQB"  # Mac Mini (updated 2026-04-10)
$syncFolder   = $SyncFolder
$installDir   = "$env:LOCALAPPDATA\FileSyncApp"
$downloadUrl  = "https://github.com/syncthing/syncthing/releases/download/v2.0.16/syncthing-windows-amd64-v2.0.16.zip"
$zipPath      = "$installDir\syncthing.zip"
$syncthingExe = "$installDir\syncthing-windows-amd64-v2.0.16\syncthing.exe"
$startupFolder = [Environment]::GetFolderPath('Startup')

function Wait-SyncthingReady {
    Write-Host "  Waiting for Syncthing to start..." -NoNewline
    for ($i = 0; $i -lt 30; $i++) {
        try {
            Invoke-RestMethod "http://127.0.0.1:8384/rest/system/ping" -Headers @{ "X-API-Key" = $script:apiKey } | Out-Null
            Write-Host " Ready"
            return $true
        } catch { Start-Sleep -Seconds 1; Write-Host "." -NoNewline }
    }
    Write-Host " Timeout"
    return $false
}

function Get-SyncthingApiKey {
    $configPath = "$env:LOCALAPPDATA\Syncthing\config.xml"
    for ($i = 0; $i -lt 15; $i++) {
        if (Test-Path $configPath) {
            $xml = [xml](Get-Content $configPath)
            $key = $xml.configuration.gui.apikey
            if ($key) { return $key }
        }
        Start-Sleep -Seconds 1
    }
    return $null
}

# ── Step 1: สร้างโฟลเดอร์ที่จำเป็น ──────────────────────────
Write-Host "[1/6] Creating folders..."
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
New-Item -ItemType Directory -Force -Path $syncFolder | Out-Null
Write-Host "  $syncFolder"

# ── Step 2: ดาวน์โหลด Syncthing ────────────────────────────
if (-not (Test-Path $syncthingExe)) {
    Write-Host "[2/6] Downloading Syncthing..."
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath
    Expand-Archive -Path $zipPath -DestinationPath $installDir -Force
    Remove-Item $zipPath
    Write-Host "  Done"
} else {
    Write-Host "[2/6] Syncthing already installed, skipping."
}

# ── Step 3: Start Syncthing (ครั้งแรก สร้าง config) ─────────
Write-Host "[3/6] Starting Syncthing..."
Start-Process -FilePath $syncthingExe -ArgumentList "--no-browser" -WindowStyle Hidden
Start-Sleep -Seconds 3

$script:apiKey = Get-SyncthingApiKey
if (-not $script:apiKey) {
    Write-Host "  ERROR: Could not get API key. Exiting."
    exit 1
}
Write-Host "  API key: $($script:apiKey.Substring(0,8))..."

if (-not (Wait-SyncthingReady)) {
    Write-Host "  ERROR: Syncthing did not start. Exiting."
    exit 1
}

$headers = @{ "X-API-Key" = $script:apiKey }

# ── Step 4: เพิ่ม Kay เป็น device ───────────────────────────
Write-Host "[4/6] Adding Kay's device..."
$deviceBody = @{
    deviceID = $kayDeviceId
    name     = "Mac Mini"
    addresses = @("dynamic")
    autoAcceptFolders = $false
} | ConvertTo-Json

Write-Host "  Client: $PersonName"
if ($DepartmentName -ne "") {
    Write-Host "  Department: $DepartmentName"
}

try {
    Invoke-RestMethod "http://127.0.0.1:8384/rest/config/devices" `
        -Method POST -Headers $headers `
        -ContentType "application/json" -Body $deviceBody | Out-Null
    Write-Host "  Added: Mac Mini"
} catch {
    Write-Host "  Already exists or error: $($_.Exception.Message)"
}

# ── Step 5: Share โฟลเดอร์ไปหา Kay ──────────────────────────
Write-Host "[5/6] Sharing folder with Kay..."
$safeClient = $PersonName -replace '[^a-zA-Z0-9]', '-'

if ($DepartmentName -ne "") {
    $safeDept  = $DepartmentName -replace '[^a-zA-Z0-9]', '-'
    $folderId  = "origo-sync-$safeClient-$safeDept".ToLower()
    $folderLabel = "Origo Sync - $PersonName ($DepartmentName)"
} else {
    $folderId    = "origo-sync-$safeClient".ToLower()
    $folderLabel = "Origo Sync - $PersonName"
}

$folderBody = @{
    id      = $folderId
    label   = $folderLabel
    path    = $syncFolder
    type    = "sendonly"
    devices = @(
        @{ deviceID = $kayDeviceId }
    )
} | ConvertTo-Json -Depth 5

try {
    Invoke-RestMethod "http://127.0.0.1:8384/rest/config/folders" `
        -Method POST -Headers $headers `
        -ContentType "application/json" -Body $folderBody | Out-Null
    Write-Host "  Shared: $syncFolder → Mac Mini"
} catch {
    Write-Host "  Already exists or error: $($_.Exception.Message)"
}

# ── Step 6: ตั้ง Startup ─────────────────────────────────────
Write-Host "[6/6] Setting up auto-start..."
Stop-Process -Name "syncthing" -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$s = (New-Object -COM WScript.Shell).CreateShortcut("$startupFolder\Syncthing.lnk")
$s.TargetPath = $syncthingExe
$s.Arguments = "--no-browser"
$s.WindowStyle = 7
$s.Save()
Write-Host "  Syncthing will auto-start on login"

# Start Syncthing again (hidden)
Start-Process -FilePath $syncthingExe -ArgumentList "--no-browser" -WindowStyle Hidden

# ── Done ─────────────────────────────────────────────────────
Write-Host ""
Write-Host "============================================"
Write-Host " Setup complete!"
Write-Host " Client      : $PersonName"
if ($DepartmentName -ne "") {
    Write-Host " Department  : $DepartmentName"
    Write-Host " Mac Mini path: incoming/$PersonName/Data Raw/$DepartmentName/"
}
Write-Host " Sync folder : $syncFolder"
Write-Host " Folder ID   : $folderId"
Write-Host " Drop files in that folder to send to Kay."
Write-Host " Waiting for Kay to accept the connection."
Write-Host "============================================"
