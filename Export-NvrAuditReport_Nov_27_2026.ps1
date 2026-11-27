# 1. Ensure ImportExcel module is available for native .xlsx creation & styling
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module for .xlsx export..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
}
Import-Module ImportExcel

# Ignore self-signed SSL certificates & enable modern TLS for pre-flight credential check
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

# 2. Interactive Prompts (Held in RAM only, never saved to disk)
do {
    $ExportUserName = Read-Host "Enter Full Name of User Exporting This Report (Required)"
    if ([string]::IsNullOrWhiteSpace($ExportUserName)) {
        Write-Host "  [!] User name cannot be empty. Please manually enter your name." -ForegroundColor Yellow
    }
} while ([string]::IsNullOrWhiteSpace($ExportUserName))
$ExportUserName = $ExportUserName.Trim()

$RawInput = Read-Host "Enter NVR IP(s) or Domain(s) (comma-separated for simultaneous export, e.g., 192.168.1.1, 192.168.1.2)"

$HostList = $RawInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique
if ($HostList.Count -eq 0) {
    Write-Host "No NVR address provided. Exiting." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# Pre-flight Credential Validation Loop (Keeps asking until username & password are valid)
$Credential = $null
$SecurePass = $null
$AuthValid  = $false

while (-not $AuthValid) {
    $User       = Read-Host "Enter NVR Login Username"
    $SecurePass = Read-Host "Enter NVR Login Password" -AsSecureString
    $Credential = New-Object System.Management.Automation.PSCredential ($User, $SecurePass)

    Write-Host "Verifying credentials..." -ForegroundColor Cyan
    $WrongCreds     = $false
    $AnyReachable   = $false

    foreach ($TestHost in $HostList) {
        $TestUrl = if ($TestHost -notmatch "^https?://") { "http://$TestHost" } else { $TestHost.TrimEnd('/') }
        try {
            $null = Invoke-RestMethod -Uri "$TestUrl/ISAPI/System/deviceInfo" -Credential $Credential -TimeoutSec 8 -ErrorAction Stop
            $AnyReachable = $true
        } catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($statusCode -eq 401 -or $_.Exception.Message -match "401|Unauthorized") {
                $WrongCreds = $true
                break
            } else {
                Write-Host "  [!] Warning: Could not reach $TestHost ($($_.Exception.Message))" -ForegroundColor Yellow
            }
        }
    }

    if ($WrongCreds) {
        Write-Host "`n  [!] Username or password is wrong. Please re-enter the username and password.`n" -ForegroundColor Red
        if ($SecurePass) { $SecurePass.Dispose() }
        $Credential = $null
    } elseif (-not $AnyReachable) {
        Write-Host "`n  [!] None of the entered NVR addresses could be reached. Check IP/network and try again." -ForegroundColor Red
        if ($SecurePass) { $SecurePass.Dispose() }
        Read-Host "Press Enter to exit"
        exit
    } else {
        Write-Host "  [OK] Credentials verified." -ForegroundColor Green
        $AuthValid = $true
    }
}

$DesktopPath = [Environment]::GetFolderPath("Desktop")

# 3. Worker ScriptBlock executed simultaneously for every NVR
$NvrWorkerScript = {
    param(
        [string]$RawHost,
        [string]$ExportUserName,
        [PSCredential]$Credential,
        [string]$DesktopPath
    )

    Import-Module ImportExcel

    # Ignore self-signed SSL certificates & enable modern TLS inside worker job
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    [System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

    function Invoke-IsapiGet {
        param([string]$Uri, [PSCredential]$Cred, [int]$Timeout = 10)
        try {
            return [xml](Invoke-RestMethod -Uri $Uri -Credential $Cred -TimeoutSec $Timeout -ErrorAction Stop)
        } catch {
            return $null
        }
    }

    function Invoke-IsapiPost {
        param([string]$Uri, [string]$Body, [PSCredential]$Cred, [int]$Timeout = 10)
        try {
            return [xml](Invoke-RestMethod -Uri $Uri -Method Post -Body $Body -ContentType "application/xml" -Credential $Cred -TimeoutSec $Timeout -ErrorAction Stop)
        } catch {
            return $null
        }
    }

    function Get-LatestClipInRange {
        param(
            [string]$NvrUrl,
            [int]$TrackId,
            [datetime]$StartUtc,
            [datetime]$EndUtc,
            [PSCredential]$Cred
        )

        $startStr = $StartUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $endStr   = $EndUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")

        $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<CMSearchDescription>
    <searchID>$([guid]::NewGuid().ToString())</searchID>
    <trackList><trackID>$TrackId</trackID></trackList>
    <timeSpanList><timeSpan><startTime>$startStr</startTime><endTime>$endStr</endTime></timeSpan></timeSpanList>
    <maxResults>100</maxResults>
    <searchResultPosition>0</searchResultPosition>
</CMSearchDescription>
"@
        $res = Invoke-IsapiPost -Uri "$NvrUrl/ISAPI/ContentMgmt/search" -Body $xml -Cred $Cred -Timeout 8
        if (-not $res -or -not $res.CMSearchResult.matchList.searchMatchItem) { return $null }

        $items = @($res.CMSearchResult.matchList.searchMatchItem)
        $pos   = 0
        $guard = 0

        # Safely step forward if >100 clips exist in this window (hard-capped at 10 pages so it can NEVER hang)
        while ($items.Count -eq 100 -and $guard -lt 10) {
            $guard++
            $pos += 100
            $pageXml = @"
<?xml version="1.0" encoding="utf-8"?>
<CMSearchDescription>
    <searchID>$([guid]::NewGuid().ToString())</searchID>
    <trackList><trackID>$TrackId</trackID></trackList>
    <timeSpanList><timeSpan><startTime>$startStr</startTime><endTime>$endStr</endTime></timeSpan></timeSpanList>
    <maxResults>100</maxResults>
    <searchResultPosition>$pos</searchResultPosition>
</CMSearchDescription>
"@
            $nextRes = Invoke-IsapiPost -Uri "$NvrUrl/ISAPI/ContentMgmt/search" -Body $pageXml -Cred $Cred -Timeout 8
            if (-not $nextRes -or -not $nextRes.CMSearchResult.matchList.searchMatchItem) { break }
            $nextItems = @($nextRes.CMSearchResult.matchList.searchMatchItem)
            if ($nextItems.Count -eq 0) { break }
            $items = $nextItems
        }

        return ($items | Select-Object -Last 1)
    }

    function Get-RecordingSpan {
        param(
            [string]$NvrUrl,
            [int]$TrackId,
            [PSCredential]$Cred
        )

        $oldestRec = "No Footage"
        $newestRec = "No Footage"
        $retDays   = 0
        $dtOld     = $null
        $dtNew     = $null

        # 1. Oldest clip (first match from 2015)
        $oldestXml = @"
<?xml version="1.0" encoding="utf-8"?>
<CMSearchDescription>
    <searchID>$([guid]::NewGuid().ToString())</searchID>
    <trackList><trackID>$TrackId</trackID></trackList>
    <timeSpanList><timeSpan><startTime>2015-01-01T00:00:00Z</startTime><endTime>2035-01-01T00:00:00Z</endTime></timeSpan></timeSpanList>
    <maxResults>1</maxResults>
    <searchResultPosition>0</searchResultPosition>
</CMSearchDescription>
"@
        $recSearch = Invoke-IsapiPost -Uri "$NvrUrl/ISAPI/ContentMgmt/search" -Body $oldestXml -Cred $Cred -Timeout 8
        $firstItem = $recSearch.CMSearchResult.matchList.searchMatchItem | Select-Object -First 1

        if (-not ($firstItem -and $firstItem.timeSpan.startTime)) {
            return @{ Oldest = $oldestRec; Newest = $newestRec; Days = $retDays }
        }

        try {
            $dtOld     = ([datetime]::Parse($firstItem.timeSpan.startTime)).ToLocalTime()
            $oldestRec = $dtOld.ToString("yyyy-MM-dd HH:mm")
        } catch {
            $oldestRec = $firstItem.timeSpan.startTime
        }

        # 2. Newest clip: check last 24 hours first, then expand progressively if offline
        $nowUtc   = (Get-Date).ToUniversalTime().AddHours(2)
        $lastItem = $null

        foreach ($lookbackDays in @(1, 7, 30, 180, 730)) {
            $startUtc = $nowUtc.AddDays(-$lookbackDays)
            $lastItem = Get-LatestClipInRange -NvrUrl $NvrUrl -TrackId $TrackId -StartUtc $startUtc -EndUtc $nowUtc -Cred $Cred
            if ($lastItem) { break }
        }

        if (-not $lastItem) { $lastItem = $firstItem }

        $endRaw = if ($lastItem.timeSpan.endTime) { $lastItem.timeSpan.endTime } else { $lastItem.timeSpan.startTime }
        if ($endRaw) {
            try {
                $dtNew     = ([datetime]::Parse($endRaw)).ToLocalTime()
                $newestRec = $dtNew.ToString("yyyy-MM-dd HH:mm")
            } catch {
                $newestRec = $endRaw
            }
        }

        if ($dtOld -and $dtNew) {
            $retDays = [math]::Max(0, [math]::Round(($dtNew - $dtOld).TotalDays, 1))
        } elseif ($dtOld) {
            $retDays = [math]::Max(0, [math]::Round(((Get-Date) - $dtOld).TotalDays, 1))
        }

        return @{
            Oldest = $oldestRec
            Newest = $newestRec
            Days   = $retDays
        }
    }

    function Format-AuditSheet {
        param (
            [OfficeOpenXml.ExcelPackage]$ExcelPackage,
            [string]$SheetName,
            [int]$StartRow = 1
        )
        $ws = $ExcelPackage.Workbook.Worksheets[$SheetName]
        if (-not $ws -or -not $ws.Dimension) { return }

        $ws.View.ShowGridLines = $false

        $maxRow = $ws.Dimension.End.Row
        $maxCol = $ws.Dimension.End.Column
        if ($maxRow -lt $StartRow) { return }

        $tableRange = $ws.Cells[$StartRow, 1, $maxRow, $maxCol]
        $tableRange.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $tableRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $tableRange.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $tableRange.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

        for ($r = ($StartRow + 1); $r -le $maxRow; $r++) {
            for ($c = 1; $c -le $maxCol; $c++) {
                $val = [string]$ws.Cells[$r, $c].Text
                if ($val -match "^(Online|Normal|ok|Active|Synced)$") {
                    $ws.Cells[$r, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $ws.Cells[$r, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(198, 239, 206))
                    $ws.Cells[$r, $c].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 97, 0))
                    $ws.Cells[$r, $c].Style.Font.Bold = $true
                }
                elseif ($val -match "^(Offline|Abnormal|Error|Uninitialized|Exception|Video Loss|Failed)$") {
                    $ws.Cells[$r, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $ws.Cells[$r, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 199, 206))
                    $ws.Cells[$r, $c].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(156, 0, 6))
                    $ws.Cells[$r, $c].Style.Font.Bold = $true
                }
            }
        }
    }

    $NvrUrl  = if ($RawHost -notmatch "^https?://") { "http://$RawHost" } else { $RawHost.TrimEnd('/') }
    $CleanIP = ($RawHost -replace "^https?://", "" -replace "[:/\\]", "_")

    $DeviceInfo = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/System/deviceInfo" -Cred $Credential -Timeout 15
    if (-not $DeviceInfo) {
        return [PSCustomObject]@{
            Host       = $RawHost
            Success    = $false
            OutputFile = $null
            Error      = "Could not authenticate or reach $NvrUrl"
        }
    }

    $Channels      = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/ContentMgmt/InputProxy/channels" -Cred $Credential -Timeout 15
    $ChannelStatus = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/ContentMgmt/InputProxy/channels/status" -Cred $Credential
    $NetworkInfo   = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/System/Network/interfaces" -Cred $Credential
    $SystemStatus  = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/System/status" -Cred $Credential
    $TimeInfo      = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/System/time" -Cred $Credential
    $NtpServers    = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/System/time/ntpServers" -Cred $Credential
    $HddStatus     = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/ContentMgmt/Storage/hdd" -Cred $Credential
    $UsersXml      = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/Security/users" -Cred $Credential
    $StreamCaps    = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/Streaming/channels" -Cred $Credential

    $NvrName   = $DeviceInfo.DeviceInfo.deviceName
    $NvrModel  = $DeviceInfo.DeviceInfo.model
    $NvrSerial = $DeviceInfo.DeviceInfo.serialNumber
    $NvrFw     = $DeviceInfo.DeviceInfo.firmwareVersion
    $NvrMac    = $DeviceInfo.DeviceInfo.macAddress

    # --- 1. NVR DETAILS ---
    $PrimaryNic = $NetworkInfo.NetworkInterfaceList.NetworkInterface | Select-Object -First 1
    $NvrLocalIP = if ($PrimaryNic.IPAddress.ipAddress) { $PrimaryNic.IPAddress.ipAddress } else { $CleanIP }
    $SubnetMask = if ($PrimaryNic.IPAddress.subnetMask) { $PrimaryNic.IPAddress.subnetMask } else { "N/A" }
    $Gateway    = if ($PrimaryNic.IPAddress.DefaultGateway.ipAddress) { $PrimaryNic.IPAddress.DefaultGateway.ipAddress } else { "N/A" }
    $UptimeHrs  = if ($SystemStatus.DeviceStatus.deviceUpTime) { [math]::Round([double]$SystemStatus.DeviceStatus.deviceUpTime / 3600, 1) } else { "N/A" }
    $TimeMode   = if ($TimeInfo.Time.timeMode) { $TimeInfo.Time.timeMode } else { "N/A" }
    $LocalTime  = if ($TimeInfo.Time.localTime) { $TimeInfo.Time.localTime } else { "N/A" }
    $TimeZone   = if ($TimeInfo.Time.timeZone) { $TimeInfo.Time.timeZone } else { "N/A" }
    $NtpHost    = if ($NtpServers.NTPServerList.NTPServer.ipAddress) { ($NtpServers.NTPServerList.NTPServer.ipAddress -join ", ") } elseif ($NtpServers.NTPServerList.NTPServer.hostName) { ($NtpServers.NTPServerList.NTPServer.hostName -join ", ") } else { "N/A" }

    $NvrDetailsList = @(
        [PSCustomObject]@{
            Device_Name        = $NvrName
            Device_ID          = $DeviceInfo.DeviceInfo.deviceID
            Model              = $NvrModel
            Serial_Number      = $NvrSerial
            Firmware_Version   = $NvrFw
            Firmware_Date      = $DeviceInfo.DeviceInfo.firmwareReleasedDate
            Hardware_Version   = $DeviceInfo.DeviceInfo.hardwareVersion
            MAC_Address        = $NvrMac
            Login_Address      = $RawHost
            Local_IP_Address   = $NvrLocalIP
            Subnet_Mask        = $SubnetMask
            Default_Gateway    = $Gateway
            Uptime_Hours       = $UptimeHrs
            System_Time        = $LocalTime
            Time_Zone          = $TimeZone
            Time_Sync_Mode     = $TimeMode
            NTP_Server         = $NtpHost
            Report_Exported_By = $ExportUserName
        }
    )

    # Map Channel Online Status & Hardware
    $StatusMap = @{}
    $CamHwMap  = @{}
    if ($ChannelStatus) {
        foreach ($st in $ChannelStatus.InputProxyChannelStatusList.InputProxyChannelStatus) {
            $StatusMap[$st.id] = if ($st.online -eq "true") { "Online" } else { "Offline" }
            $CamHwMap[$st.id]  = @{
                Model    = $st.model
                Serial   = $st.serialNumber
                Firmware = $st.firmwareVersion
            }
        }
    }

    # Map Main Stream Video Settings
    $StreamMap = @{}
    if ($StreamCaps -and $StreamCaps.StreamingChannelList.StreamingChannel) {
        foreach ($sc in $StreamCaps.StreamingChannelList.StreamingChannel) {
            if ($sc.id -match "^(\d+)01$") {
                $cId = $Matches[1]
                $vid = $sc.Video
                $res = if ($vid.videoResolutionWidth -and $vid.videoResolutionHeight) { "$($vid.videoResolutionWidth)x$($vid.videoResolutionHeight)" } else { "N/A" }
                $fps = if ($vid.maxFrameRate) { [math]::Round([double]$vid.maxFrameRate / 100, 0) } else { "N/A" }
                $br  = if ($vid.constantBitRate) { $vid.constantBitRate } elseif ($vid.vbrUpperCap) { $vid.vbrUpperCap } else { "N/A" }

                $StreamMap[$cId] = @{
                    Resolution = $res
                    Codec      = if ($vid.videoCodecType) { $vid.videoCodecType } else { "N/A" }
                    FPS        = $fps
                    Bitrate    = $br
                }
            }
        }
    }

    # --- 2. CAMERAS ---
    $NvrCamerasList = @()
    foreach ($Channel in $Channels.InputProxyChannelList.InputProxyChannel) {
        $chId  = $Channel.id
        $state = if ($StatusMap.ContainsKey($chId)) { $StatusMap[$chId] } else { "Unknown" }
        $strm  = if ($StreamMap.ContainsKey($chId)) { $StreamMap[$chId] } else { @{ Resolution="N/A"; Codec="N/A"; FPS="N/A"; Bitrate="N/A" } }

        $camModel = $Channel.sourceInputPortDescriptor.model
        $camSer   = $Channel.sourceInputPortDescriptor.serialNumber
        $camFw    = $Channel.sourceInputPortDescriptor.firmwareVersion

        if (-not $camModel -and $CamHwMap.ContainsKey($chId)) { $camModel = $CamHwMap[$chId].Model }
        if (-not $camSer   -and $CamHwMap.ContainsKey($chId)) { $camSer   = $CamHwMap[$chId].Serial }
        if (-not $camFw    -and $CamHwMap.ContainsKey($chId)) { $camFw    = $CamHwMap[$chId].Firmware }

        if ($state -eq "Online" -and (-not $camModel -or -not $camSer -or -not $camFw)) {
            $proxyDev = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/ContentMgmt/InputProxy/channels/$chId/deviceInfo" -Cred $Credential -Timeout 6
            if ($proxyDev -and $proxyDev.DeviceInfo) {
                if (-not $camModel) { $camModel = $proxyDev.DeviceInfo.model }
                if (-not $camSer)   { $camSer   = $proxyDev.DeviceInfo.serialNumber }
                if (-not $camFw)    { $camFw    = $proxyDev.DeviceInfo.firmwareVersion }
            }
        }

        if ($state -eq "Online" -and (-not $camModel -or -not $camSer)) {
            $vidIn = Invoke-IsapiGet -Uri "$NvrUrl/ISAPI/System/Video/inputs/channels/$chId" -Cred $Credential -Timeout 6
            if ($vidIn -and $vidIn.VideoInputChannel) {
                if (-not $camModel) { $camModel = $vidIn.VideoInputChannel.model }
                if (-not $camSer)   { $camSer   = $vidIn.VideoInputChannel.serialNumber }
                if (-not $camFw)    { $camFw    = $vidIn.VideoInputChannel.firmwareVersion }
            }
        }

        $camModel = if ($camModel) { $camModel } else { "Not Exposed by Protocol" }
        $camSer   = if ($camSer)   { $camSer }   else { "Not Exposed by Protocol" }
        $camFw    = if ($camFw)    { $camFw }    else { "Not Exposed by Protocol" }

        # Recording Retention Span
        $tid  = ([int]$chId * 100) + 1
        $span = Get-RecordingSpan -NvrUrl $NvrUrl -TrackId $tid -Cred $Credential

        $NvrCamerasList += [PSCustomObject]@{
            NVR_Name         = $NvrName
            NVR_Model        = $NvrModel
            NVR_Serial       = $NvrSerial
            Firmware         = $NvrFw
            MAC_Address      = $NvrMac
            Channel_ID       = $chId
            Camera_Name      = $Channel.name
            Camera_IP        = $Channel.sourceInputPortDescriptor.ipAddress
            Management_Port  = $Channel.sourceInputPortDescriptor.managePortNo
            Protocol         = $Channel.sourceInputPortDescriptor.proxyProtocol
            Status           = $state
            Camera_Model     = $camModel
            Camera_Serial    = $camSer
            Camera_Firmware  = $camFw
            Resolution       = $strm.Resolution
            Codec            = $strm.Codec
            FPS              = $strm.FPS
            Bitrate_Kbps     = $strm.Bitrate
            Oldest_Recording = $span.Oldest
            Newest_Recording = $span.Newest
            Retention_Days   = $span.Days
        }
    }

    # --- 3. NVR STORAGE / HARD DRIVES ---
    $NvrStorageList = @()
    if ($HddStatus -and $HddStatus.hddList.hdd) {
        foreach ($hdd in $HddStatus.hddList.hdd) {
            $capGB  = [math]::Round([double]$hdd.capacity / 1024, 2)
            $freeGB = [math]::Round([double]$hdd.freeSpace / 1024, 2)
            $usedGB = [math]::Round(($capGB - $freeGB), 2)

            $NvrStorageList += [PSCustomObject]@{
                NVR_Name     = $NvrName
                HDD_ID       = $hdd.id
                HDD_Name     = $hdd.hddName
                Status       = $hdd.status
                Type         = $hdd.hddType
                Property     = $hdd.property
                Capacity_GB  = $capGB
                Used_GB      = $usedGB
                FreeSpace_GB = $freeGB
            }
        }
    } else {
        $NvrStorageList += [PSCustomObject]@{
            NVR_Name     = $NvrName
            HDD_ID       = "N/A"
            HDD_Name     = "No HDD data returned"
            Status       = "N/A"
            Type         = "N/A"
            Property     = "N/A"
            Capacity_GB  = 0
            Used_GB      = 0
            FreeSpace_GB = 0
        }
    }

    # --- 4. NVR USER ACCOUNTS ---
    $NvrUsersList = @()
    if ($UsersXml -and $UsersXml.UserList.User) {
        foreach ($u in $UsersXml.UserList.User) {
            $NvrUsersList += [PSCustomObject]@{
                NVR_Name      = $NvrName
                User_ID       = $u.id
                Username      = $u.userName
                Account_Level = $u.userLevel
                Bond_IP       = if ($u.bondIpAddress) { $u.bondIpAddress } else { "Any" }
                Bond_MAC      = if ($u.bondMacAddress) { $u.bondMacAddress } else { "Any" }
            }
        }
    }

    # --- 5. EXPORT SEPARATE EXCEL WORKBOOK FOR THIS NVR ---
    $Timestamp  = Get-Date -Format "MMM-dd-yyyy_HH_mm_ss"
    $OutputFile = Join-Path -Path $DesktopPath -ChildPath "NVR_login_${CleanIP}_${Timestamp}.xlsx"

    $OnlineCount  = ($NvrCamerasList | Where-Object { $_.Status -eq "Online" }).Count
    $OfflineCount = ($NvrCamerasList | Where-Object { $_.Status -ne "Online" }).Count
    $TotalCap     = ($NvrStorageList | Measure-Object -Property Capacity_GB -Sum).Sum
    $TotalUsed    = ($NvrStorageList | Measure-Object -Property Used_GB -Sum).Sum
    $UsedPct      = if ($TotalCap -gt 0) { [math]::Round(($TotalUsed / $TotalCap) * 100, 1) } else { 0 }

    # Sheet 1: NVR_Details
    $ExcelPkg = $NvrDetailsList | Export-Excel -Path $OutputFile -WorksheetName "NVR_Details" -StartRow 5 -TableStyle "Medium2" -AutoSize -PassThru
    $wsNvr = $ExcelPkg.Workbook.Worksheets["NVR_Details"]

    $wsNvr.Cells["A1"].Value = "NVR AUDIT & HEALTH REPORT - $NvrName ($CleanIP)"
    $wsNvr.Cells["A1:H1"].Merge = $true
    $wsNvr.Cells["A1"].Style.Font.Size = 14
    $wsNvr.Cells["A1"].Style.Font.Bold = $true
    $wsNvr.Cells["A1"].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(31, 78, 120))

    $wsNvr.Cells["A2"].Value = "Report Exported By: $ExportUserName   |   Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')   |   Cameras: $OnlineCount Online / $OfflineCount Offline   |   Storage Used: $TotalUsed GB / $TotalCap GB ($UsedPct%)"
    $wsNvr.Cells["A2:L2"].Merge = $true
    $wsNvr.Cells["A2"].Style.Font.Bold = $true
    $wsNvr.Cells["A2"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $wsNvr.Cells["A2"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217, 225, 242))

    Format-AuditSheet -ExcelPackage $ExcelPkg -SheetName "NVR_Details" -StartRow 5
    Close-ExcelPackage $ExcelPkg

    # Sheet 2: Cameras
    $ExcelPkg = $NvrCamerasList | Export-Excel -Path $OutputFile -WorksheetName "Cameras" -TableStyle "Medium2" -AutoSize -FreezeTopRow -PassThru
    Format-AuditSheet -ExcelPackage $ExcelPkg -SheetName "Cameras" -StartRow 1
    Close-ExcelPackage $ExcelPkg

    # Sheet 3: NVR_Storage
    $ExcelPkg = $NvrStorageList | Export-Excel -Path $OutputFile -WorksheetName "NVR_Storage" -TableStyle "Medium2" -AutoSize -FreezeTopRow -PassThru
    Format-AuditSheet -ExcelPackage $ExcelPkg -SheetName "NVR_Storage" -StartRow 1
    Close-ExcelPackage $ExcelPkg

    # Sheet 4: Users
    if ($NvrUsersList.Count -gt 0) {
        $ExcelPkg = $NvrUsersList | Export-Excel -Path $OutputFile -WorksheetName "Users" -TableStyle "Medium2" -AutoSize -FreezeTopRow -PassThru
        Format-AuditSheet -ExcelPackage $ExcelPkg -SheetName "Users" -StartRow 1
        Close-ExcelPackage $ExcelPkg
    }

    return [PSCustomObject]@{
        Host       = $RawHost
        Success    = $true
        OutputFile = $OutputFile
        Error      = $null
    }
}

$Jobs = @()
$ExportedFiles = @()

try {
    Write-Host "`nLaunching $($HostList.Count) simultaneous NVR extraction job(s)..." -ForegroundColor Cyan

    # Launch all NVRs in parallel at the exact same time
    foreach ($RawHost in $HostList) {
        Write-Host "  [+] Started parallel worker for: $RawHost" -ForegroundColor Gray
        $Jobs += Start-Job -ScriptBlock $NvrWorkerScript -ArgumentList $RawHost, $ExportUserName, $Credential, $DesktopPath
    }

    Write-Host "`nWaiting for all NVRs to complete simultaneously..." -ForegroundColor Yellow
    $Jobs | Wait-Job -Timeout 120 | Out-Null

    # Collect results from all parallel workers
    foreach ($Job in $Jobs) {
        if ($Job.State -eq "Running") {
            Stop-Job -Job $Job -ErrorAction SilentlyContinue
            Write-Host "  [TIMEOUT] Worker exceeded 120s and was stopped." -ForegroundColor Red
            Remove-Job -Job $Job -Force
            continue
        }

        $Result = Receive-Job -Job $Job
        Remove-Job -Job $Job -Force

        if ($Result -and $Result.Success) {
            $ExportedFiles += $Result.OutputFile
            Write-Host "  [OK] $($Result.Host) -> $($Result.OutputFile)" -ForegroundColor Green
            Invoke-Item $Result.OutputFile
        } else {
            $errHost = if ($Result.Host) { $Result.Host } else { "Unknown NVR" }
            $errMsg  = if ($Result.Error) { $Result.Error } else { "Job failed unexpectedly" }
            Write-Host "  [FAILED] $errHost : $errMsg" -ForegroundColor Red
        }
    }

    if ($ExportedFiles.Count -eq 0) {
        throw "No NVR devices responded successfully. Check IP, port, or credentials."
    }

    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host " ALL REPORTS EXPORTED SIMULTANEOUSLY" -ForegroundColor Green
    Write-Host " Report Exported By : $ExportUserName" -ForegroundColor Cyan
    Write-Host " Total NVR Files    : $($ExportedFiles.Count)" -ForegroundColor White
    foreach ($f in $ExportedFiles) {
        Write-Host "   -> $f" -ForegroundColor White
    }
    Write-Host "============================================================" -ForegroundColor Green

} catch {
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Clean up any leftover jobs & purge sensitive credentials from memory
    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
    if ($SecurePass) { $SecurePass.Dispose() }
    Remove-Variable -Name ExportUserName, User, SecurePass, Credential -ErrorAction SilentlyContinue
    Read-Host "`nPress Enter to close this window"
}