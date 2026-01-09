# encoding: utf-8 with BOM
# mouse_wheel_scroll_reverse.ps1
# DOS风格鼠标滚轮方向配置工具

# 确保PowerShell以正确的编码读取文件
if ([System.Text.Encoding]::Default.BodyName -ne 'utf-8') {
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}
# Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# chcp 437 | Out-Null
# # 设置控制台输入输出为 DOS 代码页
# [Console]::InputEncoding = [System.Text.Encoding]::Default
# [Console]::OutputEncoding = [System.Text.Encoding]::Default

# 读取外部配置文件
$configFilePath = "vendors.json"

try {
    if (Test-Path $configFilePath) {
        $config = Get-Content -Path $configFilePath -Raw | ConvertFrom-Json
        $VendorMap = @{}
        $ProductMap = @{}

        # 将PSObject转换为哈希表
        $config.vendors.PSObject.Properties | ForEach-Object {
            $VendorMap[$_.Name] = $_.Value
        }

        $config.products.PSObject.Properties | ForEach-Object {
            $ProductMap[$_.Name] = $_.Value
        }
    } else {
        Write-Host "❌ CONFIG FILE NOT FOUND: $configFilePath" -ForegroundColor Red
        Write-Host "📁 CREATING DEFAULT CONFIG FILE..." -ForegroundColor Yellow

        # 创建默认配置
        $defaultConfig = @{
            vendors = @{
                "046D" = "LOGITECH"
                "045E" = "MICROSOFT"
                "093A" = "PIXART"
                "1532" = "RAZER"
                "1038" = "STEELSERIES"
                "1B1C" = "CORSAIR"
                "258A" = "RAPOO"
                "04D9" = "HOLTEK"
                "2734" = "DELL"
                "256C" = "PERIXX"
            }
            products = @{
                "046D_C52B" = "UNIFYING RECEIVER"
                "046D_C08B" = "G502 HERO GAMING MOUSE"
                "046D_C092" = "M325 WIRELESS MOUSE"
                "046D_C332" = "G502 PROTEUS SPECTRUM"
                "046D_C247" = "G100S OPTICAL GAMING"
                "045E_079B" = "SCULPT ERGONOMIC"
                "1532_006F" = "DEATHADDER ESSENTIAL"
                "1038_1824" = "RIVAL 310"
                "258A_001F" = "VT950 PRO"
            }
        }

        $defaultConfig | ConvertTo-Json -Depth 3 | Out-File -FilePath $configFilePath -Encoding UTF8
        Write-Host "✅ DEFAULT CONFIG FILE CREATED SUCCESSFULLY!" -ForegroundColor Green

        # 使用默认配置
        $VendorMap = $defaultConfig.vendors
        $ProductMap = $defaultConfig.products
    }
} catch {
    Write-Host "❌ ERROR READING CONFIG FILE: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "🔧 USING DEFAULT VALUES..." -ForegroundColor Yellow

    # 使用默认值
    # ==== 厂商数据库 ====
    $VendorMap = @{
        "046D" = "LOGITECH"
        "045E" = "MICROSOFT"
        "093A" = "PIXART"
        "1532" = "RAZER"
        "1038" = "STEELSERIES"
        "1B1C" = "CORSAIR"
        "258A" = "RAPOO"
        "04D9" = "HOLTEK"
        "2734" = "DELL"
        "256C" = "PERIXX"
    }

    # ==== 产品型号数据库 ====
    $ProductMap = @{
        "046D_C52B" = "UNIFYING RECEIVER"
        "046D_C08B" = "G502 HERO GAMING MOUSE"
        "046D_C092" = "M325 WIRELESS MOUSE"
        "046D_C332" = "G502 PROTEUS SPECTRUM"
        "046D_C247" = "G100S OPTICAL GAMING"
        "045E_079B" = "SCULPT ERGONOMIC"
        "1532_006F" = "DEATHADDER ESSENTIAL"
        "1038_1824" = "RIVAL 310"
        "258A_001F" = "VT950 PRO"
    }
}


function Get-VendorName {
    param([string]$vid)
    if ($VendorMap.ContainsKey($vid)) {
        return $VendorMap[$vid]
    }
    return "UNKNOWN"
}

function Get-ProductName {
    param([string]$vid, [string]$productId)
    $key = "${vid}_${productId}".ToUpper()
    if ($ProductMap.ContainsKey($key)) {
        return $ProductMap[$key]
    }

    $vendor = Get-VendorName $vid
    if ($vendor -eq "UNKNOWN") {
        return "HID MOUSE"
    } else {
        return "$vendor MOUSE"
    }
}

function Clear-Screen {
    Clear-Host
}

function Show-Header {
    Write-Host "╔══════════════════════════════════════════════════════════════════════════╗" -ForegroundColor White
    Write-Host "║                        MOUSE WHEEL SCROLL REVERSE                        ║" -ForegroundColor White
    Write-Host "║                             DOS EDITION  v0.1                            ║" -ForegroundColor White
    Write-Host "╚══════════════════════════════════════════════════════════════════════════╝" -ForegroundColor White
    Write-Host ""
}

function Show-MouseList {
    param([ref]$mouseListRef, [string]$errorMessage = $null)

    Clear-Screen
    Show-Header

    if ($errorMessage) {
        Write-Host "ERROR: $errorMessage" -ForegroundColor Red
        Write-Host ""
    }

    Write-Host "SCANNING MOUSE DEVICES... 📡" -ForegroundColor Yellow

    $mouseDevices = Get-PnpDevice -Class Mouse -ErrorAction SilentlyContinue |
                    Where-Object { $_.Status -eq 'OK' -and $_.InstanceId -like 'HID\VID_*&PID_*' }

    if ($mouseDevices.Count -eq 0) {
        Write-Host ""
        Write-Host "❌ ERROR: NO ACTIVE MOUSE DEVICES FOUND." -ForegroundColor Red
        Write-Host ""
        Write-Host "PRESS ANY KEY TO EXIT... " -ForegroundColor Gray -NoNewline
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return $false
    }

    $mouseList = @()
    $index = 1
    foreach ($device in $mouseDevices) {
        if ($device.InstanceId -match 'VID_([0-9A-F]{4})&PID_([0-9A-F]{4})') {
            $vid = $matches[1].ToUpper()
            $productId = $matches[2].ToUpper()

            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($device.InstanceId)\Device Parameters"
            $flipProp = Get-ItemProperty -Path $regPath -Name "FlipFlopWheel" -ErrorAction SilentlyContinue
            $isInverted = if ($flipProp -and $flipProp.FlipFlopWheel -eq 1) { $true } else { $false }

            $mouseList += [PSCustomObject]@{
                ID = $index
                Vendor = Get-VendorName $vid
                Product = Get-ProductName $vid $productId
                HardwareId = "VID_$vid&PID_$productId"
                InstanceId = $device.InstanceId
                RegPath = $regPath
                IsInverted = $isInverted
                FriendlyName = $device.FriendlyName
            }
            $index++
        }
    }

    # 显示表格
    Write-Host ""
    Write-Host "🔍 FOUND $($mouseList.Count) MOUSE DEVICE(S):" -ForegroundColor Green
    Write-Host ""
    Write-Host "  ID  VENDOR             PRODUCT                    HWID                 INVERT" -ForegroundColor White
    Write-Host "  --  ------             -------                    ----                 ------" -ForegroundColor DarkGray

    foreach ($mouse in $mouseList) {
        $invertStr = if ($mouse.IsInverted) { "YES" } else { "NO " }
        $idStr = $mouse.ID.ToString().PadRight(3)
        $vendorStr = $mouse.Vendor.PadRight(18)
        $productStr = $mouse.Product.PadRight(26)
        $hwidStr = $mouse.HardwareId.PadRight(20)

        Write-Host "  $idStr $vendorStr $productStr $hwidStr $invertStr" -ForegroundColor White
    }

    Write-Host ""
    Write-Host "📋 OPTIONS:" -ForegroundColor Cyan
    Write-Host "  • ENTER MOUSE ID TO CONFIGURE" -ForegroundColor White
    Write-Host "  • PRESS [ENTER] TO REFRESH" -ForegroundColor White
    Write-Host "  • TYPE 'Q' TO EXIT" -ForegroundColor White
    Write-Host ""
    Write-Host "SELECT MOUSE ID (OR PRESS ENTER/TYPE Q): " -ForegroundColor Yellow -NoNewline

    $mouseListRef.Value = $mouseList
    return $true
}

function Set-FlipFlopWheel {
    param(
        [string]$regPath,
        [bool]$invert
    )

    $value = if ($invert) { 1 } else { 0 }

    try {
        if (-not (Test-Path $regPath)) {
            return $false
        }

        Set-ItemProperty -Path $regPath -Name "FlipFlopWheel" -Value $value -Type DWord -Force
        return $true
    }
    catch {
        return $false
    }
}

function Show-ResultScreen {
    param(
        [object]$selectedMouse,
        [bool]$targetInvert,
        [bool]$success
    )

    Clear-Screen
    Show-Header

    if ($success) {
        Write-Host "✅ OPERATION SUCCESSFUL!" -ForegroundColor Green
        Write-Host ""
        Write-Host "🖱️ DEVICE: $($selectedMouse.Vendor) $($selectedMouse.Product)" -ForegroundColor White
        Write-Host "🔧 HARDWARE ID: $($selectedMouse.HardwareId)" -ForegroundColor White
        Write-Host "⚙️ NEW STATE: $(if ($targetInvert) { 'INVERTED' } else { 'NORMAL' })" -ForegroundColor White
        Write-Host ""
        Write-Host "🔄 TO APPLY CHANGES:" -ForegroundColor Yellow
        Write-Host "   • UNPLUG AND RECONNECT MOUSE" -ForegroundColor White
        Write-Host "   • OR RESTART COMPUTER" -ForegroundColor White
        Write-Host ""
        Write-Host "🛡️ ADMIN RIGHTS: $(if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { 'YES' } else { 'NO' })" -ForegroundColor Cyan
    } else {
        Write-Host "❌ OPERATION FAILED!" -ForegroundColor Red
        Write-Host ""
        Write-Host "🖱️ DEVICE: $($selectedMouse.Vendor) $($selectedMouse.Product)" -ForegroundColor White
        Write-Host "🔧 HARDWARE ID: $($selectedMouse.HardwareId)" -ForegroundColor White
        Write-Host "⚙️ TARGET STATE: $(if ($targetInvert) { 'INVERTED' } else { 'NORMAL' })" -ForegroundColor White
        Write-Host ""
        Write-Host "⚠️ ERROR: ACCESS DENIED OR REGISTRY ERROR" -ForegroundColor Red
        Write-Host ""
        Write-Host "💡 SOLUTION: RUN AS ADMINISTRATOR" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "🔄 PRESS ANY KEY TO RETURN TO MAIN MENU... " -ForegroundColor White -NoNewline
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ====== 主循环 ======
do {
    # 获取鼠标列表
    $mouseList = $null
    if (-not (Show-MouseList -mouseListRef ([ref]$mouseList))) {
        break
    }

    $choice = Read-Host

    # 检查是否退出
    if ($choice -eq 'Q' -or $choice -eq 'q' -or $choice -eq 'QUIT' -or $choice -eq 'quit') {
        break
    }

    # 检查是否为空（刷新）
    if ($choice -eq '') {
        continue
    }

    # 检查输入是否为数字
    if (-not ($choice -match '^\d+$')) {
        # 显示错误并重新显示列表
        Show-MouseList -mouseListRef ([ref]$mouseList) -errorMessage "INVALID INPUT: PLEASE ENTER A VALID MOUSE ID NUMBER"
        continue
    }

    $selectedId = [int]$choice
    $selectedMouse = $mouseList | Where-Object { $_.ID -eq $selectedId }

    if ($null -eq $selectedMouse) {
        # 显示错误并重新显示列表
        Write-Host ""
        Write-Host "❌ INVALID MOUSE ID $selectedId" -ForegroundColor Red
        Write-Host "🔄 PRESS ANY KEY TO TRY AGAIN... " -ForegroundColor Yellow -NoNewline
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        continue
    }

    # 滚轮方向选择循环
    do {
        Clear-Screen
        Show-Header

        Write-Host "🎯 SELECTED DEVICE:" -ForegroundColor Cyan
        Write-Host "   $($selectedMouse.Vendor) $($selectedMouse.Product)" -ForegroundColor White
        Write-Host "   $($selectedMouse.HardwareId)" -ForegroundColor White
        Write-Host "   CURRENT: $(if ($selectedMouse.IsInverted) { 'INVERTED' } else { 'NORMAL' })" -ForegroundColor White
        Write-Host ""
        Write-Host "🖱 SELECT WHEEL DIRECTION:" -ForegroundColor Yellow
        Write-Host "   1) 🔄 NORMAL (DEFAULT)" -ForegroundColor White
        Write-Host "   2) 🔄 INVERTED (NATURAL SCROLL)" -ForegroundColor White
        Write-Host "   3) 🏠 RETURN TO MAIN MENU" -ForegroundColor Gray
        Write-Host ""
        Write-Host "ENTER CHOICE (1/2/3): " -ForegroundColor Yellow -NoNewline

        $directionChoice = Read-Host
            # 检查输入是否为数字
            if (-not ($directionChoice -match '^[1-3]$')) {
                Write-Host ""
                Write-Host "❌ INVALID INPUT: PLEASE ENTER 1, 2, OR 3" -ForegroundColor Red
                Write-Host "🔄 PRESS ANY KEY TO TRY AGAIN... " -ForegroundColor Yellow -NoNewline
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                continue
            }

        switch ($directionChoice) {
            "1" {
                $targetInvert = $false
            }
            "2" {
                $targetInvert = $true
            }
            "3" {
                    break  # 跳出方向选择循环，返回主菜单
            }
            }

            # 如果选择了返回，继续主循环
        if ($directionChoice -eq "3") {
                break
        }

        # 执行修改
        $success = Set-FlipFlopWheel -regPath $selectedMouse.RegPath -invert $targetInvert

        # 显示结果并等待返回
        Show-ResultScreen -selectedMouse $selectedMouse -targetInvert $targetInvert -success $success
            break  # 完成操作后返回主菜单

    } while ($true)
} while ($true)

Write-Host ""
Write-Host "👋 GOODBYE!" -ForegroundColor Green
