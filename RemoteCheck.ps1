# ------------------------------
# RemoteCheck.ps1
# 功能: 查詢遠端主機名稱、埠狀態與進程資訊，並以表格方式呈現
# ------------------------------

### 1. 檢查與設置執行策略 ###
$originalPolicy = Get-ExecutionPolicy
Write-Host "Current Execution Policy: $originalPolicy"

Write-Host "Temporarily setting Execution Policy to Bypass..."
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

### 2. 定義目標電腦與埠 ###
$targetsFile = "C:\Users\2019051401\Desktop\targets.txt"
$portsFile = "C:\Users\2019051401\Desktop\ports.txt"

# 讀取目標電腦
if (Test-Path $targetsFile) {
    Write-Host "Reading targets from file: $targetsFile"
    $args = Get-Content -Path $targetsFile
} else {
    Write-Host "Target file not found. Using default targets."
    $args = "10.231.110.1", "10.231.250.21"
}

# 讀取埠列表
if (Test-Path $portsFile) {
    Write-Host "Reading ports from file: $portsFile"
    $ports = Get-Content -Path $portsFile
} else {
    Write-Host "Ports file not found. Using default port: 3389"
    $ports = 3389
}

# 初始化結果表格
$results = @()

### 3. 遍歷每個目標 ###
foreach ($target in $args) {
    Write-Host "`nProcessing target: $target" -ForegroundColor Cyan

    # 嘗試通過 WMI 獲取主機名稱
    try {
        $hostname = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName $target).Name
    } catch {
        $hostname = "N/A"
        Write-Host "Unable to retrieve host name for $target" -ForegroundColor Red
    }

    # 檢查指定埠的進程資訊
    foreach ($port in $ports) {
        try {
            # 嘗試檢查進程
            $processDetails = Get-WmiObject -Class Win32_Process -Filter "Name LIKE '%'" -ComputerName $target | Where-Object {
                $_.CommandLine -match "${port}"
            }

            if ($processDetails) {
                foreach ($process in $processDetails) {
                    # 將結果新增到表格
                    $results += [PSCustomObject]@{
                        HostName      = $hostname
                        IPAddress     = $target
                        Port          = $port
                        ProcessName   = $process.Name
                        ProcessPath   = $process.ExecutablePath
                        PID           = $process.ProcessId
                    }
                }
            } else {
                # 如果無匹配進程，記錄空白
                $results += [PSCustomObject]@{
                    HostName      = $hostname
                    IPAddress     = $target
                    Port          = $port
                    ProcessName   = "No process"
                    ProcessPath   = "N/A"
                    PID           = "N/A"
                }
            }
        } catch {
            $results += [PSCustomObject]@{
                HostName      = $hostname
                IPAddress     = $target
                Port          = $port
                ProcessName   = "Error retrieving"
                ProcessPath   = "N/A"
                PID           = "N/A"
            }
        }
    }
}

### 4. 輸出表格 ###
Write-Host "`nFinal Results:" -ForegroundColor Green
$results | Format-Table -AutoSize

### 5. 保存結果到 CSV ###
$outputFile = "C:\Users\2019051401\Desktop\RemoteCheckResults.csv"
$results | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
Write-Host "`nResults have been saved to $outputFile"

### 6. 恢復執行策略 ###
Write-Host "`nRestoring original Execution Policy: $originalPolicy"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy $originalPolicy -Force

Write-Host "Script execution complete."
