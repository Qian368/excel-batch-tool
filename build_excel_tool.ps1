<#
.SYNOPSIS
  自动编译Excel批量操作工具的Inno Setup脚本

.DESCRIPTION
  该脚本用于自动调用ISCC.exe编译excel_tool.iss安装程序脚本，包含错误处理和日志记录功能。

.PARAMETER ISCCPath
  Inno Setup编译器路径，默认为D:\Program Files (x86)\Inno Setup 6\ISCC.exe

.PARAMETER ISSScriptPath
  ISS脚本文件路径，默认为当前目录下的excel_tool.iss

.EXAMPLE
  .\build_excel_tool.ps1
  使用默认路径编译ISS脚本

.EXAMPLE
  .\build_excel_tool.ps1 -ISCCPath "C:\Program Files\Inno Setup 6\ISCC.exe"
  指定ISCC.exe路径编译ISS脚本
#>

param(
    [string]$ISCCPath = "D:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    [string]$ISSScriptPath = "$PSScriptRoot\excel_tool.iss"
)

# 初始化日志文件
$logFile = "$PSScriptRoot\build_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$timestamp] $message" | Out-File -FilePath $logFile -Append
    Write-Host "[$timestamp] $message"
}

# 检查ISCC.exe是否存在
if (-not (Test-Path $ISCCPath)) {
    $errorMsg = "错误: 找不到ISCC.exe，请确认Inno Setup已安装且路径正确。当前路径: $ISCCPath"
    Write-Log $errorMsg
    throw $errorMsg
}

# 检查ISS脚本是否存在
if (-not (Test-Path $ISSScriptPath)) {
    $errorMsg = "错误: 找不到ISS脚本文件，请确认路径正确。当前路径: $ISSScriptPath"
    Write-Log $errorMsg
    throw $errorMsg
}

# 新增编码检查函数
function Check-FileEncoding {
    param([string]$filePath)
    $content = Get-Content -Path $filePath -Raw -Encoding Byte
    # 检查UTF-8 BOM
    if ($content[0] -eq 0xEF -and $content[1] -eq 0xBB -and $content[2] -eq 0xBF) {
        return "UTF8"
    }
    return "ANSI"
}

# 开始编译
Write-Log "开始编译ISS脚本: $ISSScriptPath"
Write-Log "使用ISCC路径: $ISCCPath"

try {
    # 设置控制台输出编码为UTF-8
    $OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = [System.Text.Encoding]::UTF8
    
    # 调用ISCC.exe编译ISS脚本
    $process = Start-Process -FilePath $ISCCPath -ArgumentList "`"$ISSScriptPath`"" -NoNewWindow -Wait -PassThru
    
    if ($process.ExitCode -eq 0) {
        Write-Log "编译成功完成!"
    }
    else {
        $errorMsg = "编译失败，退出代码: $($process.ExitCode)"
        Write-Log $errorMsg
        throw $errorMsg
    }
}
catch {
    $errorMsg = "编译过程中发生错误: $_"
    Write-Log $errorMsg
    throw $errorMsg
}

Write-Log "编译流程结束"

# 在编译前添加编码检查
$encoding = Check-FileEncoding -filePath $ISSScriptPath
if ($encoding -ne "ANSI") {
    Write-Log "警告: ISS文件编码为$encoding，正在转换为ANSI..."
    $content = Get-Content -Path $ISSScriptPath -Raw
    $content | Out-File -FilePath $ISSScriptPath -Encoding Default -Force
    Write-Log "编码转换完成"
}