@powershell "$s=[scriptblock]::create((gc %~f0|?{$_.readcount -gt 1})-join\"`n\");&$s" %*&goto:eof
Add-Type -Assembly System.Windows.Forms
#################################################################
# ドラッグ&ドロップでpowershellスクリプトのショートカットを作成する。
#  (Windows PowerShell用)
#################################################################
# ↓ここからPowerShellスクリプト
if ($args.Length -eq 0) {
    Write-Warning "PowerShellスクリプトをドラッグ＆ドロップして起動してください。"
    pause
    exit 1
}

if ((Test-Path $args[0] -PathType Leaf) -eq $false) {
    Write-Warning "引数のスクリプトが存在しません。"
    pause
    exit 1
}

$scriptPath = Resolve-Path -LiteralPath $args[0]

if ((Get-Item -LiteralPath $scriptPath).Extension -ne ".ps1") {
    Write-Warning "引数のスクリプトはPowerShellスクリプト(拡張子:.ps1)ではありません。"
    pause
    exit 1
}

$outputDirPath = Split-Path -Parent $scriptPath
$scriptBaseName = (Get-Item $scriptPath).BaseName
$outputShortcutPath = Join-Path -Path $outputDirPath -ChildPath "${scriptBaseName}.lnk"

# ショートカットの作成
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($outputShortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Unrestricted -File ${scriptPath}"
$shortcut.IconLocation = "$PSHOME\powershell.exe"
$shortcut.WorkingDirectory = (Split-Path -Parent $scriptPath )
$shortcut.WindowStyle = 7

$shortcut.Save()

pause