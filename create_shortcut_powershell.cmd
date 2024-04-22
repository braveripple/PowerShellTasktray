@powershell "$s=[scriptblock]::create((gc %~f0|?{$_.readcount -gt 1})-join\"`n\");&$s" %*&goto:eof
Add-Type -Assembly System.Windows.Forms
#################################################################
# �h���b�O&�h���b�v��powershell�X�N���v�g�̃V���[�g�J�b�g���쐬����B
#  (Windows PowerShell�p)
#################################################################
# ����������PowerShell�X�N���v�g
if ($args.Length -eq 0) {
    Write-Warning "PowerShell�X�N���v�g���h���b�O���h���b�v���ċN�����Ă��������B"
    pause
    exit 1
}

if ((Test-Path $args[0] -PathType Leaf) -eq $false) {
    Write-Warning "�����̃X�N���v�g�����݂��܂���B"
    pause
    exit 1
}

$scriptPath = Resolve-Path -LiteralPath $args[0]

if ((Get-Item -LiteralPath $scriptPath).Extension -ne ".ps1") {
    Write-Warning "�����̃X�N���v�g��PowerShell�X�N���v�g(�g���q:.ps1)�ł͂���܂���B"
    pause
    exit 1
}

$outputDirPath = Split-Path -Parent $scriptPath
$scriptBaseName = (Get-Item $scriptPath).BaseName
$outputShortcutPath = Join-Path -Path $outputDirPath -ChildPath "${scriptBaseName}.lnk"

# �V���[�g�J�b�g�̍쐬
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($outputShortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Unrestricted -File ${scriptPath}"
$shortcut.IconLocation = "$PSHOME\powershell.exe"
$shortcut.WorkingDirectory = (Split-Path -Parent $scriptPath )
$shortcut.WindowStyle = 7

$shortcut.Save()

pause