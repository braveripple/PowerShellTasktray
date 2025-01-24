param (
    [string] $USER,
    [string] $PASSWORD,
    [string] $DB_NAME,
    [string] $HOST_NAME = "localhost",
    [string] $PORT_NO = "5432"
)

add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms

$PSQLEDIT_PATH = "C:\app\psqledit_5_32\psqledit.exe"

$now = Get-Date

$REG_PATH = "HKCU:\Software\OGAWA\POSTGRESQL\PSqlEdit\CONNECT_INFO\"

$preUser = Get-ItemPropertyValue $REG_PATH -Name "USER-CUR"
$preDBName = Get-ItemPropertyValue $REG_PATH -Name "DBNAME-CUR"
$preHostName = Get-ItemPropertyValue $REG_PATH -Name "HOST-CUR"
$prePortNo = Get-ItemPropertyValue $REG_PATH -Name "PORT-CUR"
$prePassword = Get-ItemPropertyValue $REG_PATH -Name "PASSWD-0"

Write-Debug "ユーザ名：$preUser"
Write-Debug "DB名：$preDBName"
Write-Debug "ホスト名：$preHostName"
Write-Debug "ポート番号：$prePortNo"

Set-ItemProperty $REG_PATH -Name "USER-CUR" -Value $USER
Set-ItemProperty $REG_PATH -Name "DBNAME-CUR" -Value $DB_NAME
Set-ItemProperty $REG_PATH -Name "HOST-CUR" -Value $HOST_NAME
Set-ItemProperty $REG_PATH -Name "PORT-CUR" -Value $PORT_NO
Set-ItemProperty $REG_PATH -Name "PASSWD-0" -Value ""

# アプリケーションを起動
Start-Process -FilePath $PSQLEDIT_PATH

# 対象プロセスが見つかるまで待機
do {
    $processes = Get-Process psqledit | Where-Object {
        $_.StartTime -ge $now -and $_.Path -eq $PSQLEDIT_PATH
    }
    Start-Sleep 0.5
} while ($null -eq $processes)

# 最新のPSqlEditのプロセスを取得
$p = $processes | Sort-Object -Property Id -Descending | Select-Object -First 1

# PSqlEditを最前面化＆キー送信
$success = $false
$retryCount = 1
$maxRetries = 10

Start-Sleep 0.5

while (-not $success -and $retryCount -lt $maxRetries) {
    try {
        [Microsoft.VisualBasic.Interaction]::AppActivate($p.Id)
        $success = $true
        Write-Debug "最前面化に成功しました。"
    } catch {
        Write-Debug "最前面化に失敗しました。再試行します... ($retryCount/$maxRetries)"
        Start-Sleep 0.5
        $retryCount++
    }
}

# 最前面に来た段階で環境変数を元に戻す
Set-ItemProperty $REG_PATH -Name "USER-CUR" -Value $preUser
Set-ItemProperty $REG_PATH -Name "DBNAME-CUR" -Value $preDBName
Set-ItemProperty $REG_PATH -Name "HOST-CUR" -Value $preHostName
Set-ItemProperty $REG_PATH -Name "PORT-CUR" -Value $prePortNo
Set-ItemProperty $REG_PATH -Name "PASSWD-0" -Value $prePassword

if ($success) {
    [System.Windows.Forms.SendKeys]::SendWait($PASSWORD)
    Write-Debug "PASSWORDを送信しました。"
    
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Write-Debug "ENTERキーを送信しました。"
}
