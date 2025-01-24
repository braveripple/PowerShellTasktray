using namespace System.Management.Automation
param (
    [string] $DBConfigFile
)
<###############################################################
PSqlEditランチャー
###############################################################>

Add-Type -AssemblyName System.Windows.Forms;

$APPLICATION_NAME = "PSqlEdit ランチャー"
$MUTEX_NAME = 'CE634DBE-31E2-4E65-838A-11CF22BDBBC4' + $APPLICATION_NAME;

if ([string]::IsNullOrWhiteSpace($DBConfigFile)) {
    $DBConfigFile = "$PSScriptRoot/接続先.tsv"
}
if (!(Test-Path -LiteralPath $DBConfigFile -PathType Leaf)) {
    Write-Error "接続先設定ファイルが存在しません：$DBConfigFile"
    exit 1
}

$EDITOR_PATH = "notepad.exe"
$mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME);

# 接続先の選択画面を開く
function displayAccessPoint {
    try {
        # JSONまたはTSVファイルから接続先情報を取得
        if ((Get-Item -LiteralPath $DBConfigFile).Extension.ToLower() -eq ".json") {
            $dbList = Get-Content -LiteralPath $DBConfigFile -Encoding utf8 | ConvertFrom-Json
        } else {
            $dbList = Import-Csv -LiteralPath $DBConfigFile -Delimiter "`t" -Encoding utf8;
        }
        
        if ($dbList.Length -gt 0) {
            $columns = $dbList[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
            $i = 1            
            foreach ($db in $dbList) {
                # 環境変数を値に変換する
                foreach ($column in $columns) {
                    $db."$column" = [regex]::Replace($db."$column", '\${?Env:([^}]+)}?', { $(get-item "Env:$($args.Groups[1].Value)").Value })
                }
                # ID列を追加する
                $db | Add-Member -Name Id -MemberType NoteProperty -Value ("{0:00}" -f $i++)
            }
        }
        # 選択画面表示。パスワードは画面に表示しない。
        $select = $dbList | 
            Select-Object Id, User, DBName, Host, PortNo, Memo | 
            Out-GridView -OutputMode Multiple -Title "${APPLICATION_NAME} 接続先の選択"

        $select | ForEach-Object {
            $tmp = $_
            Write-Debug $tmp
            $db = $dbList | 
                Where-Object { $_.Id -eq $tmp.Id } | Select-Object -First 1
            Write-Debug $db

            .\PSqlEditLoginW.ps1 -USER $db.User -PASSWORD $db.Password -DB_NAME $db.DBName -HOST_NAME $db.Host -PORT_NO $db.PortNo

            Start-Sleep 1
        }
    }
    catch {
        Write-Error $_.ToString();
        $notifyIcon.BalloonTipText = $_.ToString();
        $notifyIcon.ShowBalloonTip(5000);
    }

}
function displayDBConfigFile {
    try {
        & $EDITOR_PATH $DBConfigFile
    }
    catch {
        $notifyIcon.BalloonTipText = $_.ToString();
        $notifyIcon.ShowBalloonTip(5000);
    }
}

function openScriptDir {
    try {
        Invoke-Item $PSScriptRoot
    }
    catch {
        $notifyIcon.BalloonTipText = $_.ToString();
        $notifyIcon.ShowBalloonTip(5000);
    }
}

function displayTooltip {
    try {
        Write-Debug "  make context start"
        # コンテキスト作成
        $appContext = New-Object System.Windows.Forms.ApplicationContext;
        Write-Debug "  make context end"
        ####################################################
        # 通知アイコン作成
        ####################################################
        Write-Debug "  set icon start"
        # TeraTermのアイコンを設定する
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell.exe).Source)
        $notifyIcon = [System.Windows.Forms.NotifyIcon]@{
            Icon           = $icon;
            Text           = $APPLICATION_NAME;
            BalloonTipIcon = 'None';
        };
        Write-Debug "  set icon end"

        ####################################################
        # アイコン左クリック時のイベントを設定
        ####################################################
        Write-Debug "  set left click event start"
        $notifyIcon.add_Click( {
                if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    displayAccessPoint
                }
            });
        Write-Debug "  set left click event end"
        ####################################################
        # アイコン右クリック時のコンテキストメニューの設定
        ####################################################
        Write-Debug "  set right click event start"
        $menuItem_exit = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'Exit' };
        $menuItem_editDBConfigFile = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '接続先を編集' };
        $menuItem_openScriptDir = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'プログラムの場所を開く' };
        
        $notifyIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip;
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_editDBConfigFile);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_openScriptDir);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_exit);

        $menuItem_exit.add_Click( {
            $appContext.ExitThread();
        });

        $menuItem_editDBConfigFile.add_Click( {
            displayDBConfigFile
        });

        $menuItem_openScriptDir.add_Click( {
            openScriptDir
        });

        Write-Debug "  set right click event end"
        $notifyIcon.Visible = $true;

        # タスクトレイの表示
        Write-Debug "  run start"
        [void][System.Windows.Forms.Application]::Run($appContext);
        Write-Debug "  run end"

        $notifyIcon.Visible = $false;

    }
    finally {
        Write-Debug "  dispose start"
        $notifyIcon.Dispose();
        Write-Debug "  release mutex start"
        $mutex.ReleaseMutex();
    }
}

# タスクバー非表示
function hiddenTaskber {
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    $asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
}

try {
    # タイトルバーの書き換え
    $Host.UI.RawUI.WindowTitle = $APPLICATION_NAME
    # 多重起動チェック
    if ($mutex.WaitOne(0, $false)) {
        Write-Debug "hiddenTaskber START"
        hiddenTaskber
        Write-Debug "hiddenTaskber END"
        Write-Debug "displayTooltip START"
        displayTooltip
        Write-Debug "displayTooltip END"
        
        $retcode = 0;
    }
    else {
        Write-Error "同じプログラムがすでに起動中です。"
        $retcode = 255;
    }
}
finally {
    $mutex.Dispose();
}
exit $retcode;