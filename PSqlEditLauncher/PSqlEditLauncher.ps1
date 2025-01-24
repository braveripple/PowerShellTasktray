<###############################################################
PSqlEdit ランチャー
###############################################################>

using namespace System.Management.Automation
Add-Type -AssemblyName System.Windows.Forms;

$APPLICATION_NAME = "PSqlEdit ランチャー"
$MUTEX_NAME = 'CE634DBE-31E2-4E65-838A-11CF22BDBBC4' + $APPLICATION_NAME;

$DB_LIST_FILE = "$PSScriptRoot/接続先.tsv"
$SCRIPT_FILE = $MyInvocation.MyCommand.Path
$SCRIPT_NAME = [System.IO.Path]::GetFileNameWithoutExtension($SCRIPT_FILE)
$EDITOR_PATH = "notepad.exe"

$mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME);

# 接続先の選択画面を開く
function displayAccessPoint {
    try {
        # JSONまたはTSVファイルから接続先情報を取得
        if ((Get-Item -LiteralPath $DB_LIST_FILE).Extension.ToLower() -eq ".json") {
            $dbList = Get-Content -LiteralPath $DB_LIST_FILE -Encoding utf8 | ConvertFrom-Json
        } else {
            $dbList = Import-Csv -LiteralPath $DB_LIST_FILE -Delimiter "`t" -Encoding utf8;
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
            Write-Host $tmp
            $db = $dbList | 
                Where-Object { $_.Id -eq $tmp.Id } | Select-Object -First 1
            Write-Host $db

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
function displayDBFile {
    try {
        & $EDITOR_PATH $DB_LIST_FILE
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
        Write-Host "  make context start"
        # コンテキスト作成
        $appContext = New-Object System.Windows.Forms.ApplicationContext;
        Write-Host "  make context end"
        ####################################################
        # 通知アイコン作成
        ####################################################
        Write-Host "  set icon start"
        # TeraTermのアイコンを設定する
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Program Files (x86)\teraterm\ttermpro.exe")
        $notifyIcon = [System.Windows.Forms.NotifyIcon]@{
            Icon           = $icon;
            Text           = $APPLICATION_NAME;
            BalloonTipIcon = 'None';
        };
        Write-Host "  set icon end"

        ####################################################
        # アイコン左クリック時のイベントを設定
        ####################################################
        Write-Host "  set left click event start"
        $notifyIcon.add_Click( {
                if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    displayAccessPoint
                }
            });
        Write-Host "  set left click event end"
        ####################################################
        # アイコン右クリック時のコンテキストメニューの設定
        ####################################################
        Write-Host "  set right click event start"
        $menuItem_exit = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'Exit' };
        $menuItem_editdbFile = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '接続先を編集' };
        $menuItem_openScriptDir = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'プログラムの場所を開く' };
        
        $notifyIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip;
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_editdbFile);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_openScriptDir);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_exit);

        $menuItem_exit.add_Click( {
            $appContext.ExitThread();
        });

        $menuItem_editdbFile.add_Click( {
            displayDBFile
        });

        $menuItem_openScriptDir.add_Click( {
            openScriptDir
        });

        Write-Host "  set right click event end"
        $notifyIcon.Visible = $true;

        # タスクトレイの表示
        Write-Host "  run start"
        [void][System.Windows.Forms.Application]::Run($appContext);
        Write-Host "  run end"

        $notifyIcon.Visible = $false;

    }
    finally {
        Write-Host "  dispose start"
        $notifyIcon.Dispose();
        Write-Host "  release mutex start"
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
    Write-Host $SCRIPT_FILE
    if (!(Test-Path -LiteralPath $DB_LIST_FILE -PathType Leaf)) {
        Write-Error "接続先設定ファイルが存在しません：$DB_LIST_FILE"
        exit 1
    }
    # タイトルバーの書き換え
    $Host.UI.RawUI.WindowTitle = $APPLICATION_NAME
    # 多重起動チェック
    if ($mutex.WaitOne(0, $false)) {
        Write-Host "hiddenTaskber START"
        hiddenTaskber
        Write-Host "hiddenTaskber END"
        Write-Host "displayTooltip START"
        displayTooltip
        Write-Host "displayTooltip END"
        
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