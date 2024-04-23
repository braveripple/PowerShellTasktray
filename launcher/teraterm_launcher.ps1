<###############################################################
Tera Term ランチャー
###############################################################>
#Requires -Version 3

using namespace System.Management.Automation
Add-Type -AssemblyName System.Windows.Forms;

$APPLICATION_NAME = "Tera Term ランチャー"
$MUTEX_NAME = 'CE634DBE-31E2-4E65-838A-11CF22BDBBC4' + $APPLICATION_NAME;

$TERATERM_PATH = "C:\Program Files (x86)\teraterm\ttermpro.exe"
$SERVER_FILE = "$PSScriptRoot/接続先.tsv"
$SCRIPT_FILE = $MyInvocation.MyCommand.Path
$EDITOR_PATH = "notepad.exe"
$LOGDIR_PATH = "$Env:OneDriveConsumer\Documents\Logs\TeratermLog"

$mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME);

# 接続先の選択画面を開く
function displayAccessPoint {
    try {
        # TSVファイルから接続先情報を取得
        $severList = Import-Csv -Path $SERVER_FILE -Delimiter "`t" -Encoding utf8
        # 選択画面表示。パスワードは画面に表示しない。
        $select = $SeverList | Select-Object User, Host, Kanji, Macro, Memo | Out-GridView -OutputMode Multiple -Title "接続先の選択"
        $select | ForEach-Object {
            $tmp = $_
            Write-Host $_
            $serverData = $severList | Where-Object { ($_.Host -eq $tmp.Host) -and ($_.User -eq $tmp.User) } | Select-Object -First 1
            $macroPath = "$PSScriptRoot\$($serverData.Macro)"
            $macroArgs = ""
            if (Test-Path $macroPath -PathType Leaf) {
                $tmpMacroPath = "$macroPath.$($serverData.User).ttl"
                (Get-Content $macroPath) -replace "<USERNAME>","$($serverData.User)" > $tmpMacroPath
                $macroArgs = "/M=$tmpMacroPath"
            }
            Write-Host $serverData
            & $TERATERM_PATH `
            /ssh "$($serverData.Host)" /2 /auth=password `
            /user="$($serverData.User)" /passwd="$($serverData.Password)" /L="$LOGDIR_PATH\&h_%Y%m%d_%H%M%S.log" `
            /KR="$($serverData.Kanji)" /KT="$($serverData.Kanji)" `
            "$macroArgs"
            Start-Sleep 1
        }
    }
    catch {
        $notifyIcon.BalloonTipText = $_.ToString();
        $notifyIcon.ShowBalloonTip(5000);
    }

}
function displayServerFile {
    try {
        & $EDITOR_PATH $SERVER_FILE
    }
    catch {
        $notifyIcon.BalloonTipText = $_.ToString();
        $notifyIcon.ShowBalloonTip(5000);
    }
}
function displayScriptFile {
    try {
        & $EDITOR_PATH $SCRIPT_FILE
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
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($TERATERM_PATH)
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
        $menuItem_editServerFile = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '接続先を編集' };
        $menuItem_editScriptFile = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'スクリプトを編集' };
        
        $notifyIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip;
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_editServerFile);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_editScriptFile);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_exit);

        $menuItem_exit.add_Click( {
            $appContext.ExitThread();
        });
        $menuItem_editServerFile.add_Click( {
            displayServerFile
        });
        $menuItem_editScriptFile.add_Click( {
            displayScriptFile
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
    if (!(Test-Path -LiteralPath $SERVER_FILE -PathType Leaf)) {
        Write-Error "接続先設定ファイルが存在しません：$SERVER_FILE"
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