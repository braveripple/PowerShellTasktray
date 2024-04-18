# PowerShellTasktray

## 結局何で作ればいいのか

PowerShellのタスクトレイ常駐プログラムはWindows PowerShellで作ること。

PowerShell CoreおよびPowreShell 7以降で動かすと変なエラーを吐く。

* 文字コード：UTF-8 BOMにすること。
* 改行コード：Windowsに倣ってCRLFでいいと思う。

## プロンプト残る問題の解決

PowerShellのタスクトレイ常駐プログラムをダブルクリックで起動すると、プロンプトが残る。

プロンプトを出さないようにするにはショートカット経由でプログラムを実行する。
```リンク先
powershell.exe -NoProfile -ExecutionPolicy Unrestricted <PowerShellのタスクトレイ常駐プログラムのフルパス>
```

* 参考ページ：[私PowerShellだけど、君のタスクトレイで暮らしたい](https://qiita.com/magiclib/items/cc2de9169c781642e52d)


* 参考ページ2：[タスクトレイで指令を待ち続ける健気な PowerShell スクリプト | Aqua Ware つぶやきブログ](https://aquasoftware.net/blog/?p=1244)
