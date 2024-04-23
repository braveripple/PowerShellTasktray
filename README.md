# PowerShellTasktray

## 結局何で作ればいいのか

可能な限りWindows PowerShellとPowerShell7の両方で動作するスクリプトを目指す。

* 文字コード：UTF-8にする。
* 改行コード：LFにする。

## プロンプト残る問題の解決

PowerShellのタスクトレイ常駐プログラムをダブルクリックで起動すると、プロンプトが残る。

プロンプトを出さないようにするにはショートカット経由でプログラムを実行する。
```リンク先
powershell.exe -NoProfile -ExecutionPolicy Unrestricted <PowerShellのタスクトレイ常駐プログラムのフルパス>
```

* 参考ページ：[私PowerShellだけど、君のタスクトレイで暮らしたい](https://qiita.com/magiclib/items/cc2de9169c781642e52d)


* 参考ページ2：[タスクトレイで指令を待ち続ける健気な PowerShell スクリプト | Aqua Ware つぶやきブログ](https://aquasoftware.net/blog/?p=1244)
