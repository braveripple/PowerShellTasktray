# PowerShellTasktray

PowerShellで作った常駐プログラム置き場

可能な限りWindows PowerShell/PowerShell7以降の両方で動くように以下の対応をしている。

* スクリプトをUTF8BOM付きで保存
  
他何かあったら書く

## プロンプト残る問題の解決

PowerShellのタスクトレイ常駐プログラムをダブルクリックで起動すると、プロンプトが残る。

プロンプトを出さないようにするにはショートカット経由でプログラムを実行する。
```リンク先
powershell.exe -NoProfile -ExecutionPolicy Unrestricted <PowerShellのタスクトレイ常駐プログラムのフルパス>
```

* 参考ページ：[私PowerShellだけど、君のタスクトレイで暮らしたい](https://qiita.com/magiclib/items/cc2de9169c781642e52d)


* 参考ページ2：[タスクトレイで指令を待ち続ける健気な PowerShell スクリプト | Aqua Ware つぶやきブログ](https://aquasoftware.net/blog/?p=1244)
