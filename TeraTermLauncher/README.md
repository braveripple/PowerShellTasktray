# TeraTermランチャー

TeraTermの起動と接続処理をPowerShellで担う常駐スクリプト。

接続先一覧はOut-GridViewで作成。

接続情報は以下を設定可能

* Host：サーバ名またはIPアドレス
* User：ユーザ名
* Password：パスワード
* Kanji：サーバの文字コード
* Macro：マクロファイル(.ttl)を相対パスで指定。未指定の場合は空文字にする
* Memo：サーバの説明など
  
