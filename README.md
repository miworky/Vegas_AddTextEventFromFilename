# Vegas_AddTextEventFromFilename
# ●概要

このプログラムは、Vegas Pro 18.0 のプラグインです。  
このプラグインを使用すると、タイムラインに配置してある「みてねからダウンロードしたメディアファイル」のファイル名を元に、テロップを生成できます。

# ●背景

みてねから画像や動画をダウンロードして動画を作成したり、ブルーレイに焼いたりする際に、いくつか課題があります。
課題の一つは、画像や動画のテロップを作成するのが大変だということです。
テロップは画像や動画1つ1つに対して1つずつ手動で追加する必要があるので、画像や動画の数が多いと非常に時間がかかります。

このプラグインを使用すれば、この作業を自動化できます。

# ●動画作成手順

1)みてねからファイルをダウンロードする

  https://github.com/miworky/miteneDownloader

を使用してダウンロードしてください。
  ダウンロードしたファイルは、「YYYY-MM-DDThhmmss_1つめのコメント」というファイル名になります。
  
2)ダウンロードしたファイルを Vegas Pro のメディアプールに取り込み、タイムラインに貼り付けます。

　ダウンロードしたファイル名に日付が含まれているので、これだけでみてねからダウンロードしたファイルを撮影日時順にタイムラインに配置できます。

3)撮影日とコメントをテキストイベントに登録します（本プログラムを使用します）

4)オリジナルの高解像度のファイルに差し替えます

   https://github.com/miworky/Vegas_ReplaceMediaFiles

を使用すると、自動でオリジナルの高解像度のファイルに差し替えできます。

5)お好きなBGMを貼り付けます

6)動画として書き出したり、ブルーレイに焼いたりします。

7)テキストイベントの時刻と内容をテキストファイルに書き出します

  https://github.com/miworky/Vegas_ExportTextEvent

　を使用すると自動でテキストイベントの時刻と内容をテキストファイルに書き出せます。

# ●開発環境

VisualStudio 2019 C#

# ●ビルド方法

1)AddTextEventFromFilename.sln を VisualStudio2019 で開きます

2)Release, Any CPUでビルド

AddTextEventFromFilename\bin\Releaseに成果物ができあがります。


# ●デプロイ方法

C:\ProgramData\VEGAS Pro\Script Menu
に以下のファイルをコピーします：

AddTextEventFromFilename\bin\Release\AddTextEventFromFilename.dll

# ●実行方法

1)Vegas Pro 18.0 で、あらかじめ「みてねからダウンロードしたファイル」をタイムラインのトラック2以降に配置しておきます

注意：テキストイベントはトラック1に追加されます。

2)Vegas Pro 18.0 から本プラグインを実行します

3)ファイル選択ダイアログが開くので、作成するログのファイル名を指定します

4)しばらく時間が経った後に、「終了しました」というポップアップが開けば終了です

　2で選択したファイルに作業ログが出力されています。
 
 　作業ログには以下の内容が出力されます：

タイムラインの時刻　撮影日　静止画・動画のコメント
     
