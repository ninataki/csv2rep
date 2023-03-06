## SCO報告資料をcsvからエクセル加工するVBA

- RedmineからCSVエキスポート
  - クエリ表示する http://r4pm.soft.fujitsu.com/projects/a-list/issues?query_id=480
  - CSVエキスポート（UTF-8形式、最新コメントを追加）する
- マクロのインポート
  - .vbs拡張子のファイルにして保存し、エクセルの「開発タブ」VBA画面を表示しファイル→インポート。
    または、Module1を開いてコピペ）
  - 新しいファイルを毎回加工するのが面倒ならPERSONAL.XLSBに入れておく
- マクロの実行
  - 実行はF5または▽の再生ボタンみたいなのを押す。[ALL_xxx] で全実行される

- エクセルで保存
  - yyyymmdd_SCOreport.xlsxなど
