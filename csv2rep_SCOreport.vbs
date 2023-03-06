'Attribute VB_Name = "SCO報告資料"
'想定フォーマット
'題名   状況(最新コメント)    優先度(確度)  ステータス  課題あり    担当者    担当幹部    KFA g法主

Sub ALL_SCO報告書作成()
    Call 設定変更
    Call 全体整形
    Call 列入れ替え
    Call 題名と状況の整形
    Call イコールの除去
    Call オートフィルタ
    Call タイトルの色
End Sub



'文字のある範囲を選択、折り返す、左・中央ぞろえにする、目盛り線を消す、枠線をつける。上下幅：3列ほどにする。
Sub 設定変更()
    ActiveWindow.DisplayGridlines = False '目盛線を非表示
End Sub

Sub 全体整形()
    With Range("A1").CurrentRegion 'アクティブ領域の定義
    .Borders.LineStyle = xlContinuous '罫線をつける
    .HorizontalAlignment = xlLeft '左揃え
    .VerticalAlignment = xlCenter '上下中央揃え
    .WrapText = True '折り返し
    .ColumnWidth = 10 '横幅
    .RowHeight = 30 '上下幅
    .Interior.ColorIndex = xlNone
    End With
End Sub

'「題名」列を前にする、最新のコメント→状況
Sub 列入れ替え()
    If Range("B1").Value = "題名" Then '入替済みなら処理しない
        Exit Sub
    End If
    Columns("B").Cut '最新のコメント=列B →Cへ
    Columns("C").Insert Shift:=xlToRight
    Columns("C").Cut '題名=列C →Bへ
    Columns("B").Insert Shift:=xlToRight
    Range("C1").Value = "状況" 'C1セルをリネーム
End Sub

'最新のコメントと題名を横に伸ばす。最新のコメントは上揃え
Sub 題名と状況の整形()
    Range("B1").EntireColumn.ColumnWidth = 40  '横幅調整
    Range("C1").EntireColumn.ColumnWidth = 130 '状況の横幅調整
    Range("C1").Value = "状況" 'C1セルを変更
    Range("C1").EntireColumn.VerticalAlignment = xlTop '状況は上揃えにする
    Range("C1").VerticalAlignment = xlCenter 'タイトルだけ中央揃えに直す
End Sub

Sub イコールの除去()
    Cells.Replace What:="=", Replacement:="'"
End Sub

Sub オートフィルタ()
  Range("A1").AutoFilter
End Sub

'タイトルを濃い緑にする
Sub タイトルの色()
    Range("A1").AutoFilter Field:=1, Criteria1:="#" 'タイトル行を選択
    With Range("A1").CurrentRegion
    .Interior.ThemeColor = xlThemeColorAccent6 'offceカラーの一番右(緑)
    .Interior.TintAndShade = 0.5 '2番目に薄い色
    End With
    ActiveSheet.ShowAllData 'オートフィルタ全表示
End Sub

