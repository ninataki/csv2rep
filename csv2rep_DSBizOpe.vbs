'Attribute VB_Name = "報告資料"
'想定フォーマット
'題名    状況    ステータス  課題あり    担当者  期日    開始日  更新日  優先度  親チケット

Sub ALL_市場企画部報告書作成()
    Call 設定変更
    Call 全体整形
    Call 列入れ替え
    Call 題名と状況の整形
    Call 親チケット処理
    Call タイトルの色
    Call ステータス処理
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
    .RowHeight = 80 '上下幅
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


'フィルタ化する、親チケット(「親チケット」列が空白)を選択する。
'タイトルと親チケットを上下幅1行にする、折り返さない、薄緑にする、親チケットのタイトル以外を削除、「ステータス」を非空白にする
'色見 <https://kosapi.com/post-3405/>
Sub 親チケット処理()

    '親チケットをフィルタして整形
    Range("A1").AutoFilter Field:=11, Criteria1:="" '親チケットが空白行
    With Range("A1").CurrentRegion
    .Interior.ThemeColor = xlThemeColorAccent6 'offceカラーの一番右(緑)
    .Interior.TintAndShade = 0.8 '1番目に薄い色
    .RowHeight = 30 '上下幅を小さくする
    .WrapText = False '折り返さない
    End With
    
    '状況の不要な箇所を削除
    Call タイトル除く選択(3)  '３列目＝状況
    Selection.ClearContents '削除

    'ステータスを非空白にする
    Call タイトル除く選択(4) '４列目＝ステータス
    Selection.Value = "-"

    ActiveSheet.ShowAllData 'オートフィルタ全表示

End Sub

'タイトルを濃い緑にする
Sub タイトルの色()
    
    Range("A1").AutoFilter Field:=11, Criteria1:="親*" 'タイトル行を選択
    With Range("A1").CurrentRegion
    .Interior.ThemeColor = xlThemeColorAccent6 'offceカラーの一番右(緑)
    .Interior.TintAndShade = 0.5 '2番目に薄い色
    End With

    ActiveSheet.ShowAllData 'オートフィルタ全表示
    
End Sub


'ステータスに色を付ける「Doing→緑、課題ありは黄色、」「Review→水色」「ToD→薄緑」「Backlog→灰色」
'ステータスの「Backlog」を非表示にする
' ColorIndex  <https://tripbowl.com/excel-vba/label-color-change/#color>
Sub ステータス処理()
    
    'ステータスの色付け
    Range("A1").AutoFilter Field:=4, Criteria1:="Review"
    Call タイトル除く選択(4)
    Selection.Interior.ColorIndex = 33 '青

    Range("A1").AutoFilter Field:=4, Criteria1:="Backlog"
    Call タイトル除く選択(4)
    Selection.Interior.ColorIndex = 15 'グレー

    Range("A1").AutoFilter Field:=4, Criteria1:="ToDo"
    Call タイトル除く選択(4)
    Selection.Interior.ColorIndex = 43 '薄い緑

    Range("A1").AutoFilter Field:=4, Criteria1:="Doing"
    Call タイトル除く選択(4)
    Selection.Interior.ColorIndex = 10 '緑

    Range("A1").AutoFilter Field:=5, Criteria1:="はい" 'ステータス Doing & 課題あり
    Call タイトル除く選択(4)
    Selection.Interior.ColorIndex = 6 '黄色
    
    'オートフィルタ設定しBacklogを除外して表示
    Range("A1").AutoFilter 'フィルタ解除
    Range("A1").AutoFilter Field:=4, Criteria1:="<>Backlog"
    
End Sub

Function タイトル除く選択(ByVal num As Integer)
    Range("A1").CurrentRegion.Columns(num).Select 'ステータス列Dを選択
    Selection.Offset(1, 0).Select '選択領域を一つ下にずらす
    Selection.Resize(Selection.Rows.Count - 1).Select '選択領域を一つ減らす
End Function

