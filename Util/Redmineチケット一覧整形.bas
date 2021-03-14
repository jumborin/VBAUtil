Option VBASupport 1
' 上記はLibreOfficeでの開発用のため、MicrosoftOfficeで実施する場合、削除すること。

Option Explicit

' ===========================================================
' CSVファイルを整形するマクロ
' 前提条件：対象シートがアクティブになること
' ===========================================================
Sub Redmine取り込みマクロ()
  Dim 最終行,最終列,ステータス列No,ループ処理用変数 as Integer
  最終行 = Cells(1, 1).End(xlDown).Row
  最終列 = Cells(1, 1).End(xlToRight).Column
  Const ステータス列名 = "ステータス"
  Const 完了済ステータス = "Resolved"

  ' 罫線を引く
  Range(Cells(1,1),Cells(最終行,最終列)).Borders.Weight = xlContinuous
  
  ' タイトル行に背景色をつける
  Range(Cells(1,1),Cells(1,最終列)).Interior.ColorIndex = 3
  
  ' 列幅を整理する
  Columns.EntireColumn.Autofit
  
  ' 1行目の値が「ステータス」となっている列を探し、ステータス列Noに設定
  For ループ処理用変数 = 1 to 最終列
    If Cells(1,ループ処理用変数) = ステータス列名 then
      ステータス列No =  ループ処理用変数
    End if
  Next
  
  ' 完了済ステータスになっている行の背景色を灰色にする
  For ループ処理用変数 = 1 to 最終行
    If Cells(ループ処理用変数,ステータス列No) = 完了済ステータス Then
      Range(Cells(ループ処理用変数,1),Cells(ループ処理用変数,最終列)).Interior.ColorIndex = 16
    End If
  Next
  
  ' オートフィルタで絞り込みされている場合、オートフィルタを1度解除し、再設定する
  If ActiveSheet.AutoFilterMode = True Then
    Range(Cells(1,1),Cells(最終行,1)).AutoFilter
  End If
  Range(Cells(1,1),Cells(最終行,1)).AutoFilter
  
End Sub