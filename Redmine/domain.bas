Option Explicit

' CSVデータのデータ内容
Public チケットNo as String
Public ステータス as String
Public 発行日 as String
Public 期限 as String
Public 担当者 as String

' 処理用
Dim タイトル配列() As Variant
Dim データ配列() as Variant
Const フィールド数 = 5

' コンストラクタ
Friend Sub Class_Initialize()
  ReDim タイトル配列(フィールド数)
  ReDim データ配列(フィールド数)
  タイトル配列 = Array("チケットNo","ステータス","発行日","期限","担当者")
End Sub

'データをセットする。
Friend Sub データセット(ByVal 配列 as Variant)
  チケットNo = 配列(0)
  ステータス = 配列(1)
  発行日 = 配列(2)
  期限 = 配列(3)
  担当者 = 配列(4)
  データ配列 = Array(チケットNo,ステータス,発行日,期限,担当者)
End Sub

' 以下は変更不要

' タイトル行をセルに出力する。
Friend Sub タイトル行を出力(ByVal タイトル行No as Integer)
  Dim i as Variant
  For i=0 to UBound(タイトル配列)
    Cells(タイトル行No,i+1) = タイトル配列(i)
  Next データ
End Sub

' 任意の順番でセルに出力する。
Friend Sub セルに出力(ByVal 行No as Integer)
  Dim i as Variant
  For i=0 to UBound(データ配列)
    Cells(行No,i+1) = データ配列(i)
  Next データ
  Call セルの設定(行No)
End Sub

' セルのレイアウトを設定する。
Private Sub セルの設定(ByVal 行No as Integer)
  Const 灰色 = 15
  Const 黄色 = 6
  If ステータス = "完了" Then
    Range(Cells(行No,1),Cells(行No,UBound(データ配列)-1)).Interior.ColorIndex = 灰色
  ElseIf DateValue(期限) < Date Then
    Range(Cells(行No,1),Cells(行No,UBound(データ配列)-1)).Interior.ColorIndex = 黄色
  End If
  Range(Cells(行No,1),Cells(行No,UBound(データ配列)-1)).Borders.LineStyle = xlContinuous
End Sub
