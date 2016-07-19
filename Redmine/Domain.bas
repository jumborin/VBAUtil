Option Explicit

' CSVデータのデータ内容
Public チケットNo As String
Public ステータス As String
Public 発行日 As String
Public 期限 As String
Public 担当者 As String

' 処理用
Dim ワークシート As New Worksheet
Dim タイトル配列() As Variant
Dim データ配列() As Variant
Public Const フィールド数 = 5
Const 灰色 = 15
Const 黄色 = 6

' コンストラクタ
Public Sub Class_Initialize()
  ReDim タイトル配列(フィールド数)
  ReDim データ配列(フィールド数)
  タイトル配列 = Array("チケットNo", "ステータス", "発行日", "期限", "担当者")
End Sub

'出力先シートをセットする。
Public Sub 出力先シートセット(ByVal シート名 As String)
  Set ワークシート = Worksheets(シート名)
End Sub

'データをセットする。
Public Sub データセット(ByVal 配列 As Variant)
  チケットNo = 配列(0)
  ステータス = 配列(1)
  発行日 = 配列(2)
  期限 = 配列(3)
  担当者 = 配列(4)
  データ配列 = Array(チケットNo, ステータス, 発行日, 期限, 担当者)
End Sub

' 以下は変更不要

' タイトル行をセルに出力する。
Public Sub タイトル行を出力(ByVal ワークシート As Worksheet, ByVal タイトル行No As Integer)
  Dim i As Variant
  For i = 0 To UBound(タイトル配列)
    ワークシート.Cells(タイトル行No, i + 1) = タイトル配列(i)
  Next データ
  ワークシート.Range(Cells(タイトル行No,1),Cells(タイトル行No,UBound(タイトル配列+1)).Interior.ColorIndex = 灰色
  ワークシート.Range(Cells(タイトル行No,1),Cells(タイトル行No,UBound(タイトル配列+1)).Borders.LineStyle = xlContinuous
End Sub

' 任意の順番でセルに出力する。
Public Sub セルに出力(ByVal ワークシート As Worksheet, ByVal 行No As Integer)
  Dim i As Integer
  For i = 0 To UBound(データ配列)
    ワークシート.Cells(行No, i + 1) = データ配列(i)
  Next データ
  Call セルの設定(行No)
End Sub

' セルのレイアウトを設定する。
Private Sub セルの設定(ByVal 行No As Integer)
  If ステータス = "完了" Then
    Range(Cells(行No, 1), Cells(行No, UBound(データ配列) - 1)).Interior.ColorIndex = 灰色
  ElseIf DateValue(期限) < Date Then
    Range(Cells(行No, 1), Cells(行No, UBound(データ配列) - 1)).Interior.ColorIndex = 黄色
  End If
  Range(Cells(行No, 1), Cells(行No, UBound(データ配列) - 1)).Borders.LineStyle = xlContinuous
End Sub
