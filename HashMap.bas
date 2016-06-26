'HashMap

' キーのリスト
Dim keyList as New Collection

' 値のリスト
Dim valueList as New Collection

' コンストラクタ
Public Sub Class_Initialize()
End Sub

' デストラクタ
Public Sub Class_Terminate()
End Sub

' 要素を取得する
Public Function get(ByVal key as String) as String
  Dim value as String
  Dim idx as Integer
  For idx = 0 to keyList.Count
    If keyList.Item(idx) key then
      get = valueList.Item(idx)
    End If
  next idx
  get = ""
End Function

' 要素を追加する
Public Sub put(ByVal key as String, ByVal value as String)
  keyList.Add(key)
  valueList.Add(value)
End Sub

' 要素数を取得する
Public Function count() as Long
  count = keyList.Count
End Function

' キーの重複を除外する
Public Sub ChangeSet()
  Dim key as String
  Dim idx as Integer
  For Each key as keyList
    For idx = 0 to keyList.Count
      If keyList.Item(idx) = key then
        keyList.delete(idx)
        valueList.delete(idx)
        Exit For
      End If
    Next idx
  Next key
End Sub
