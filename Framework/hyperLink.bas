Option Explicit

' 指定のセルに指定のアドレスのハイパーリンクを貼る。
Sub createHyperLinkForUrl(ByVal range as Range, ByVal address as String, ByVal text as String)
  ActiveSheet.Hyperlinks.Add Anchor:=range, Address:=address, TextToDisplay:=text
End Sub

' 指定のセルに指定のメールアドレスのハイパーリンクを貼る。
Sub createHyperLinkForMailAddress(ByVal range as Range, ByVal address as String)
  ActiveSheet.Hyperlinks.Anchor:=range, Address:="mailto:" & address
End Sub

' 指定セルのハイパーリンクを削除する。
Sub deleteHyperLink(ByVal range as Range)
  range.Hyperlinks.Delete
End Sub
