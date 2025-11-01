Attribute VB_Name = "Module31"
Sub AddComment()
  Dim cmt As Comment
  Set cmt = ActiveCell.Comment
  If cmt Is Nothing Then
    ActiveCell.AddComment Text:=""
    Set cmt = ActiveCell.Comment
    With cmt.Shape.TextFrame.Characters.Font
      .Name = "Vandra"
      .Size = 11
      .Bold = True
      .ColorIndex = 0
    End With
  End If
  SendKeys "+{F2}"
End Sub
