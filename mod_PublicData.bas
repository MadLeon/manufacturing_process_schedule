' In modPublicData module
Option Explicit

Public lastEditedRow As Long

Public Function GetLastEditedRow() As Long
    GetLastEditedRow = lastEditedRow
End Function

Public Sub SetLastEditedRow(ByVal row As Long)
    lastEditedRow = row
End Sub