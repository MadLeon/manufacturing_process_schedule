' Class module: EventHandler_JobSelectorCombo
Option Explicit

Public WithEvents Combo As MSForms.ComboBox
Public DrawingDescriptions As Collection
Public JobSelectorForm As Object ' Reference to the JobSelector form

Private Sub Combo_Change()
    Dim selectedDrawingNumber As String
    Dim description As String

    selectedDrawingNumber = Combo.Text

    ' Retrieve the description based on the selected drawing number
    On Error Resume Next
    description = DrawingDescriptions(selectedDrawingNumber)
    On Error GoTo 0

    ' Update the Tag property of the Select button with the new description
    If JobSelectorForm.Controls.Item("btnSelect") Is Not Nothing Then
        Dim selectButton As MSForms.CommandButton
        Set selectButton = JobSelectorForm.Controls.Item("btnSelect")

        Dim parts() As String
        parts = Split(selectButton.Tag, "|")

        If UBound(parts) >= 2 Then
            selectButton.Tag = parts(0) & "|" & parts(1) & "|" & description
        End If
    End If
End Sub
