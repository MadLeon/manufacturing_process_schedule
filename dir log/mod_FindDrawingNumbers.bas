Option Explicit

' --- Configuration Variables ---
Const DB_PATH As String = "\\rtdnas2\OE\jobs.db"

Public Sub FindDrawingNumbers(job_number As String)
  ' Find related drawing_number based on job_number
  Dim dbPath As String, sql As String, results As Variant, i As Long, part_number As String
  ' Use a Dictionary to store drawing numbers and their descriptions
  Dim drawingInfo As Object ' Late binding for Dictionary
  Set drawingInfo = CreateObject("Scripting.Dictionary")
  Dim drawingNumber As String, description As String
  Dim selectedRow As Long

  ' 1. Initialize database connection
  dbPath = DB_PATH ' Assuming the database file is in the same directory as the workbook
  If Not InitializeSQLite(dbPath) Then
    MsgBox "Failed to initialize database connection.", vbCritical
    Exit Sub
  End If

  ' 2. Find part_number, oe_number, and po_number in the jobs table based on job_number
  sql = "SELECT part_number, oe_number, po_number, part_description FROM jobs WHERE job_number = '" & job_number & "'"
  results = ExecuteSQL(sql)
  If IsNull(results) Then
    MsgBox "No record found for job_number " & job_number & ".", vbInformation
    GoTo Cleanup
  End If

  part_number = results(0)(0) ' Get part_number
  Dim oe_number As String, po_number As String, part_description As String
  oe_number = results(0)(1)
  po_number = results(0)(2)
  part_description = results(0)(3) 'Part Description from Jobs Table

  ' Add the initial part_number to the drawingInfo dictionary with its part_description
  drawingInfo(part_number) = part_description

  ' 3. Find drawing_number and description in the assemblies table based on part_number
  sql = "SELECT drawing_number, description FROM assemblies WHERE part_number = '" & part_number & "'"
  results = ExecuteSQL(sql)
  If IsNull(results) Then
    MsgBox "No drawings found for part_number " & part_number & ".", vbInformation
    GoTo Cleanup
  End If

  ' Add drawing_number and description to the drawingInfo dictionary
  For i = LBound(results) To UBound(results)
    drawingNumber = results(i)(0)
    description = results(i)(1)
    If Not drawingInfo.Exists(drawingNumber) Then
      drawingInfo(drawingNumber) = description
    End If
  Next i

  ' 4. Display drawing_number in the JobSelector form
  Load JobSelector
  With JobSelector
    .Caption = "Select Drawing Number" ' Select Drawing Number
    ' Clear existing controls
    While .Controls.Count > 0
      Unload .Controls(0)
    Wend

    ' Add a ComboBox
    Dim combo As MSForms.ComboBox
    Set combo = .Controls.Add("Forms.ComboBox.1")
    With combo
      .Name = "cboDrawingNumber"
      .Top = 12
      .Left = 12
      .Width = 200
      ' First item is always the part_number
      .AddItem part_number
      ' Add the rest of the drawing numbers
      For i = 0 To drawingInfo.Count - 1
          If i > 0 Then 'Skip the first since we already added part_number
            .AddItem drawingInfo.Keys()(i)
          End If
      Next i
    End With

    ' Add a Select Button
    Dim selectButton As MSForms.CommandButton
    Set selectButton = .Controls.Add("Forms.CommandButton.1")
    
    Dim buttonHandler As EventHandler_JobSelectorButton
    Set buttonHandler = New EventHandler_JobSelectorButton
    
    With selectButton
      .Name = "btnSelect"
      .Caption = "Select"
      .Top = 40
      .Left = 12
      .Width = 70
      .Height = 24
      ' Event handler for Select Button
      ' Pass the drawingInfo dictionary and selected drawing number to the Tag
      .Tag = oe_number & "|" & po_number & "|" & part_number 'Store initial values
      Set buttonHandler.Btn = selectButton
    End With

    ' Expand the form
    .Width = 230
    .Height = 100
    .Show

  End With

Cleanup:
  ' 5. Close database connection
  CloseSQLite
End Sub

Private Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  ' Check if the string is in the array
  Dim element As Variant
  For Each element In arr
    If element = stringToBeFound Then
      IsInArray = True
      Exit Function
    End If
  Next element
  IsInArray = False
End Function