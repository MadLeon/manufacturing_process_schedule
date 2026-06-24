Option Explicit

' Declare a variable to remember which TextBox was last focused
Private lastFocusedTextBox As MSForms.TextBox

' Module-level flag
Private showSTSymbol As Boolean




Private Sub txtTolerance_Enter()
    Set lastFocusedTextBox = Me.txtTolerance
    Debug.Print "Focused: txtTolerance"
End Sub
Private Sub txtPrimary_Enter()
    Set lastFocusedTextBox = Me.txtPrimary
    Debug.Print "Focused: txtPrimary"
End Sub
Private Sub txtSecondary_Enter()
    Set lastFocusedTextBox = Me.txtSecondary
    Debug.Print "Focused: txtSecondary"
End Sub
Private Sub txtTertiary_Enter()
    Set lastFocusedTextBox = Me.txtTertiary
    Debug.Print "Focused: txtTertiary"
End Sub

Private Sub UpdatePreview()
    Dim symbolStr As String
    Dim tolStr As String
    Dim primaryStr As String
    Dim secondaryStr As String
    Dim tertiaryStr As String
    Dim previewStr As String
    
    ' 1. Get current values
    symbolStr = Me.cboSymbol.Value
    tolStr = Me.txtTolerance.Value
    Me.txtPrimary.Value = UCase(Me.txtPrimary.Value)
    Me.txtTertiary.Value = UCase(Me.txtTertiary.Value)
    Me.txtSecondary.Value = UCase(Me.txtSecondary.Value)
    Me.txtTertiary.Value = UCase(Me.txtTertiary.Value)
    primaryStr = UCase(Me.txtPrimary.Value)
    secondaryStr = UCase(Me.txtSecondary.Value)
    tertiaryStr = UCase(Me.txtTertiary.Value)
    
    ' Exit early if everything is empty
    If symbolStr = "" And tolStr = "" And primaryStr = "" And secondaryStr = "" And tertiaryStr = "" And Not showSTSymbol Then
        Me.txtPreview.Value = ""
        Exit Sub
    End If
    
    ' 2. Add "~" to symbol
    If symbolStr <> "" Then symbolStr = symbolStr & "~"
    
    ' 3. Transform tolerance and datums
    If tolStr <> "" Then tolStr = TransformTolerance(tolStr)
    primaryStr = TransformDatum(primaryStr)
    secondaryStr = TransformDatum(secondaryStr)
    tertiaryStr = TransformDatum(tertiaryStr)
    
    ' 4. Build preview
    Dim sections As Variant
    sections = Array(tolStr, primaryStr, secondaryStr, tertiaryStr)
    
    previewStr = symbolStr
    Dim part As Variant
    For Each part In sections
        If part <> "" Then
            If previewStr <> "" Then previewStr = previewStr & "|"
            previewStr = previewStr & part
        End If
    Next part
    
    ' 5. Wrap with braces if non-empty
    If previewStr <> "" Then previewStr = "{" & previewStr & "}"
    
    ' 6. Append ST symbol if flag is True
    If showSTSymbol Then previewStr = previewStr & "Æ"
    
    ' 7. Update preview textbox
    Me.txtPreview.Value = previewStr
End Sub




Private Function TransformTolerance(tol As String) As String
    Dim i As Long, ch As String, result As String
    result = ""
    
    For i = 1 To Len(tol)
        ch = Mid$(tol, i, 1)
        Select Case ch
            Case "Ø": result = result & "Ø_"
            Case ".": result = result & "`."
            Case Else: result = result & ch & "~"
        End Select
    Next i
    
    TransformTolerance = result
End Function

Private Sub UserForm_Initialize()
    ' Populate the symbol dropdown
    With Me.cboSymbol
        .Font.Name = "y14.5-2018"
        .Font.Size = 18
        .AddItem "¿"   ' Position
        .AddItem "µ"   ' Flatness
        .AddItem "´"   ' Perpendicularity
        .AddItem "¶"   ' Parallelism
        .AddItem "·"   ' Concentricity
        .AddItem "¸"   ' Circular Runout
        .AddItem "¹"   ' Total Runout
        .AddItem "º"   ' Profile of a Surface
        .AddItem "»"   ' Profile of a Line
        .AddItem "¼"   ' Angularity
        .AddItem "½"   ' Cylindricity
        .AddItem "¾"   ' Symmetry
        .AddItem "³"   ' True Position (Alternate)
        .AddItem "²"   ' Straightness
    End With
    
    ' Setup preview box
    With Me.txtPreview
        .Font.Name = "y14.5-2018"
        .Font.Size = 18
        .Locked = True
        .Enabled = True
        .BackColor = vbWhite
        .Value = ""
    End With
    
    Me.txtPrimary.Font.Name = "y14.5-2018"
    Me.txtPrimary.Font.Size = 18
    Me.txtSecondary.Font.Name = "y14.5-2018"
    Me.txtSecondary.Font.Size = 18
    Me.txtTertiary.Font.Name = "y14.5-2018"
    Me.txtTertiary.Font.Size = 18
End Sub


Private Sub cboSymbol_Change()
    Dim sym As String
    sym = Me.cboSymbol.Value
    
    Dim txt As String
    txt = Me.txtPreview.Value
    
    Dim startPos As Long, endPos As Long
    
    startPos = InStr(txt, "{")
    endPos = InStr(txt, "~")
    
    If startPos > 0 And endPos > startPos Then
        ' Remove existing symbol(s) between { and ~
        txt = Left$(txt, startPos) & sym & Mid$(txt, endPos)
        Me.txtPreview.Value = txt
    ElseIf startPos > 0 Then
        ' If only { exists, append symbol before ~
        txt = Left$(txt, startPos) & sym & Mid$(txt, startPos + 1)
        Me.txtPreview.Value = txt
    Else
        ' Fallback: just prepend symbol
        Me.txtPreview.Value = sym & txt
    End If
    
    UpdatePreview
    
    Me.txtTolerance.SetFocus
End Sub

Private Function TransformDatum(datum As String) As String
    Dim i As Long
    Dim result As String
    Dim ch As String
    
    ' Return empty string if input is empty
    If datum = "" Then
        TransformDatum = ""
        Exit Function
    End If
    
    ' Loop through each character
    For i = 1 To Len(datum)
        ch = Mid$(datum, i, 1)
        result = result & ch & "~"
    Next i
    
    ' Return transformed string
    TransformDatum = result
End Function

Private Sub txtPrimary_Change()
    UpdatePreview
End Sub

Private Sub txtSecondary_Change()
    UpdatePreview
End Sub

Private Sub txtTertiary_Change()
    UpdatePreview
End Sub

Private Sub ToggleSymbol(symbol As String, Optional PrependOnly As Boolean = False)
    ' =============================================
    ' Determine which TextBox should be updated
    ' =============================================
    Dim targetTxt As MSForms.TextBox
    
    If Not lastFocusedTextBox Is Nothing Then
        Set targetTxt = lastFocusedTextBox
        Debug.Print "Last focused TextBox: " & targetTxt.Name
    Else
        Set targetTxt = Me.txtTolerance
        Debug.Print "No last focused TextBox, defaulting to txtTolerance"
    End If
    
    ' Ensure the font is correct
    targetTxt.Font.Name = "y14.5-2018"
    
    ' Get current text
    Dim txt As String
    txt = Trim(targetTxt.Value)
    Debug.Print "Before toggle: " & txt
    
    ' =============================================
    ' Handle diameter symbol (PrependOnly = True)
    ' =============================================
    If PrependOnly Then
        ' Diameter always modifies txtTolerance
        txt = Trim(Me.txtTolerance.Value)
        Debug.Print "Handling diameter, original txtTolerance: " & txt
        
        If Left$(txt, 1) = symbol Then
            txt = Mid$(txt, 2) ' Remove existing Ø
            Debug.Print "Diameter exists, removed"
        Else
            txt = symbol & txt ' Prepend Ø
            Debug.Print "Diameter added"
        End If
        
        Me.txtTolerance.Value = txt
        Me.txtTolerance.SetFocus
        Me.txtTolerance.SelStart = Len(txt) ' Cursor at end
        
    ' =============================================
    ' Handle other symbols (insert/remove at cursor)
    ' =============================================
    Else
        Dim curPos As Long
        curPos = targetTxt.SelStart
        Debug.Print "Cursor position: " & curPos
        
        If InStr(txt, symbol) > 0 Then
            ' Remove all occurrences of this symbol
            txt = Replace(txt, symbol, "")
            targetTxt.Value = txt
            targetTxt.SelStart = Application.Max(0, curPos - 1)
            Debug.Print "Symbol existed, removed. New value: " & txt
        Else
            ' Insert symbol at cursor position
            txt = Left$(txt, curPos) & symbol & Mid$(txt, curPos + 1)
            targetTxt.Value = txt
            targetTxt.SelStart = curPos + 1
            Debug.Print "Symbol inserted. New value: " & txt
        End If
        
        ' Refocus on the active TextBox
        targetTxt.SetFocus
    End If
End Sub


Private Sub UpdateTolerancePreview()
    Dim tol As String
    tol = Me.txtTolerance.Value  ' Get current tolerance text
    
    ' Build transformed tolerance string
    Dim newTol As String
    Dim i As Long, ch As String
    newTol = ""
    
    For i = 1 To Len(tol)
        ch = Mid$(tol, i, 1)
        
        If ch = "." Then
            newTol = newTol & "`."
        Else
            ' Treat Ø and all other characters the same way: append ~
            newTol = newTol & ch & "~"
        End If
    Next i
    
    ' Update txtPreview between | and }
    Dim txt As String
    txt = Me.txtPreview.Value
    
    Dim pipePos As Long, endBracePos As Long
    pipePos = InStr(txt, "|")        ' Find start of tolerance
    endBracePos = InStr(txt, "}")    ' Find end of tolerance
    
    If pipePos > 0 And endBracePos > pipePos Then
        ' Replace old tolerance with new transformed string
        txt = Left$(txt, pipePos) & newTol & Mid$(txt, endBracePos)
    Else
        ' Fallback: ensure braces exist and append
        txt = "{~_|" & newTol & "|}"
    End If
    
    Me.txtPreview.Value = txt
End Sub


Private Sub txtTolerance_Change()
    'UpdateTolerancePreview
    UpdatePreview
End Sub


' Diameter (always prepend)
Private Sub btnDiameter_Click()
    ToggleSymbol "Ø", True   ' skip targetBox, set IsDiameter = True
End Sub

' Other symbols (insert/remove at cursor)
Private Sub btnÅ_Click()
    ToggleSymbol "Å"
End Sub


Private Sub btnÒ_Click()
    ToggleSymbol "Ò"
End Sub

Private Sub btnÎ_Click()
    ToggleSymbol "Î"
End Sub

Private Sub btnÌ_Click()
    ToggleSymbol "Ì"
End Sub

Private Sub btnÃ_Click()
    ToggleSymbol "Ã"
End Sub

Private Sub btnË_Click()
    ToggleSymbol "Ë"
End Sub

Private Sub btnÑ_Click()
    ToggleSymbol "Ñ"
End Sub

Private Sub btnÆ_Click()
    ' Toggle the module-level flag
    showSTSymbol = Not showSTSymbol
    Debug.Print "ST symbol toggled. New state: " & showSTSymbol
    ' Rebuild preview with the current flag
    UpdatePreview
End Sub

Private Sub btnEnter_Click()
    Dim gdtString As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim shp As Shape
    
    Dim tolValue As String
    Dim rowNum As Long
    Dim i As Long, ch As String, cleanTol As String
    
    ' Get current GDT preview text
    gdtString = Me.txtPreview.Value
    
    ' =========================================
    ' Force replace any Enter/newline with ¶
    ' =========================================
    gdtString = Replace(gdtString, vbCrLf, ChrW(182))
    gdtString = Replace(gdtString, vbCr, ChrW(182))
    gdtString = Replace(gdtString, vbLf, ChrW(182))
    
    ' Safety check
    If gdtString = "" Then
        MsgBox "Preview is empty.", vbExclamation
        Exit Sub
    End If
    
    ' Ensure a cell is selected
    If ActiveCell Is Nothing Then
        MsgBox "No active cell selected.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    Set cell = ActiveCell
    
    ' Add a textbox shape over the active cell
    Set shp = ws.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=cell.Left, _
        Top:=cell.Top, _
        Width:=cell.Width, _
        Height:=cell.Height)
    
    ' Set textbox properties
    With shp.TextFrame2
        .TextRange.Text = gdtString
        .TextRange.Font.Name = "Y14.5-2018"
        .TextRange.Font.Size = 12
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse                      ' Disable word wrap
        .AutoSize = msoAutoSizeNone               ' Prevent auto resize
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With

    ' Remove fill and border for a clean GDT look
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' Allow overflow (Excel doesn't strictly "overflow" shapes,
    ' but disabling word wrap + fixed size gives visual overflow effect)
    shp.Placement = xlMoveAndSize
    
    ' Hide the form
    Me.Hide
    
    ' Add the upper tolerance
    rowNum = ActiveCell.Row
    tolValue = Me.txtTolerance.Value
    cleanTol = ""
    For i = 1 To Len(tolValue)
        ch = Mid$(tolValue, i, 1)
        If (ch Like "[0-9.-]") Then cleanTol = cleanTol & ch
    Next i
    If cleanTol <> "" Then
        ws.Cells(rowNum, "D").Value = cleanTol
    End If
    
    ' -------------------------
    ' Call formula logic (G/H columns) for this row
    ' -------------------------
    HandleGDTFormula ws, rowNum
    
    ' -------------------------
    ' Call tool logic (I column) using the gdt string
    ' -------------------------
    HandleGDTTool ws, rowNum
    
    Debug.Print "Inserted textbox at " & cell.Address & " with text: " & gdtString
End Sub

