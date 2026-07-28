' =========================================================
' Module: Data_SharedValues
' =========================================================
Option Explicit

' Number of items to input
Public ItemCount As Long

' Batch values (array of strings)
Public BatchValues() As String

' Start row of data table
Public startRow As Long

' Current number of sheets (excluding "Data")
Public SheetCount As Long

' Collection of merged header info objects
Public MergedHeaders As Collection

' Collection of target cells that match the selected distance and properties
Public TargetCells As Collection

Public AlertToleranceShown As Boolean

' Set by dir_signature/Handler_SaveLock.HandleCtrlSSave right before an explicit
' Ctrl+S save; consumed (and reset) by dir_signature/Function_DatabaseSync as
' soon as it's read, so it can never leak into a later, unrelated save. Tells
' the DB sync that this save should mark the DIR's part_attachment as
' 'completed' rather than leaving its status untouched.
Public IsCtrlSSave As Boolean