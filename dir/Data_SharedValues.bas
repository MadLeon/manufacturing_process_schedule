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


