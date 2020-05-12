Attribute VB_Name = "Module1"
Option Explicit

Sub Clear()
Attribute Clear.VB_ProcData.VB_Invoke_Func = " \n14"
Dim wb As Workbook
Dim ws As Worksheet

Set wb = ThisWorkbook
Set ws = zapisnik
    ws.Range("B6:K65,N6").Select
    Selection.ClearContents
    ws.Range("B6").Select
End Sub
