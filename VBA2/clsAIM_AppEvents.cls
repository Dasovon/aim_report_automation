'=== Class Module: clsAIM_AppEvents ===
Option Explicit
Public WithEvents App As Application

'====================================================================
'  GLOBAL EVENT HANDLER: Sync Inspection Status Across All Sheets
'====================================================================
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo CleanUp

    Const woCol As Long = 1 ' Work Order always in column A
    Dim inspCol As Long, i As Long
    Dim woNumber As String, newStatus As String

    ' Find the "Inspection Status" column
    inspCol = 0
    For i = 1 To Sh.UsedRange.Columns.Count
        If LCase(Trim(Sh.Cells(1, i).Value)) = "inspection status" Then
            inspCol = i
            Exit For
        End If
    Next i
    If inspCol = 0 Then Exit Sub
    If Target.Row = 1 Or Target.Column <> inspCol Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub

    ' Work order + new status
    woNumber = Trim(Sh.Cells(Target.Row, woCol).Value)
    newStatus = Trim(Target.Value)
    If woNumber = "" Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = "↻ Syncing Work Order " & woNumber & " = " & newStatus

    ' Push update to every sheet in the same workbook
    Dim ws As Worksheet, otherInspCol As Long, lastRow As Long, k As Long
    For Each ws In Sh.Parent.Worksheets
        otherInspCol = 0
        For i = 1 To ws.UsedRange.Columns.Count
            If LCase(Trim(ws.Cells(1, i).Value)) = "inspection status" Then
                otherInspCol = i
                Exit For
            End If
        Next i
        If otherInspCol = 0 Then GoTo NextSheet

        lastRow = ws.Cells(ws.Rows.Count, woCol).End(xlUp).Row
        For k = 2 To lastRow
            If Trim(ws.Cells(k, woCol).Value) = woNumber Then
                If ws.Cells(k, otherInspCol).Value <> newStatus Then
                    ws.Cells(k, otherInspCol).Value = newStatus
                End If
                Exit For
            End If
        Next k
NextSheet:
    Next ws

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
