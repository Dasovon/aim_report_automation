Attribute VB_Name = "mod_AIM_Formatter"
'===============================================
' Module: mod_AIM_Formatter
' Purpose: Format AIM CSV exports for inspections
' Adds Floor, Room, Inspection Status,
' sorts, colors, and splits by building.
' Works on both Windows and macOS (no ActiveX, no CreateObject)
'===============================================
Option Explicit

Private Function KeyExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim tmp
    tmp = col.Item(key) ' will raise error if missing
    KeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub Run_AIM_Formatter()
    Dim ws As Worksheet, newWS As Worksheet
    Dim lastRow As Long, descCol As Long, floorCol As Long, roomCol As Long, propCol As Long
    Dim i As Long, descText As String, floorVal As String, roomVal As String
    Dim bldgCode As String, bldgName As String
    Dim uniqProps As Collection, key As Variant
    Dim todayName As String, floorRankCol As Long, roomRankCol As Long, inspectionCol As Long

    On Error GoTo ErrorHandler

    Dim originalErrorCheck As Boolean
    originalErrorCheck = Application.ErrorCheckingOptions.NumberAsText
    Application.ErrorCheckingOptions.NumberAsText = False

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set ws = ActiveSheet

    '=== Find columns ===
    descCol = 0
    propCol = 0
    For i = 1 To ws.UsedRange.Columns.Count
        Select Case LCase(Trim(ws.Cells(1, i).Value))
            Case "description": descCol = i
            Case "property": propCol = i
        End Select
    Next i

    If descCol = 0 Then
        MsgBox "Description column not found. Please ensure your CSV has a 'Description' column.", vbCritical
        GoTo CleanExit
    End If

    '=== Add new columns ===
    floorCol = ws.UsedRange.Columns.Count + 1
    roomCol = floorCol + 1
    inspectionCol = roomCol + 1

    ws.Cells(1, floorCol).Value = "Floor"
    ws.Cells(1, roomCol).Value = "Room"
    ws.Cells(1, inspectionCol).Value = "Inspection Status"
    With ws.Cells(1, inspectionCol)
        .Interior.Color = RGB(200, 200, 200)
        .Font.Bold = True
    End With

    floorRankCol = inspectionCol + 1
    roomRankCol = inspectionCol + 2
    ws.Cells(1, floorRankCol).Value = "__FloorRank"
    ws.Cells(1, roomRankCol).Value = "__RoomRank"

    lastRow = ws.Cells(ws.Rows.Count, descCol).End(xlUp).Row
    Application.StatusBar = "Processing " & (lastRow - 1) & " work orders..."

    '=== Parse each row ===
    For i = 2 To lastRow
        descText = CStr(ws.Cells(i, descCol).Value)
        floorVal = "": roomVal = ""

        '----------------------------------------
        ' Floor/Room extraction (pure VBA)
        '----------------------------------------
        Dim floorPos As Long, roomPos As Long, afterText As String

        floorPos = InStr(1, descText, "Floor:", vbTextCompare)
        If floorPos = 0 Then floorPos = InStr(1, descText, "Flr:", vbTextCompare)
        If floorPos > 0 Then
            afterText = Mid$(descText, floorPos + 6)
            If InStr(afterText, " ") > 0 Then
                floorVal = Trim(Split(afterText, " ")(0))
            Else
                floorVal = Trim(afterText)
            End If
        End If

        roomPos = InStr(1, descText, "Room:", vbTextCompare)
        If roomPos = 0 Then roomPos = InStr(1, descText, "Rm:", vbTextCompare)
        If roomPos > 0 Then
            afterText = Mid$(descText, roomPos + 5)
            If InStr(afterText, " ") > 0 Then
                roomVal = Trim(Split(afterText, " ")(0))
            Else
                roomVal = Trim(afterText)
            End If
        End If
        '----------------------------------------

        '=== Fallback logic for floors ===
        If floorVal = "" And roomVal <> "" Then
            If Len(roomVal) >= 4 And IsNumeric(Left$(roomVal, 2)) Then
                Select Case Left$(roomVal, 2)
                    Case "10": floorVal = "10"
                    Case "11": floorVal = "11"
                    Case "12": floorVal = "12"
                End Select
            End If
            If floorVal = "" Then
                Select Case Left$(roomVal, 1)
                    Case "0": floorVal = "B"
                    Case "1": floorVal = "1"
                    Case "2": floorVal = "2"
                    Case "3": floorVal = "3"
                    Case "4": floorVal = "4"
                    Case "5": floorVal = "5"
                    Case "6": floorVal = "6"
                    Case "7": floorVal = "7"
                    Case "8": floorVal = "8"
                    Case "9": floorVal = "9"
                End Select
            End If
        End If

        '=== Assign parsed values ===
        Select Case UCase$(floorVal)
            Case "SF": ws.Cells(i, floorCol).Value = "SF"
            Case "0", "B": ws.Cells(i, floorCol).Value = "B"
            Case Else: ws.Cells(i, floorCol).Value = floorVal
        End Select

        If roomVal <> "" Then
            ws.Cells(i, roomCol).NumberFormat = "@"
            ws.Cells(i, roomCol).Value = roomVal
        Else
            ws.Cells(i, roomCol).ClearContents
        End If

        ws.Cells(i, inspectionCol).Value = "Pending"

        '=== Sorting helpers ===
        Dim fRank As Long
        Select Case UCase$(CStr(ws.Cells(i, floorCol).Value))
            Case "B": fRank = 0
            Case "SF": fRank = 99
            Case Else
                If Trim$(CStr(ws.Cells(i, floorCol).Value)) = "" Then
                    fRank = 999
                ElseIf IsNumeric(ws.Cells(i, floorCol).Value) Then
                    fRank = Val(CStr(ws.Cells(i, floorCol).Value))
                Else
                    fRank = 999
                End If
        End Select
        ws.Cells(i, floorRankCol).Value = fRank

        Dim rText As String, rRank As Long, numPrefix As String
        rText = UCase$(Trim$(CStr(ws.Cells(i, roomCol).Value)))
        If rText = "" Then
            rRank = 999999
        ElseIf InStr(rText, "HALL") > 0 Or InStr(rText, "STR") > 0 Or InStr(rText, "ELEV") > 0 Then
            rRank = 700000
        Else
            Dim j As Long, ch As String
            numPrefix = ""
            For j = 1 To Len(rText)
                ch = Mid$(rText, j, 1)
                If ch Like "[0-9]" Then
                    numPrefix = numPrefix & ch
                Else
                    Exit For
                End If
            Next j
            If Len(numPrefix) > 0 Then
                rRank = CLng(numPrefix)
            Else
                rRank = 600000
            End If
        End If
        ws.Cells(i, roomRankCol).Value = rRank

        '=== Building detection ===
        bldgCode = "": bldgName = ""
        If propCol > 0 Then
            bldgCode = Trim$(CStr(ws.Cells(i, propCol).Value))
            If bldgCode <> "" And IsNumeric(bldgCode) Then bldgCode = CStr(CLng(bldgCode))
        End If

        If bldgCode = "" Then
            If InStr(1, descText, "Emerging Technologies Building", vbTextCompare) > 0 Then
                bldgCode = "270": bldgName = "ETB"
            ElseIf InStr(1, descText, "Wisenbaker Engineering Building", vbTextCompare) > 0 Then
                bldgCode = "682": bldgName = "WEB"
            ElseIf InStr(1, descText, "H.J. (Bill) and Reta Haynes Engineering Building", vbTextCompare) > 0 Then
                bldgCode = "492": bldgName = "HEB"
            End If
        Else
            Select Case bldgCode
                Case "270": bldgName = "ETB"
                Case "682": bldgName = "WEB"
                Case "492": bldgName = "HEB"
            End Select
        End If

        If bldgCode <> "" And bldgName <> "" Then
            If propCol = 0 Then
                propCol = ws.UsedRange.Columns.Count + 1
                ws.Cells(1, propCol).Value = "Property"
            End If
            ws.Cells(i, propCol).Value = bldgCode & "-" & bldgName
        End If

        If i Mod 10 = 0 Then
            Application.StatusBar = "Processing row " & i & " of " & lastRow
        End If
    Next i

    '=== Sort ===
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Columns(floorRankCol), Order:=xlAscending
    ws.Sort.SortFields.Add key:=ws.Columns(roomRankCol), Order:=xlAscending
    With ws.Sort
        .SetRange ws.UsedRange
        .Header = xlYes
        .Apply
    End With

    Application.DisplayAlerts = False
    ws.Columns(roomRankCol).Delete
    ws.Columns(floorRankCol).Delete
    Application.DisplayAlerts = True

    '=== Dropdown for inspection status ===
    inspectionCol = 0
    For i = 1 To ws.UsedRange.Columns.Count
        If LCase$(Trim$(CStr(ws.Cells(1, i).Value))) = "inspection status" Then
            inspectionCol = i
            Exit For
        End If
    Next i

    If inspectionCol > 0 Then
        With ws.Range(ws.Cells(2, inspectionCol), ws.Cells(lastRow, inspectionCol)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="Pending,Complete,Incomplete,Needs Review"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End If

    ws.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    ws.Cells.EntireColumn.AutoFit

    With ws.Columns(descCol)
        .ColumnWidth = 60
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    With ws.Columns(inspectionCol)
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
    End With

    '=== Split by building (no Scripting.Dictionary) ===
    If propCol > 0 Then
        Set uniqProps = New Collection

        ' collect unique property keys; store key as both item and key
        For i = 2 To ws.Cells(ws.Rows.Count, propCol).End(xlUp).Row
            Dim propVal As String
            propVal = Trim$(CStr(ws.Cells(i, propCol).Value))
            If propVal <> "" Then
                If Not KeyExists(uniqProps, propVal) Then
                    uniqProps.Add propVal, propVal
                End If
            End If
        Next i

        If uniqProps.Count > 1 Then
            Dim idx As Long
            For idx = 1 To uniqProps.Count
                key = uniqProps(idx)
                Dim shortName As String
                shortName = Split(CStr(key), "-")(1)

                On Error Resume Next
                Application.DisplayAlerts = False
                Worksheets(shortName).Delete
                Application.DisplayAlerts = True
                On Error GoTo 0

                ws.Copy After:=Sheets(Sheets.Count)
                Set newWS = ActiveSheet
                newWS.Name = shortName

                Dim r As Long, lr As Long
                lr = newWS.Cells(newWS.Rows.Count, propCol).End(xlUp).Row
                For r = lr To 2 Step -1
                    If Trim$(CStr(newWS.Cells(r, propCol).Value)) <> CStr(key) Then
                        newWS.Rows(r).Delete
                    End If
                Next r

                AddInspectionFormatting newWS, inspectionCol
            Next idx
        End If
    End If

    todayName = "WO's for " & Format(Date, "yyyy-mm-dd")
    On Error Resume Next
    ws.Name = todayName
    On Error GoTo ErrorHandler

    AddInspectionFormatting ws, inspectionCol

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.ErrorCheckingOptions.NumberAsText = originalErrorCheck

    MsgBox "AIM Formatter complete!" & vbCrLf & vbCrLf & _
           "- Floors/rooms sorted" & vbCrLf & _
           "- Inspection Status column added" & vbCrLf & _
           "- Cross-platform compatible (no ActiveX)", vbInformation, "Formatting Complete"
    Exit Sub

ErrorHandler:
    Application.ErrorCheckingOptions.NumberAsText = originalErrorCheck
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub


'====================================================
' CONDITIONAL FORMATTING HELPER
'====================================================
Sub AddInspectionFormatting(ws As Worksheet, inspectionCol As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, inspectionCol).End(xlUp).Row
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.UsedRange.Columns.Count)).FormatConditions.Delete

    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.UsedRange.Columns.Count)).FormatConditions.Add( _
            Type:=xlExpression, Formula1:="=$" & Split(ws.Cells(1, inspectionCol).Address, "$")(1) & "2=""Complete""" _
        )
        .Interior.Color = RGB(198, 239, 206)
        .StopIfTrue = False
    End With

    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.UsedRange.Columns.Count)).FormatConditions.Add( _
            Type:=xlExpression, Formula1:="=$" & Split(ws.Cells(1, inspectionCol).Address, "$")(1) & "2=""Incomplete""" _
        )
        .Interior.Color = RGB(255, 199, 206)
        .StopIfTrue = False
    End With

    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.UsedRange.Columns.Count)).FormatConditions.Add( _
            Type:=xlExpression, Formula1:="=$" & Split(ws.Cells(1, inspectionCol).Address, "$")(1) & "2=""Needs Review""" _
        )
        .Interior.Color = RGB(255, 235, 156)
        .StopIfTrue = False
    End With
End Sub


