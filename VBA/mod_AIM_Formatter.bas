Attribute VB_Name = "mod_AIM_Formatter"
'=== Module: mod_AIM_Formatter ===
Option Explicit

Private Const ENABLE_DASHBOARD As Boolean = True

'===========================================================
' MAIN ENTRY
'===========================================================
Sub Run_AIM_Formatter()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim descCol As Long, propCol As Long, dateCreatedCol As Long
    Dim floorCol As Long, roomCol As Long, ageCol As Long, inspectionCol As Long
    Dim floorRankCol As Long, roomRankCol As Long
    Dim i As Long, descText As String, floorVal As String, roomVal As String
    Dim regex As Object, matches As Object
    Dim createdDate As Variant
    Dim stage As String

    On Error GoTo FailHard
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    stage = "Get ActiveSheet"
    Set ws = ActiveSheet
    If ws Is Nothing Then GoTo CleanExit

    ' --- Identify headers safely ---
    stage = "Find headers"
    descCol = FindHeader(ws, "description", "desc")
    propCol = FindHeader(ws, "property", "building", "bldg", "building code")
    dateCreatedCol = FindHeader(ws, "date created", "created", "created date", "date_created")

    If descCol = 0 Then
        Err.Raise vbObjectError + 1000, , "Header 'Description' not found."
    End If

    ' Ensure single Age (Days)
    DeleteDuplicateAgeColumns ws
    ageCol = FindHeader(ws, "age (days)", "age", "age days")
    If ageCol = 0 Then ageCol = AppendColumn(ws, "Age (Days)")

    ' Ensure Floor, Room, Inspection
    floorCol = FindHeader(ws, "floor"): If floorCol = 0 Then floorCol = AppendColumn(ws, "Floor")
    roomCol = FindHeader(ws, "room"): If roomCol = 0 Then roomCol = AppendColumn(ws, "Room")
    inspectionCol = FindHeader(ws, "inspection status", "inspection"): If inspectionCol = 0 Then inspectionCol = AppendColumn(ws, "Inspection Status")

    HeaderShade ws, Array(floorCol, roomCol, ageCol, inspectionCol)

    ' Hidden sort helpers
    floorRankCol = AppendColumn(ws, "__FloorRank")
    roomRankCol = AppendColumn(ws, "__RoomRank")

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    lastRow = ws.Cells(ws.Rows.Count, descCol).End(xlUp).row
    If lastRow < 2 Then GoTo AfterLoop

    ' === MAIN LOOP ===
    For i = 2 To lastRow
        descText = NzStr(ws.Cells(i, descCol).Value)
        floorVal = "": roomVal = ""

        regex.Pattern = "(Flr|Floor)\s*:\s*([A-Za-z0-9]+)"
        If regex.Test(descText) Then
            Set matches = regex.Execute(descText)
            floorVal = Trim(matches(0).SubMatches(1))
        End If

        regex.Pattern = "(Rm|Room)\s*:\s*([A-Za-z0-9\-]+)"
        If regex.Test(descText) Then
            Set matches = regex.Execute(descText)
            roomVal = Trim(matches(0).SubMatches(1))
        End If

        ws.Cells(i, floorCol).Value = NormalizeFloor(floorVal, roomVal)
        ws.Cells(i, roomCol).Value = roomVal

        ' === Age calculation ===
        If dateCreatedCol > 0 Then
            createdDate = ws.Cells(i, dateCreatedCol).Value
            If IsDate(createdDate) Then
                ws.Cells(i, ageCol).Value = CLng(Application.WorksheetFunction.NetworkDays_Intl(CDate(createdDate), Date, 1))
            Else
                ws.Cells(i, ageCol).Value = vbNullString
            End If
        Else
            ws.Cells(i, ageCol).Value = vbNullString
        End If

        ' Default Inspection
        ' --- Only set Pending for empty cells to preserve existing user input ---
        If Trim(NzStr(ws.Cells(i, inspectionCol).Value)) = "" Then
            ws.Cells(i, inspectionCol).Value = "Pending"
        End If


        ws.Cells(i, floorRankCol).Value = FloorRankVal(ws.Cells(i, floorCol).Value)
        ws.Cells(i, roomRankCol).Value = RoomRankVal(ws.Cells(i, roomCol).Value)
    Next i


AfterLoop:
    ' === Sort ===
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Columns(floorRankCol), Order:=xlAscending
    ws.Sort.SortFields.Add key:=ws.Columns(roomRankCol), Order:=xlAscending
    With ws.Sort
        .SetRange ws.UsedRange
        .Header = xlYes
        .Apply
    End With

    ' Delete helpers
    Application.DisplayAlerts = False
    ws.Columns(roomRankCol).Delete
    ws.Columns(floorRankCol).Delete
    Application.DisplayAlerts = True

    ' Dropdown for Inspection
    If inspectionCol > 0 And lastRow >= 2 Then
        With ws.Range(ws.Cells(2, inspectionCol), ws.Cells(lastRow, inspectionCol)).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="Pending,Complete,Incomplete,Needs Review"
        End With
    End If

    ' === Ensure a spacer between Inspection Status and Age (Days), moving Age right by 1 ===
    ' Re-detect current columns first
    inspectionCol = FindHeader(ws, "inspection status", "inspection")
    ageCol = FindHeader(ws, "age (days)", "age", "age days")
    
    If ageCol > 0 And inspectionCol > 0 Then
        ' Only insert a spacer if Age is immediately to the right of Inspection Status
        If ageCol = inspectionCol + 1 Then
            ' Insert a blank column at the current Age position (this pushes Age one column to the right)
            ws.Columns(ageCol).Insert Shift:=xlToRight
            ' After insert, Age (Days) is now at ageCol + 1
            ageCol = ageCol + 1
        End If
    End If
    
    ' Make sure Age (Days) has NO validation (Excel sometimes drags it along)
    If ageCol > 0 And lastRow >= 2 Then
        On Error Resume Next
        ws.Range(ws.Cells(2, ageCol), ws.Cells(lastRow, ageCol)).Validation.Delete
        On Error GoTo 0
    End If
    
    ' Reapply Age colors now that its position is final
    If ageCol > 0 And lastRow >= 2 Then
        ApplyAgeColors ws, ageCol, lastRow
    End If
    
    ' Reapply Inspection Status dropdown (in case the insert disturbed validation range)
    inspectionCol = FindHeader(ws, "inspection status", "inspection")
    If inspectionCol > 0 And lastRow >= 2 Then
        With ws.Range(ws.Cells(2, inspectionCol), ws.Cells(lastRow, inspectionCol)).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="Pending,Complete,Incomplete,Needs Review"
            .InCellDropdown = True
        End With
    End If
        
        ' === Remove any data validation from Age (Days) ===
    If ageCol > 0 Then
        On Error Resume Next
        ws.Range(ws.Cells(2, ageCol), ws.Cells(lastRow, ageCol)).Validation.Delete
        On Error GoTo 0
    End If
    
    ' === Layout & Borders ===
    With ws.UsedRange
        .Font.name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Columns(descCol)
        .ColumnWidth = 50
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
    
    ws.Rows(1).RowHeight = 20
    ws.Cells.EntireColumn.AutoFit

' === Reapply dropdown validation to Inspection Status only ===
If inspectionCol > 0 And lastRow >= 2 Then
    With ws.Range(ws.Cells(2, inspectionCol), ws.Cells(lastRow, inspectionCol)).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Pending,Complete,Incomplete,Needs Review"
        .InCellDropdown = True
    End With
End If

' === Apply formatting ===
If ageCol > 0 And lastRow >= 2 Then ApplyAgeColors ws, ageCol, lastRow
AddInspectionFormatting ws, inspectionCol
If ENABLE_DASHBOARD Then CreateDashboard ws.Parent, ws.name


    ' === Colors ===
    If ageCol > 0 And lastRow >= 2 Then ApplyAgeColors ws, ageCol, lastRow
    AddInspectionFormatting ws, inspectionCol
    If ENABLE_DASHBOARD Then CreateDashboard ws.Parent, ws.name

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

FailHard:
    MsgBox "Stage: " & stage & vbCrLf & "Row: " & i & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical, "AIM Formatter"
    Resume CleanExit
End Sub

'===========================================================
' AGE COLORING
'===========================================================
Private Sub ApplyAgeColors(ByVal ws As Worksheet, ByVal ageCol As Long, ByVal lastRow As Long)
    Dim c As Range, v As Variant
    Dim overdueRed As Long: overdueRed = RGB(255, 153, 153)
    Dim gR As Double, gG As Double, gB As Double
    Dim rR As Double, rG As Double, rB As Double
    Dim frac As Double, rMix As Double, gMix As Double, bMix As Double

    gR = 198: gG = 239: gB = 206
    rR = 255: rG = 153: rB = 153

    ws.Range(ws.Cells(2, ageCol), ws.Cells(lastRow, ageCol)).Interior.ColorIndex = xlNone
    ws.Columns(ageCol).NumberFormat = "0"

    For Each c In ws.Range(ws.Cells(2, ageCol), ws.Cells(lastRow, ageCol))
        v = c.Value
        If IsNumeric(v) Then
            If v >= 30 Then
                c.Interior.Color = overdueRed
            ElseIf v >= 1 And v <= 29 Then
                frac = (CDbl(v) - 1#) / 28#
                rMix = gR + (rR - gR) * frac
                gMix = gG + (rG - gG) * frac
                bMix = gB + (rB - gB) * frac
                c.Interior.Color = RGB(CLng(rMix), CLng(gMix), CLng(bMix))
            End If
        End If
    Next c
End Sub

'===========================================================
' INSPECTION STATUS FORMATTING
'===========================================================
Public Sub AddInspectionFormatting(ws As Worksheet, inspectionCol As Long)
    Dim lastRow As Long, lastCol As Long, ageCol As Long
    Dim rng1 As Range, rng2 As Range, shadeRange As Range

    If inspectionCol = 0 Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, inspectionCol).End(xlUp).row
    If lastRow < 2 Then Exit Sub
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    ageCol = FindHeader(ws, "age (days)", "age", "age days")

    If ageCol > 1 Then Set rng1 = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ageCol - 1))
    If ageCol > 0 And ageCol < lastCol Then Set rng2 = ws.Range(ws.Cells(2, ageCol + 1), ws.Cells(lastRow, lastCol))
    If Not rng1 Is Nothing And Not rng2 Is Nothing Then
        Set shadeRange = Union(rng1, rng2)
    ElseIf Not rng1 Is Nothing Then
        Set shadeRange = rng1
    ElseIf Not rng2 Is Nothing Then
        Set shadeRange = rng2
    Else
        Exit Sub
    End If

    shadeRange.FormatConditions.Delete
    With shadeRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & Split(ws.Cells(1, inspectionCol).Address, "$")(1) & "2=""Complete""")
        .Interior.Color = RGB(198, 239, 206)
    End With
    With shadeRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & Split(ws.Cells(1, inspectionCol).Address, "$")(1) & "2=""Incomplete""")
        .Interior.Color = RGB(255, 199, 206)
    End With
    With shadeRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & Split(ws.Cells(1, inspectionCol).Address, "$")(1) & "2=""Needs Review""")
        .Interior.Color = RGB(255, 235, 156)
    End With
End Sub

'===========================================================
' DASHBOARD CREATION
'===========================================================
Public Sub CreateDashboard(wb As Workbook, dataSheetName As String)
    Dim dashWS As Worksheet, dataWS As Worksheet
    Dim lastRow As Long
    Dim woCol As Long, inspectionCol As Long, ageCol As Long
    Dim woColL As String, inspColL As String, ageColL As String
    Dim dashName As String

    On Error Resume Next
    Set dataWS = wb.Worksheets(dataSheetName)
    On Error GoTo 0
    If dataWS Is Nothing Then Exit Sub

    woCol = FindHeader(dataWS, "work order", "workorder", "wo")
    inspectionCol = FindHeader(dataWS, "inspection status", "inspection")
    ageCol = FindHeader(dataWS, "age (days)", "age", "age days")
    If woCol = 0 Then Exit Sub

    Application.DisplayAlerts = False
    If WorksheetExists("Dashboard", wb) Then wb.Worksheets("Dashboard").Delete
    Application.DisplayAlerts = True

    Set dashWS = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    dashWS.name = "Dashboard"

    With dashWS.Range("A1:F1")
        .Merge
        .Value = "Work Order Analytics Dashboard"
        .Font.Size = 18
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 30
    End With

    lastRow = dataWS.Cells(dataWS.Rows.Count, woCol).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    woColL = ColumnLetter(woCol)
    inspColL = ColumnLetter(inspectionCol)
    ageColL = ColumnLetter(ageCol)

    dashWS.Range("A3").Value = "SUMMARY STATISTICS"
    dashWS.Range("A3").Font.Bold = True
    dashWS.Range("A3").Font.Size = 14

    SafeFormula dashWS, "B5", "=COUNTA('" & dataSheetName & "'!" & woColL & ":" & woColL & ")-1"
    SafeFormula dashWS, "B6", "=COUNTIF('" & dataSheetName & "'!" & inspColL & ":" & inspColL & ",""Pending"")"
    SafeFormula dashWS, "B7", "=COUNTIF('" & dataSheetName & "'!" & inspColL & ":" & inspColL & ",""Complete"")"
    SafeFormula dashWS, "B8", "=COUNTIF('" & dataSheetName & "'!" & inspColL & ":" & inspColL & ",""Incomplete"")"
    SafeFormula dashWS, "B9", "=COUNTIF('" & dataSheetName & "'!" & inspColL & ":" & inspColL & ",""Needs Review"")"

    dashWS.Range("D3").Value = "AGE ANALYSIS"
    dashWS.Range("D3").Font.Bold = True
    dashWS.Range("D3").Font.Size = 14

    SafeFormula dashWS, "E5", "=IFERROR(ROUND(AVERAGE('" & dataSheetName & "'!" & ageColL & ":" & ageColL & "),1),""n/a"")"
    SafeFormula dashWS, "E6", "=IFERROR(COUNTIF('" & dataSheetName & "'!" & ageColL & ":" & ageColL & ","">=30""),""n/a"")"

    dashWS.Columns("A:F").AutoFit
    dashWS.Range("A5:B9").Borders.LineStyle = xlContinuous
    dashWS.Range("D5:E6").Borders.LineStyle = xlContinuous
End Sub

Private Sub SafeFormula(ws As Worksheet, ByVal addr As String, ByVal formulaText As String)
    On Error Resume Next
    ws.Range(addr).Formula = formulaText
    If Err.Number <> 0 Then ws.Range(addr).Value = "n/a": Err.Clear
    On Error GoTo 0
End Sub

Private Function WorksheetExists(name As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

'===========================================================
' HELPERS
'===========================================================
Private Function FindHeader(ByVal ws As Worksheet, ByVal headerName As String, ParamArray aliases()) As Long
    Dim i As Long, got As String, a As Variant, lastCol As Long
    On Error Resume Next
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1

    For i = 1 To lastCol
        On Error Resume Next
        got = CleanHeader(ws.Cells(1, i).Text)
        On Error GoTo 0
        If got = CleanHeader(headerName) Then FindHeader = i: Exit Function
        For Each a In aliases
            If got = CleanHeader(a) Then FindHeader = i: Exit Function
        Next a
    Next i
End Function

Private Function CleanHeader(ByVal s As String) As String
    CleanHeader = LCase(Trim(Replace(Replace(Replace(CStr(s), vbCr, ""), vbLf, ""), vbTab, "")))
End Function

Private Function AppendColumn(ByVal ws As Worksheet, ByVal title As String) As Long
    Dim col As Long
    col = ws.UsedRange.Columns.Count + 1
    ws.Cells(1, col).Value = title
    AppendColumn = col
End Function

Private Sub HeaderShade(ByVal ws As Worksheet, ByVal cols As Variant)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If cols(i) > 0 Then
            With ws.Cells(1, cols(i))
                .Interior.Color = RGB(200, 200, 200)
                .Font.Bold = True
            End With
        End If
    Next i
End Sub

Public Function NormalizeFloor(ByVal floorVal As String, ByVal roomVal As String) As String
    Dim f As String
    f = UCase(Trim(floorVal))
    If f = "" And Len(roomVal) > 0 Then
        Select Case Left(roomVal, 1)
            Case "0": f = "B"
            Case "1" To "9": f = Left(roomVal, 1)
        End Select
    End If
    Select Case f
        Case "0", "B", "BASEMENT": NormalizeFloor = "B"
        Case "SF", "ROOF": NormalizeFloor = "SF"
        Case Else: NormalizeFloor = f
    End Select
End Function

Public Function FloorRankVal(ByVal f As Variant) As Long
    Dim t As String
    t = UCase(Trim(CStr(f)))
    Select Case True
        Case t = "B": FloorRankVal = 0
        Case t = "SF": FloorRankVal = 99
        Case IsNumeric(t): FloorRankVal = Val(t)
        Case Else: FloorRankVal = 999
    End Select
End Function

Public Function RoomRankVal(ByVal r As Variant) As Long
    Dim t As String, j As Long, ch As String, numPrefix As String
    t = UCase(Trim(CStr(r)))
    If t = "" Then RoomRankVal = 999999: Exit Function
    numPrefix = ""
    For j = 1 To Len(t)
        ch = Mid$(t, j, 1)
        If ch Like "[0-9]" Then numPrefix = numPrefix & ch Else Exit For
    Next j
    If Len(numPrefix) > 0 Then RoomRankVal = CLng(numPrefix) Else RoomRankVal = 600000
End Function

Public Function NzStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Public Function ColumnLetter(colNum As Long) As String
    If colNum < 1 Then ColumnLetter = "" Else ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

Private Sub DeleteDuplicateAgeColumns(ByVal ws As Worksheet)
    Dim i As Long, found As Boolean, hdr As String
    found = False
    For i = ws.UsedRange.Columns.Count To 1 Step -1
        hdr = CleanHeader(ws.Cells(1, i).Value)
        Select Case hdr
            Case "age (days)", "age days", "age"
                If found Then ws.Columns(i).Delete Else found = True
        End Select
    Next i
End Sub


