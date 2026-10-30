Attribute VB_Name = "SalesConsolidator"
'==============================================================================
' Sales Consolidator - Excel VBA
'
' One click turns many monthly sheets (or many workbooks in a folder) into one
' clean Master table plus a Summary:
'   1. ConsolidateSheets  - stacks every data sheet with the expected header
'   2. ConsolidateFolder  - same, but reads every .xlsx/.xlsm/.xls in a folder
'   3. CleanMaster        - trims, fixes casing, converts text dates/numbers,
'                           flags bad rows, removes duplicate invoice numbers
'   4. BuildSummary       - revenue by region and by product + a chart
'   RunAll                - 1 + 3 + 4 in one go (assign it to a button)
'
' Everything is processed in memory arrays, so 100k+ rows take seconds,
' not minutes. No references to set: Scripting.Dictionary is late-bound.
'==============================================================================
Option Explicit

Private Const MASTER_SHEET As String = "Master"
Private Const SUMMARY_SHEET As String = "Summary"
Private Const ISSUES_SHEET As String = "Issues"
Private Const CONTROL_SHEET As String = "Control"
Private Const N_COLS As Long = 7

' Expected header of every source sheet (column order matters)
Private Function SourceHeader() As Variant
    SourceHeader = Array("Date", "Region", "Rep", "Product", "Units", "Unit Price", "Invoice No")
End Function

'------------------------------------------------------------------------------
Public Sub RunAll()
    Dim t As Double: t = Timer
    On Error GoTo Fail
    DoConsolidateSheets False
    DoCleanMaster False
    BuildSummary
    MsgBox "Done in " & Format(Timer - t, "0.0") & " s." & vbCrLf & _
           RowCount(MASTER_SHEET) & " clean rows, " & RowCount(ISSUES_SHEET) & " issues to review.", _
           vbInformation, "Sales Consolidator"
    Exit Sub
Fail:
    SpeedUp False
    MsgBox "Stopped: " & Err.Description, vbExclamation, "Sales Consolidator"
End Sub

Public Sub ConsolidateSheets()
    DoConsolidateSheets True
End Sub

Public Sub CleanMaster()
    On Error GoTo Fail
    DoCleanMaster True
    Exit Sub
Fail:
    SpeedUp False
    MsgBox "Stopped: " & Err.Description, vbExclamation, "Sales Consolidator"
End Sub

'------------------------------------------------------------------------------
' 1. Stack every sheet of THIS workbook whose header matches SourceHeader
'------------------------------------------------------------------------------
Private Sub DoConsolidateSheets(ByVal showMessage As Boolean)
    Dim ws As Worksheet, out As Collection
    Set out = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If Not IsOutputSheet(ws.Name) And HeaderMatches(ws) Then
            AppendSheet ws, ws.Name, out
        End If
    Next ws
    WriteMaster out
    If showMessage Then MsgBox out.Count & " rows consolidated into '" & MASTER_SHEET & "'.", vbInformation
End Sub

'------------------------------------------------------------------------------
' 2. Stack every matching sheet of every workbook in a chosen folder
'------------------------------------------------------------------------------
Public Sub ConsolidateFolder()
    Dim folder As String, f As String, wb As Workbook, ws As Worksheet
    Dim out As Collection, nFiles As Long

    With Application.FileDialog(4)          ' msoFileDialogFolderPicker
        .Title = "Choose the folder with the monthly files"
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator

    Set out = New Collection
    SpeedUp True
    On Error GoTo Fail
    f = Dir(folder & "*.xls*")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" And f <> ThisWorkbook.Name Then
            Set wb = Workbooks.Open(folder & f, ReadOnly:=True, UpdateLinks:=0)
            For Each ws In wb.Worksheets
                If Not IsOutputSheet(ws.Name) And HeaderMatches(ws) Then AppendSheet ws, f & " / " & ws.Name, out
            Next ws
            wb.Close SaveChanges:=False
            Set wb = Nothing
            nFiles = nFiles + 1
        End If
        f = Dir()
    Loop
    SpeedUp False
    WriteMaster out
    MsgBox out.Count & " rows from " & nFiles & " file(s) consolidated.", vbInformation
    Exit Sub
Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    SpeedUp False
    MsgBox "Stopped at '" & f & "': " & msg, vbExclamation
End Sub

'------------------------------------------------------------------------------
' 3. Clean the Master sheet in place; problems go to the Issues sheet
'------------------------------------------------------------------------------
Private Sub DoCleanMaster(ByVal showMessage As Boolean)
    Dim ws As Worksheet, data As Variant, n As Long, i As Long, k As Long
    Dim cleaned() As Variant, issues As Collection, seen As Object
    Dim d As Variant, units As Variant, price As Variant, inv As String, problem As String
    Dim kept As Long

    Set ws = GetSheet(MASTER_SHEET)
    n = LastUsedRow(ws) - 1
    If n < 1 Then
        If showMessage Then MsgBox "Master is empty - run ConsolidateSheets first.", vbExclamation
        Exit Sub
    End If

    SpeedUp True
    data = ws.Range("A2").Resize(n, N_COLS + 1).Value      ' +1 = Source column
    ReDim cleaned(1 To n, 1 To N_COLS + 2)                 ' + Source + Revenue
    Set issues = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For i = 1 To n
        problem = ""
        d = ToDate(data(i, 1))
        If IsEmpty(d) Then problem = AddProblem(problem, "invalid date")
        units = ToNumber(data(i, 5))
        If IsEmpty(units) Then
            problem = AddProblem(problem, "units not a number")
        ElseIf units <= 0 Then
            problem = AddProblem(problem, "units must be > 0")
        End If
        price = ToNumber(data(i, 6))
        If IsEmpty(price) Then
            problem = AddProblem(problem, "price not a number")
        ElseIf price <= 0 Then
            problem = AddProblem(problem, "price must be > 0")
        End If
        inv = UCase$(Squash(data(i, 7)))
        If Len(inv) = 0 Then problem = AddProblem(problem, "missing invoice no")
        If Len(Squash(data(i, 2))) = 0 Then problem = AddProblem(problem, "missing region")
        If Len(Squash(data(i, 4))) = 0 Then problem = AddProblem(problem, "missing product")
        If Len(problem) = 0 Then
            If seen.Exists(inv) Then problem = "duplicate invoice (first kept: " & seen(inv) & ")"
        End If

        If Len(problem) > 0 Then
            issues.Add Array(i + 1, data(i, N_COLS + 1), data(i, 1), data(i, 2), data(i, 3), _
                             data(i, 4), data(i, 5), data(i, 6), data(i, 7), problem)
        Else
            seen.Add inv, CStr(data(i, N_COLS + 1))
            kept = kept + 1
            cleaned(kept, 1) = d
            cleaned(kept, 2) = ProperName(data(i, 2))
            cleaned(kept, 3) = ProperName(data(i, 3))
            cleaned(kept, 4) = Squash(data(i, 4))
            cleaned(kept, 5) = units
            cleaned(kept, 6) = price
            cleaned(kept, 7) = inv
            cleaned(kept, 8) = data(i, N_COLS + 1)
            cleaned(kept, 9) = Application.WorksheetFunction.Round(units * price, 2)
        End If
    Next i

    ' write back
    ws.Cells.Clear
    WriteHeader ws, Array("Date", "Region", "Rep", "Product", "Units", "Unit Price", "Invoice No", "Source", "Revenue")
    If kept > 0 Then
        ws.Range("A2").Resize(kept, N_COLS + 2).Value = cleaned  ' extra empty rows are not written
        ws.Range("A2").Resize(kept, 1).NumberFormat = "yyyy-mm-dd"
        ws.Range("F2").Resize(kept, 1).NumberFormat = "#,##0.00"
        ws.Range("I2").Resize(kept, 1).NumberFormat = "#,##0.00"
    End If
    FinishTable ws

    Dim wi As Worksheet: Set wi = GetSheet(ISSUES_SHEET)
    wi.Cells.Clear
    WriteHeader wi, Array("Consolidated row", "Source sheet & row", "Date", "Region", "Rep", "Product", "Units", "Unit Price", "Invoice No", "Problem")
    If issues.Count > 0 Then
        Dim arr() As Variant, j As Long
        ReDim arr(1 To issues.Count, 1 To 10)
        Dim rowItem As Variant
        For k = 1 To issues.Count
            rowItem = issues(k)
            For j = 0 To 9
                arr(k, j + 1) = rowItem(j)
            Next j
        Next k
        wi.Range("A2").Resize(issues.Count, 10).Value = arr
        wi.Range("J2").Resize(issues.Count, 1).Font.Color = RGB(176, 0, 32)
    End If
    FinishTable wi
    SpeedUp False

    If showMessage Then MsgBox kept & " clean rows, " & issues.Count & " sent to '" & ISSUES_SHEET & "'.", vbInformation
End Sub

'------------------------------------------------------------------------------
' 4. Revenue by region and by product, with a chart
'------------------------------------------------------------------------------
Public Sub BuildSummary()
    Dim ws As Worksheet, data As Variant, n As Long, i As Long
    Dim byRegion As Object, byProduct As Object, total As Double
    Dim key As Variant, r As Long, sumWs As Worksheet

    Set ws = GetSheet(MASTER_SHEET)
    n = LastUsedRow(ws) - 1
    If n < 1 Then Exit Sub
    data = ws.Range("A2").Resize(n, 9).Value

    Set byRegion = CreateObject("Scripting.Dictionary")
    Set byProduct = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        byRegion(data(i, 2)) = byRegion(data(i, 2)) + data(i, 9)
        byProduct(data(i, 4)) = byProduct(data(i, 4)) + data(i, 9)
        total = total + data(i, 9)
    Next i

    Set sumWs = GetSheet(SUMMARY_SHEET)
    sumWs.Cells.Clear
    Dim co As ChartObject
    Do While sumWs.ChartObjects.Count > 0
        sumWs.ChartObjects(1).Delete
    Loop

    sumWs.Range("A1").Value = "Sales Summary"
    sumWs.Range("A1").Font.Size = 14: sumWs.Range("A1").Font.Bold = True
    sumWs.Range("A2").Value = "Rows: " & n & "   |   Total revenue: " & Format(total, "#,##0.00") & _
                              "   |   Updated: " & Format(Now, "yyyy-mm-dd hh:nn")

    WriteHeaderAt sumWs.Range("A4"), Array("Region", "Revenue")
    r = 5
    For Each key In byRegion.Keys
        sumWs.Cells(r, 1).Value = key
        sumWs.Cells(r, 2).Value = Round(byRegion(key), 2)
        r = r + 1
    Next key
    SortBlock sumWs.Range("A5").Resize(byRegion.Count, 2)
    sumWs.Range("B5").Resize(byRegion.Count, 1).NumberFormat = "#,##0.00"

    WriteHeaderAt sumWs.Range("D4"), Array("Product", "Revenue")
    r = 5
    For Each key In byProduct.Keys
        sumWs.Cells(r, 4).Value = key
        sumWs.Cells(r, 5).Value = Round(byProduct(key), 2)
        r = r + 1
    Next key
    SortBlock sumWs.Range("D5").Resize(byProduct.Count, 2)
    sumWs.Range("E5").Resize(byProduct.Count, 1).NumberFormat = "#,##0.00"

    Set co = sumWs.ChartObjects.Add(Left:=sumWs.Range("G4").Left, Top:=sumWs.Range("G4").Top, Width:=420, Height:=250)
    With co.Chart
        .ChartType = xlBarClustered
        .SetSourceData Source:=sumWs.Range("D4").Resize(byProduct.Count + 1, 2), PlotBy:=xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Revenue by product"
        .HasLegend = False
    End With
    sumWs.Range("A4").CurrentRegion.Columns.AutoFit
    sumWs.Range("D4").CurrentRegion.Columns.AutoFit
    sumWs.Parent.Activate
    sumWs.Activate
End Sub

'==============================================================================
' Helpers
'==============================================================================
Private Sub AppendSheet(ByVal ws As Worksheet, ByVal sourceName As String, ByVal out As Collection)
    Dim last As Long, data As Variant, i As Long, j As Long, rowArr() As Variant, blank As Boolean
    last = LastUsedRow(ws)
    If last < 2 Then Exit Sub
    data = ws.Range("A2").Resize(last - 1, N_COLS).Value
    For i = 1 To UBound(data, 1)
        blank = True
        ReDim rowArr(1 To N_COLS + 1)
        For j = 1 To N_COLS
            rowArr(j) = data(i, j)
            If Len(Squash(data(i, j))) > 0 Then blank = False
        Next j
        rowArr(N_COLS + 1) = sourceName & " row " & (i + 1)
        If Not blank Then out.Add rowArr
    Next i
End Sub

Private Sub WriteMaster(ByVal items As Collection)
    Dim ws As Worksheet, arr() As Variant, i As Long, j As Long
    Set ws = GetSheet(MASTER_SHEET)
    ws.Cells.Clear
    WriteHeader ws, Array("Date", "Region", "Rep", "Product", "Units", "Unit Price", "Invoice No", "Source")
    If items.Count = 0 Then Exit Sub
    Dim rowItem As Variant
    ReDim arr(1 To items.Count, 1 To N_COLS + 1)
    For i = 1 To items.Count
        rowItem = items(i)
        For j = 1 To N_COLS + 1
            ' Store text as text: otherwise Excel re-parses "07/01/2024" with US
            ' month/day rules on write. CleanMaster does the real conversion.
            If VarType(rowItem(j)) = vbString Then arr(i, j) = "'" & rowItem(j) Else arr(i, j) = rowItem(j)
        Next j
    Next i
    ws.Range("A2").Resize(items.Count, N_COLS + 1).Value = arr
    FinishTable ws
End Sub

Private Function HeaderMatches(ByVal ws As Worksheet) As Boolean
    Dim h As Variant, j As Long
    h = SourceHeader()
    For j = 0 To N_COLS - 1
        If StrComp(Squash(ws.Cells(1, j + 1).Value), h(j), vbTextCompare) <> 0 Then Exit Function
    Next j
    HeaderMatches = True
End Function

Private Function IsOutputSheet(ByVal sheetName As String) As Boolean
    Select Case LCase$(sheetName)
        Case LCase$(MASTER_SHEET), LCase$(SUMMARY_SHEET), LCase$(ISSUES_SHEET), LCase$(CONTROL_SHEET)
            IsOutputSheet = True
    End Select
End Function

' Accepts real Excel dates, ISO text (2024-03-05), d/m/yyyy text and "5 Mar 2024".
' Returns Empty when the value is not a real calendar date.
Private Function ToDate(ByVal v As Variant) As Variant
    Dim s As String, p() As String, y As Long, m As Long, d As Long
    If IsDate(v) And VarType(v) = vbDate Then ToDate = CDate(v): Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then
        If v > 20000 And v < 80000 Then ToDate = CDate(v): Exit Function   ' serial date stored as number
    End If
    s = Squash(v)
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ".", "/")
    If s Like "####-#*-#*" Or s Like "####/#*/#*" Then
        p = Split(Replace(s, "-", "/"), "/")
        If UBound(p) <> 2 Then Exit Function
        y = Val(p(0)): m = Val(p(1)): d = Val(p(2))
    ElseIf s Like "#*/#*/####" Then
        p = Split(s, "/")                          ' day/month/year (international)
        If UBound(p) <> 2 Then Exit Function
        d = Val(p(0)): m = Val(p(1)): y = Val(p(2))
    ElseIf s Like "#* [A-Za-z][A-Za-z][A-Za-z]* ####" Then
        p = Split(s, " ")
        If UBound(p) <> 2 Then Exit Function
        d = Val(p(0)): y = Val(p(2))
        Dim pos As Long
        pos = InStr(1, "janfebmaraprmayjunjulaugsepoctnovdec", LCase$(Left$(p(1), 3)))
        If pos = 0 Or (pos - 1) Mod 3 <> 0 Then Exit Function
        m = (pos - 1) \ 3 + 1
    Else
        Exit Function
    End If
    If y < 1900 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    If Day(DateSerial(y, m, d)) <> d Then Exit Function     ' 31/02 would roll into March
    ToDate = DateSerial(y, m, d)
End Function

' "$1,234.50", "USD 99", " 12 " -> Double. Empty when not a number.
Private Function ToNumber(ByVal v As Variant) As Variant
    Dim s As String, i As Long, c As String, out As String
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        ToNumber = CDbl(v): Exit Function
    End If
    s = Squash(v)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "[0-9]" Or c = "." Or c = "-" Then out = out & c
    Next i
    If Len(out) = 0 Or out = "." Or out = "-" Then Exit Function
    If InStr(2, out, "-") > 0 Then Exit Function                         ' minus only at the start
    If Len(out) - Len(Replace(out, ".", "")) > 1 Then Exit Function      ' at most one decimal point
    ' Val() always uses "." as the decimal point, whatever the PC's regional settings
    ToNumber = Val(out)
End Function

Private Function Squash(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = Trim$(Replace(Replace(CStr(v), vbTab, " "), Chr$(160), " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Squash = s
End Function

Private Function ProperName(ByVal v As Variant) As String
    ProperName = StrConv(Squash(v), vbProperCase)
End Function

Private Function AddProblem(ByVal current As String, ByVal msg As String) As String
    If Len(current) = 0 Then AddProblem = msg Else AddProblem = current & "; " & msg
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Set GetSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSheet.Name = sheetName
    End If
End Function

Private Sub WriteHeader(ByVal ws As Worksheet, ByVal names As Variant)
    WriteHeaderAt ws.Range("A1"), names
End Sub

Private Sub WriteHeaderAt(ByVal topLeft As Range, ByVal names As Variant)
    With topLeft.Resize(1, UBound(names) - LBound(names) + 1)
        .Value = names
        .Font.Bold = True
        .Interior.Color = RGB(232, 234, 237)
    End With
End Sub

Private Sub FinishTable(ByVal ws As Worksheet)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1").CurrentRegion.AutoFilter
    ws.Columns.AutoFit
    ws.Parent.Activate
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

Private Sub SortBlock(ByVal rng As Range)
    If rng.Rows.Count < 2 Then Exit Sub
    rng.Sort Key1:=rng.Columns(2), Order1:=xlDescending, Header:=xlNo
End Sub

Private Function RowCount(ByVal sheetName As String) As Long
    Dim ws As Worksheet: Set ws = GetSheet(sheetName)
    RowCount = Application.WorksheetFunction.Max(0, LastUsedRow(ws) - 1)
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then LastUsedRow = 0 Else LastUsedRow = c.Row
End Function

Private Sub SpeedUp(ByVal turnOn As Boolean)
    Application.ScreenUpdating = Not turnOn
    Application.EnableEvents = Not turnOn
    Application.Calculation = IIf(turnOn, xlCalculationManual, xlCalculationAutomatic)
End Sub
