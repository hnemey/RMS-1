Attribute VB_Name = "modGAAPActualForecast"
Option Explicit

'==========================================================================
' GAAP tab - "what is actual, what is forecast" formatter
'--------------------------------------------------------------------------
' Every month-end one more column of the FY block moves from forecast to
' actual.  This module does that re-formatting in one step:
'
'   * moves the "FY26 Actual" / "FY26 Forecast" banner split in row 3 and
'     re-labels each banner with the months it covers
'   * colour-codes every FY column - closed months, forecast months, and the
'     quarter / total columns that are a MIX of the two
'   * shades the month header row to match, so the split stays visible when
'     the sheet is scrolled
'   * clears the old vertical divider and draws a new heavy one on the
'     current actual / forecast boundary
'   * writes a small colour key above the block
'
' Macros
'   FormatGAAP_ActualVsForecast     format the GAAP tab (asks for the month)
'   FormatAllTabs_ActualVsForecast  same, for every tab in TARGET_SHEETS
'   ResetGAAP_Formatting            put the GAAP tab back the way it was
'   ResetAllTabs_Formatting         same, for every tab in TARGET_SHEETS
'
' Safety
'   The first time a tab is formatted, its shading and vertical borders are
'   copied to a very-hidden backup sheet ("_fmtbak_<tab>").  Every run
'   restores that baseline before painting, so repeated runs never stack up
'   and the Reset macros return the tab to its pre-macro look.  Formulas,
'   values, fonts, number formats and row/column sizes are never touched.
'==========================================================================


'---------------------- Settings you may want to change -------------------

' Tabs handled by the "AllTabs" macros (comma separated).
Private Const TARGET_SHEETS   As String = "GAAP,Combined"

' The tab handled by the plain GAAP macros.
Private Const MAIN_SHEET      As String = "GAAP"

' Fiscal year being tracked.  Used to find the block and to build the banner
' text - change to "FY27" next October.
Private Const FY_LABEL        As String = "FY26"

' Sheet layout (identical on GAAP and Combined).
Private Const BANNER_ROW      As Long = 3     ' "FY26 Actual" / "FY26 Forecast"
Private Const MONTH_ROW       As Long = 4     ' Oct, Nov, Dec, Q1, ... , Total
Private Const LEGEND_ROW      As Long = 2     ' blank row above the banner
Private Const MAX_SCAN_ROW    As Long = 200   ' safety stop when finding the last data row

' Set to False to skip the colour key above the block.
Private Const SHOW_LEGEND     As Boolean = True

' Cells that already carried a highlight before the macro ran (manual
' call-outs and the like) keep their colour.  False = paint everything.
Private Const KEEP_HIGHLIGHTS As Boolean = True

' Put the covered months in the banner, e.g. "FY26 Actual (Oct - Jun)".
Private Const BANNER_SHOWS_RANGE As Boolean = True

' Draw a medium line after each quarter column in the two header rows.
Private Const QUARTER_LINES   As Boolean = True

' Leave False when the module lives in the forecast workbook itself.
' Set to True if you keep it in PERSONAL.XLSB and run it against whichever
' workbook is in front.
Private Const USE_ACTIVE_WORKBOOK As Boolean = False

Private Const BACKUP_PREFIX   As String = "_fmtbak_"


'---------------------- Internals -----------------------------------------

Private Const ST_ACTUAL   As Long = 1
Private Const ST_FORECAST As Long = 2
Private Const ST_MIXED    As Long = 3

Private Type TBlock
    Found        As Boolean
    StartCol     As Long        ' first month column of the FY block (Oct)
    EndCol       As Long        ' the "Total" column of the FY block
    LastRow      As Long        ' last row of the statement body
    FirstMonth() As Long        ' per column: first fiscal month covered (0 = spacer)
    LastMonth()  As Long        ' per column: last fiscal month covered  (0 = spacer)
End Type

Private Type TPalette
    ActBody   As Long
    FcstBody  As Long
    MixBody   As Long
    ActHead   As Long
    FcstHead  As Long
    MixHead   As Long
    HeadText  As Long
    Divider   As Long
End Type

Private mPal As TPalette

Private Sub InitPalette()
    mPal.ActBody = RGB(233, 241, 250)    ' closed months - pale blue
    mPal.FcstBody = RGB(255, 243, 214)   ' forecast months - pale amber
    mPal.MixBody = RGB(240, 240, 240)    ' part actual / part forecast - grey
    mPal.ActHead = RGB(0, 75, 141)       ' the existing dark blue header
    mPal.FcstHead = RGB(191, 143, 0)     ' dark amber header
    mPal.MixHead = RGB(128, 128, 128)    ' grey header
    mPal.HeadText = RGB(255, 255, 255)
    mPal.Divider = RGB(0, 0, 0)
End Sub


'============================== Entry points ==============================

Public Sub FormatGAAP_ActualVsForecast()
    RunFormat MAIN_SHEET
End Sub

Public Sub FormatAllTabs_ActualVsForecast()
    RunFormat TARGET_SHEETS
End Sub

Public Sub ResetGAAP_Formatting()
    RunReset MAIN_SHEET
End Sub

Public Sub ResetAllTabs_Formatting()
    RunReset TARGET_SHEETS
End Sub


'============================== Drivers ===================================

Private Sub RunFormat(ByVal sheetList As String)
    Dim names() As String, i As Long
    Dim ws As Worksheet, blk As TBlock
    Dim lastActual As Long, done As String, skipped As String
    Dim scrn As Boolean, alerts As Boolean

    InitPalette
    names = Split(sheetList, ",")

    ' Ask for the month once, off the first readable tab, then apply the same
    ' answer everywhere so the tabs cannot drift apart.
    Set ws = Nothing
    For i = LBound(names) To UBound(names)
        Set ws = SheetOrNothing(Trim$(names(i)))
        If Not ws Is Nothing Then
            blk = FindBlock(ws)
            If blk.Found Then Exit For
        End If
        Set ws = Nothing
    Next i

    If ws Is Nothing Then
        MsgBox "Could not find a " & FY_LABEL & " month block on: " & sheetList & vbCrLf & vbCrLf & _
               "Expected '" & FY_LABEL & " Actual' or '" & FY_LABEL & " Forecast' in row " & BANNER_ROW & _
               ", month names (Oct, Nov, ...) in row " & MONTH_ROW & ", and a 'Total' column.", _
               vbExclamation, "GAAP formatter"
        Exit Sub
    End If

    lastActual = AskLastActualMonth(ws, blk)
    If lastActual < 0 Then Exit Sub                  ' cancelled

    scrn = Application.ScreenUpdating
    alerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanUp

    For i = LBound(names) To UBound(names)
        Set ws = SheetOrNothing(Trim$(names(i)))
        If ws Is Nothing Then
            skipped = skipped & vbCrLf & "  - " & Trim$(names(i)) & " (tab not found)"
        Else
            blk = FindBlock(ws)
            If Not blk.Found Then
                skipped = skipped & vbCrLf & "  - " & ws.Name & " (no " & FY_LABEL & " block)"
            Else
                EnsureBackup ws, blk
                RestoreBaseline ws
                PaintBlock ws, blk, lastActual
                done = done & vbCrLf & "  - " & ws.Name
            End If
        End If
    Next i

CleanUp:
    Application.ScreenUpdating = scrn
    Application.DisplayAlerts = alerts
    If Err.Number <> 0 Then
        MsgBox "Formatting stopped: " & Err.Description, vbExclamation, "GAAP formatter"
        Exit Sub
    End If

    MsgBox FY_LABEL & " is now marked actual through " & FiscalMonthName(lastActual) & "." & _
           vbCrLf & vbCrLf & "Formatted:" & done & _
           IIf(Len(skipped) > 0, vbCrLf & vbCrLf & "Skipped:" & skipped, ""), _
           vbInformation, "GAAP formatter"
End Sub

Private Sub RunReset(ByVal sheetList As String)
    Dim names() As String, i As Long, ws As Worksheet, n As Long
    Dim scrn As Boolean, alerts As Boolean

    scrn = Application.ScreenUpdating
    alerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanUp

    names = Split(sheetList, ",")
    For i = LBound(names) To UBound(names)
        Set ws = SheetOrNothing(Trim$(names(i)))
        If Not ws Is Nothing Then
            If RestoreBaseline(ws) Then n = n + 1
        End If
    Next i

CleanUp:
    Application.ScreenUpdating = scrn
    Application.DisplayAlerts = alerts
    If Err.Number <> 0 Then
        MsgBox "Reset stopped: " & Err.Description, vbExclamation, "GAAP formatter"
    ElseIf n = 0 Then
        MsgBox "Nothing to reset - no saved baseline was found." & vbCrLf & _
               "(The baseline is saved the first time the formatter runs.)", _
               vbInformation, "GAAP formatter"
    Else
        MsgBox "Original formatting restored on " & n & " tab(s).", vbInformation, "GAAP formatter"
    End If
End Sub


'============================== Painting ==================================

Private Sub PaintBlock(ws As Worksheet, blk As TBlock, ByVal lastActual As Long)
    Dim c As Long, r As Long, status As Long
    Dim boundaryCol As Long, firstBodyRow As Long
    Dim bodyClr As Long, headClr As Long

    firstBodyRow = MONTH_ROW + 1
    boundaryCol = 0

    For c = blk.StartCol To blk.EndCol
        If blk.LastMonth(c) > 0 Then
            status = ColumnStatus(blk.FirstMonth(c), blk.LastMonth(c), lastActual)
            If status <> ST_ACTUAL And boundaryCol = 0 Then boundaryCol = c

            Select Case status
                Case ST_ACTUAL:   bodyClr = mPal.ActBody:  headClr = mPal.ActHead
                Case ST_FORECAST: bodyClr = mPal.FcstBody: headClr = mPal.FcstHead
                Case Else:        bodyClr = mPal.MixBody:  headClr = mPal.MixHead
            End Select

            With ws.Cells(MONTH_ROW, c)
                .Interior.Color = headClr
                .Font.Color = mPal.HeadText
                .Font.Bold = True
            End With

            For r = firstBodyRow To blk.LastRow
                If Not (KEEP_HIGHLIGHTS And IsHighlighted(ws.Cells(r, c))) Then
                    ws.Cells(r, c).Interior.Color = bodyClr
                End If
            Next r
        End If
    Next c

    If boundaryCol = 0 Then boundaryCol = blk.EndCol + 1      ' whole year closed

    WriteBanner ws, blk, lastActual, boundaryCol
    DrawDividers ws, blk, boundaryCol
    If SHOW_LEGEND Then WriteLegend ws, blk
End Sub

' Row 3: re-split, re-label and re-merge the two banners.
Private Sub WriteBanner(ws As Worksheet, blk As TBlock, ByVal lastActual As Long, ByVal boundaryCol As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(BANNER_ROW, blk.StartCol), ws.Cells(BANNER_ROW, blk.EndCol))
    rng.UnMerge
    rng.ClearContents

    If boundaryCol > blk.StartCol Then
        BannerPart ws, blk.StartCol, MinL(boundaryCol - 1, blk.EndCol), _
                   FY_LABEL & " Actual" & RangeSuffix(1, lastActual), mPal.ActHead
    End If
    If boundaryCol <= blk.EndCol Then
        BannerPart ws, boundaryCol, blk.EndCol, _
                   FY_LABEL & " Forecast" & RangeSuffix(lastActual + 1, 12), mPal.FcstHead
    End If
End Sub

Private Sub BannerPart(ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, _
                       ByVal caption As String, ByVal fillClr As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(BANNER_ROW, c1), ws.Cells(BANNER_ROW, c2))
    ws.Cells(BANNER_ROW, c1).Value = caption
    If c2 > c1 Then rng.Merge
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = fillClr
        .Font.Bold = True
        .Font.Color = mPal.HeadText
    End With
End Sub

Private Function RangeSuffix(ByVal m1 As Long, ByVal m2 As Long) As String
    If Not BANNER_SHOWS_RANGE Then Exit Function
    If m1 < 1 Or m2 > 12 Or m1 > m2 Then Exit Function
    If m1 = m2 Then
        RangeSuffix = " (" & FiscalMonthName(m1) & ")"
    Else
        RangeSuffix = " (" & FiscalMonthName(m1) & " - " & FiscalMonthName(m2) & ")"
    End If
End Function

' Clears last month's divider, then draws the quarter lines and the current
' actual / forecast divider.
Private Sub DrawDividers(ws As Worksheet, blk As TBlock, ByVal boundaryCol As Long)
    Dim c As Long

    ws.Range(ws.Cells(BANNER_ROW, blk.StartCol), ws.Cells(blk.LastRow, blk.EndCol)) _
      .Borders(xlInsideVertical).LineStyle = xlLineStyleNone

    If QUARTER_LINES Then
        For c = blk.StartCol To blk.EndCol - 1
            If blk.LastMonth(c) - blk.FirstMonth(c) = 2 Then          ' a quarter column
                With ws.Range(ws.Cells(BANNER_ROW, c), ws.Cells(MONTH_ROW, c)).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = mPal.Divider
                End With
            End If
        Next c
    End If

    If boundaryCol > blk.StartCol And boundaryCol <= blk.EndCol Then
        With ws.Range(ws.Cells(BANNER_ROW, boundaryCol), ws.Cells(blk.LastRow, boundaryCol)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = mPal.Divider
        End With
    End If
End Sub

Private Sub WriteLegend(ws As Worksheet, blk As TBlock)
    LegendCell ws, blk.StartCol, 1, mPal.ActBody
    LegendCell ws, blk.StartCol + 1, 2, mPal.FcstBody
    LegendCell ws, blk.StartCol + 2, 3, mPal.MixBody
End Sub

Private Function LegendCaption(ByVal i As Long) As String
    Select Case i
        Case 1: LegendCaption = "Actual"
        Case 2: LegendCaption = "Forecast"
        Case Else: LegendCaption = "Actual + fcst"
    End Select
End Function

Private Sub LegendCell(ws As Worksheet, ByVal c As Long, ByVal captionId As Long, ByVal fillClr As Long)
    Dim cell As Range
    If c > ws.Columns.Count Then Exit Sub
    Set cell = ws.Cells(LEGEND_ROW, c)
    If Len(Trim$(CStr(cell.Value))) > 0 And Not IsLegendCell(cell) Then Exit Sub   ' don't clobber your own notes
    With cell
        .Value = LegendCaption(captionId)
        .Interior.Color = fillClr
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Italic = False
        .Font.Size = 9
        .Font.Color = RGB(0, 0, 0)
    End With
End Sub

Private Function IsLegendCell(cell As Range) As Boolean
    Dim i As Long, s As String
    s = Trim$(CStr(cell.Value))
    For i = 1 To 3
        If StrComp(s, LegendCaption(i), vbTextCompare) = 0 Then
            IsLegendCell = True
            Exit Function
        End If
    Next i
End Function

Private Sub ClearLegend(ws As Worksheet, tgt As Range)
    Dim c As Long, cell As Range
    For c = tgt.Column To tgt.Column + tgt.Columns.Count - 1
        Set cell = ws.Cells(LEGEND_ROW, c)
        If IsLegendCell(cell) Then cell.ClearContents
    Next c
End Sub


'============================== Block discovery ===========================

' Locates the fiscal-year month block: first month column through "Total".
Private Function FindBlock(ws As Worksheet) As TBlock
    Dim blk As TBlock
    Dim lastCol As Long, c As Long, startCol As Long, endCol As Long
    Dim lbl As String, idx As Long

    lastCol = ws.Cells(MONTH_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then Exit Function

    ' Preferred: first column carrying an "FYxx ..." banner whose month row
    ' says "Oct".
    For c = 1 To lastCol
        If InStr(1, CStr(MergedValue(ws.Cells(BANNER_ROW, c))), FY_LABEL, vbTextCompare) = 1 Then
            If FiscalIndexOf(ws.Cells(MONTH_ROW, c).Value) = 1 Then
                startCol = c
                Exit For
            End If
        End If
    Next c

    ' Fallback: the right-most "Oct" in the month row (the prior year sits left).
    If startCol = 0 Then
        For c = lastCol To 1 Step -1
            If FiscalIndexOf(ws.Cells(MONTH_ROW, c).Value) = 1 Then
                startCol = c
                Exit For
            End If
        Next c
    End If
    If startCol = 0 Then Exit Function

    ReDim blk.FirstMonth(startCol To lastCol)
    ReDim blk.LastMonth(startCol To lastCol)

    For c = startCol To lastCol
        lbl = LCase$(Trim$(CStr(ws.Cells(MONTH_ROW, c).Value)))
        idx = FiscalIndexOf(lbl)
        If idx > 0 Then
            blk.FirstMonth(c) = idx
            blk.LastMonth(c) = idx
        ElseIf lbl Like "q#" Then
            idx = CLng(Mid$(lbl, 2, 1))
            blk.FirstMonth(c) = (idx - 1) * 3 + 1
            blk.LastMonth(c) = idx * 3
        ElseIf lbl = "total" Then
            blk.FirstMonth(c) = 1
            blk.LastMonth(c) = 12
            endCol = c
            Exit For
        End If
    Next c

    If endCol = 0 Then Exit Function

    blk.StartCol = startCol
    blk.EndCol = endCol
    blk.LastRow = FindLastBodyRow(ws, startCol, endCol)
    blk.Found = (blk.LastRow > MONTH_ROW)
    FindBlock = blk
End Function

' Last row of the statement itself - stops above the footnotes, which can
' hold stray working formulas.
Private Function FindLastBodyRow(ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long) As Long
    Dim r As Long, stopRow As Long, lastFilled As Long

    stopRow = MAX_SCAN_ROW
    For r = MONTH_ROW + 1 To MAX_SCAN_ROW
        If IsFootnoteLabel(CStr(ws.Cells(r, 1).Value)) Then
            stopRow = r - 1
            Exit For
        End If
    Next r

    For r = MONTH_ROW + 1 To stopRow
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))) > 0 Then
            lastFilled = r
        End If
    Next r
    FindLastBodyRow = lastFilled
End Function

' Footnote rows look like "1 Adjusted Operating EBIDA ..." or "10 Prior year ...".
Private Function IsFootnoteLabel(ByVal s As String) As Boolean
    Dim i As Long
    s = Trim$(s)
    If Len(s) < 15 Then Exit Function
    Do While i < Len(s)
        If Mid$(s, i + 1, 1) Like "#" Then i = i + 1 Else Exit Do
    Loop
    If i = 0 Then Exit Function
    IsFootnoteLabel = (Mid$(s, i + 1, 1) = " ")
End Function


'============================== Month handling ============================

' 1 = Oct ... 12 = Sep.  Returns 0 when the text is not a month.
Private Function FiscalIndexOf(ByVal v As Variant) As Long
    Dim s As String, i As Long, fy As Variant
    fy = Array("oct", "nov", "dec", "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep")
    s = LCase$(Trim$(CStr(v)))
    If Len(s) < 3 Then Exit Function
    s = Left$(s, 3)
    For i = 0 To 11
        If s = fy(i) Then
            FiscalIndexOf = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function FiscalMonthName(ByVal idx As Long) As String
    Dim fy As Variant
    fy = Array("Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep")
    If idx < 1 Then
        FiscalMonthName = "(none)"
    ElseIf idx > 12 Then
        FiscalMonthName = "Sep"
    Else
        FiscalMonthName = fy(idx - 1)
    End If
End Function

' Calendar month (1 = Jan) -> fiscal position (1 = Oct).
Private Function CalendarToFiscal(ByVal calMonth As Long) As Long
    CalendarToFiscal = ((calMonth - 10 + 12) Mod 12) + 1
End Function

Private Function ColumnStatus(ByVal m1 As Long, ByVal m2 As Long, ByVal lastActual As Long) As Long
    If m2 <= lastActual Then
        ColumnStatus = ST_ACTUAL
    ElseIf m1 > lastActual Then
        ColumnStatus = ST_FORECAST
    Else
        ColumnStatus = ST_MIXED
    End If
End Function

' Asks which month is closed.  Returns 0..12, or -1 if the user cancels.
Private Function AskLastActualMonth(ws As Worksheet, blk As TBlock) As Long
    Dim currentIdx As Long, defIdx As Long
    Dim ans As String, idx As Long, prompt As String

    currentIdx = DetectCurrentLastActual(ws, blk)
    defIdx = CalendarToFiscal(Month(DateSerial(Year(Date), Month(Date) - 1, 1)))

    prompt = "Which is the last CLOSED month for " & FY_LABEL & "?" & vbCrLf & vbCrLf & _
             "Type a month name (Jun, June, Jul ...)," & vbCrLf & _
             "a calendar month number (1 = Jan ... 12 = Dec)," & vbCrLf & _
             "or 0 for none - everything forecast." & vbCrLf & vbCrLf & _
             "Currently marked actual through: " & IIf(currentIdx = 0, "(none)", FiscalMonthName(currentIdx))

    ans = Trim$(InputBox(prompt, "GAAP - actual vs. forecast", FiscalMonthName(defIdx)))
    If Len(ans) = 0 Then
        AskLastActualMonth = -1
        Exit Function
    End If

    If IsNumeric(ans) Then
        idx = CLng(Val(ans))
        If idx = 0 Then
            AskLastActualMonth = 0
            Exit Function
        End If
        If idx < 1 Or idx > 12 Then GoTo BadInput
        AskLastActualMonth = CalendarToFiscal(idx)
        Exit Function
    End If

    If LCase$(ans) = "none" Then
        AskLastActualMonth = 0
        Exit Function
    End If

    idx = FiscalIndexOf(ans)
    If idx = 0 Then GoTo BadInput
    AskLastActualMonth = idx
    Exit Function

BadInput:
    MsgBox "'" & ans & "' is not a month. Nothing was changed.", vbExclamation, "GAAP formatter"
    AskLastActualMonth = -1
End Function

' Reads the existing banner so the prompt can show where things stand.
Private Function DetectCurrentLastActual(ws As Worksheet, blk As TBlock) As Long
    Dim c As Long, endCol As Long, best As Long
    For c = blk.StartCol To blk.EndCol
        If InStr(1, CStr(MergedValue(ws.Cells(BANNER_ROW, c))), FY_LABEL & " Actual", vbTextCompare) = 1 Then
            With ws.Cells(BANNER_ROW, c).MergeArea
                endCol = .Column + .Columns.Count - 1
            End With
            Exit For
        End If
    Next c
    If endCol = 0 Then Exit Function
    If endCol > blk.EndCol Then endCol = blk.EndCol
    For c = blk.StartCol To endCol
        If blk.LastMonth(c) > best Then best = blk.LastMonth(c)
    Next c
    DetectCurrentLastActual = best
End Function


'============================== Backup / restore ==========================

Private Function BackupName(ws As Worksheet) As String
    BackupName = Left$(BACKUP_PREFIX & ws.Name, 31)
End Function

' Snapshots shading and vertical borders once, so the original look is
' always recoverable.
Private Sub EnsureBackup(ws As Worksheet, blk As TBlock)
    Dim bk As Worksheet, src As Range, cell As Range
    Dim nR As Long, nC As Long, r As Long, c As Long
    Dim gFill As Variant
    Dim gLls As Variant, gLwt As Variant, gLclr As Variant
    Dim gRls As Variant, gRwt As Variant, gRclr As Variant
    Dim prevSh As Object, prevWb As Workbook

    If Not SheetOrNothing(BackupName(ws)) Is Nothing Then Exit Sub

    Set src = ws.Range(ws.Cells(LEGEND_ROW, blk.StartCol), ws.Cells(blk.LastRow, blk.EndCol))
    nR = src.Rows.Count
    nC = src.Columns.Count
    ReDim gFill(1 To nR, 1 To nC)
    ReDim gLls(1 To nR, 1 To nC): ReDim gLwt(1 To nR, 1 To nC): ReDim gLclr(1 To nR, 1 To nC)
    ReDim gRls(1 To nR, 1 To nC): ReDim gRwt(1 To nR, 1 To nC): ReDim gRclr(1 To nR, 1 To nC)

    For r = 1 To nR
        For c = 1 To nC
            Set cell = src.Cells(r, c)
            If cell.Interior.ColorIndex = xlNone Then
                gFill(r, c) = -1
            Else
                gFill(r, c) = cell.Interior.Color
            End If
            With cell.Borders(xlEdgeLeft)
                gLls(r, c) = .LineStyle
                If .LineStyle = xlLineStyleNone Then
                    gLwt(r, c) = xlThin: gLclr(r, c) = 0
                Else
                    gLwt(r, c) = .Weight: gLclr(r, c) = .Color
                End If
            End With
            With cell.Borders(xlEdgeRight)
                gRls(r, c) = .LineStyle
                If .LineStyle = xlLineStyleNone Then
                    gRwt(r, c) = xlThin: gRclr(r, c) = 0
                Else
                    gRwt(r, c) = .Weight: gRclr(r, c) = .Color
                End If
            End With
        Next c
    Next r

    Set prevWb = ActiveWorkbook
    On Error Resume Next
    Set prevSh = ActiveSheet
    On Error GoTo 0

    Set bk = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
    bk.Name = BackupName(ws)
    bk.Range("A1").Value = ws.Name
    bk.Range("B1").Value = src.Address(False, False)
    bk.Range("C1").Value = nR
    bk.Range("D1").Value = nC
    bk.Range("E1").Value = Now
    bk.Range("F1").Value = "Original shading and vertical borders saved by modGAAPActualForecast. Do not edit or delete."

    bk.Cells(GridRow(1, nR), 1).Resize(nR, nC).Value = gFill
    bk.Cells(GridRow(2, nR), 1).Resize(nR, nC).Value = gLls
    bk.Cells(GridRow(3, nR), 1).Resize(nR, nC).Value = gLwt
    bk.Cells(GridRow(4, nR), 1).Resize(nR, nC).Value = gLclr
    bk.Cells(GridRow(5, nR), 1).Resize(nR, nC).Value = gRls
    bk.Cells(GridRow(6, nR), 1).Resize(nR, nC).Value = gRwt
    bk.Cells(GridRow(7, nR), 1).Resize(nR, nC).Value = gRclr

    ' A sheet cannot be hidden while it is the active sheet.
    On Error Resume Next
    If Not prevWb Is Nothing Then prevWb.Activate
    If Not prevSh Is Nothing Then prevSh.Activate
    On Error GoTo 0
    bk.Visible = xlSheetVeryHidden
End Sub

' Puts shading, vertical borders and the legend row back to the baseline.
Private Function RestoreBaseline(ws As Worksheet) As Boolean
    Dim bk As Worksheet, tgt As Range, cell As Range
    Dim nR As Long, nC As Long, r As Long, c As Long
    Dim gFill As Variant
    Dim gLls As Variant, gLwt As Variant, gLclr As Variant
    Dim gRls As Variant, gRwt As Variant, gRclr As Variant

    Set bk = SheetOrNothing(BackupName(ws))
    If bk Is Nothing Then Exit Function

    nR = CLng(Val(bk.Range("C1").Value))
    nC = CLng(Val(bk.Range("D1").Value))
    If nR < 1 Or nC < 1 Then Exit Function

    On Error Resume Next
    Set tgt = ws.Range(CStr(bk.Range("B1").Value))
    On Error GoTo 0
    If tgt Is Nothing Then Exit Function
    If tgt.Rows.Count <> nR Or tgt.Columns.Count <> nC Then Exit Function

    gFill = bk.Cells(GridRow(1, nR), 1).Resize(nR, nC).Value
    gLls = bk.Cells(GridRow(2, nR), 1).Resize(nR, nC).Value
    gLwt = bk.Cells(GridRow(3, nR), 1).Resize(nR, nC).Value
    gLclr = bk.Cells(GridRow(4, nR), 1).Resize(nR, nC).Value
    gRls = bk.Cells(GridRow(5, nR), 1).Resize(nR, nC).Value
    gRwt = bk.Cells(GridRow(6, nR), 1).Resize(nR, nC).Value
    gRclr = bk.Cells(GridRow(7, nR), 1).Resize(nR, nC).Value

    ' Wipe first in bulk, then put back only the cells that carried something -
    ' far quicker than setting every cell one at a time.
    tgt.Interior.Pattern = xlNone
    tgt.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    tgt.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    tgt.Borders(xlEdgeRight).LineStyle = xlLineStyleNone

    For r = 1 To nR
        For c = 1 To nC
            If CLng(Val(gFill(r, c))) <> -1 Then
                tgt.Cells(r, c).Interior.Color = CLng(Val(gFill(r, c)))
            End If
            If CLng(Val(gLls(r, c))) <> xlLineStyleNone Then
                Set cell = tgt.Cells(r, c)
                With cell.Borders(xlEdgeLeft)
                    .LineStyle = CLng(Val(gLls(r, c)))
                    .Weight = CLng(Val(gLwt(r, c)))
                    .Color = CLng(Val(gLclr(r, c)))
                End With
            End If
            If CLng(Val(gRls(r, c))) <> xlLineStyleNone Then
                Set cell = tgt.Cells(r, c)
                With cell.Borders(xlEdgeRight)
                    .LineStyle = CLng(Val(gRls(r, c)))
                    .Weight = CLng(Val(gRwt(r, c)))
                    .Color = CLng(Val(gRclr(r, c)))
                End With
            End If
        Next c
    Next r

    ClearLegend ws, tgt
    RestoreBaseline = True
End Function

' Grids are stacked down the backup sheet with a blank row between them.
Private Function GridRow(ByVal gridIndex As Long, ByVal nR As Long) As Long
    GridRow = 3 + (gridIndex - 1) * (nR + 1)
End Function


'============================== Small helpers =============================

Private Function TargetBook() As Workbook
    If USE_ACTIVE_WORKBOOK And Not ActiveWorkbook Is Nothing Then
        Set TargetBook = ActiveWorkbook
    Else
        Set TargetBook = ThisWorkbook
    End If
End Function

Private Function SheetOrNothing(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = TargetBook().Worksheets(nm)
    On Error GoTo 0
End Function

Private Function MergedValue(cell As Range) As Variant
    MergedValue = cell.MergeArea.Cells(1, 1).Value
End Function

' True when the cell already carried a deliberate highlight.
Private Function IsHighlighted(cell As Range) As Boolean
    If cell.Interior.ColorIndex = xlNone Then Exit Function
    If cell.Interior.Color = RGB(255, 255, 255) Then Exit Function
    IsHighlighted = True
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function
