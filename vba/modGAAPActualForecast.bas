Attribute VB_Name = "modGAAPActualForecast"
Option Explicit

'==========================================================================
' GAAP tab - move the FY "Actual / Forecast" split
'--------------------------------------------------------------------------
' Every month-end one more column moves from forecast to actual, and the
' "FY26 Actual" / "FY26 Forecast" banner in row 3 has to be re-split by hand.
' This module does exactly that and nothing else:
'
'   * un-merges the banner, re-merges it at the correct column, and puts the
'     same two labels back - the existing wording, unchanged
'   * copies the banner formatting from the banner cells already on the
'     sheet, so the result is identical to what is there today
'   * moves the white rule between the two banners, down through the month
'     header row as well, and keeps every other rule exactly as it is
'
' It does not add colour, month names, legends or shading, and it never
' touches formulas, values, fonts, number formats, row heights or column
' widths.  Everything outside row 3 is left alone.
'
' Macros
'   UpdateGAAP_Split            GAAP tab, asks which month closed
'   UpdateGAAP_Split_Auto       GAAP tab, no prompt - uses last month
'   UpdateAllTabs_Split         GAAP + Combined, asks once
'   UpdateAllTabs_Split_Auto    GAAP + Combined, no prompt
'
' There is no undo for macros in Excel, but there is nothing to undo here:
' run it again with a different month and the banner moves again.
'==========================================================================


'---------------------- Settings you may want to change -------------------

' Tab used by the plain GAAP macros.
Private Const MAIN_SHEET      As String = "GAAP"

' Tabs used by the "AllTabs" macros (comma separated).
Private Const TARGET_SHEETS   As String = "GAAP,Combined"

' Fiscal year being tracked.  Used to find the block, and as the label text
' only if the sheet has no banner to copy - change to "FY27" next October.
Private Const FY_LABEL        As String = "FY26"

' Sheet layout (same on GAAP and Combined).
Private Const BANNER_ROW      As Long = 3    ' "FY26 Actual" / "FY26 Forecast"
Private Const MONTH_ROW       As Long = 4    ' Oct, Nov, Dec, Q1, ... , Total

' Leave False when this module lives in the forecast workbook itself.  Set to
' True if you keep it in PERSONAL.XLSB and run it against whatever workbook
' is in front of you.
Private Const USE_ACTIVE_WORKBOOK As Boolean = False


'---------------------- Internals -----------------------------------------

Private Const ST_ACTUAL   As Long = 1
Private Const ST_FORECAST As Long = 2
Private Const ST_MIXED    As Long = 3

Private Type TBlock
    Found        As Boolean
    StartCol     As Long      ' first month column of the FY block (Oct)
    EndCol       As Long      ' the "Total" column of the FY block
    FirstMonth() As Long      ' per column: first fiscal month covered (0 = spacer)
    LastMonth()  As Long      ' per column: last fiscal month covered  (0 = spacer)
End Type

Private Type TEdge
    Has       As Boolean
    LineStyle As Long
    Weight    As Long
    ThemeOK   As Boolean
    ThemeClr  As Long
    Tint      As Double
    Colr      As Long
End Type

Private Type TBanner
    ActCell  As String        ' address of the "Actual" banner cell
    ActText  As String
    FcstCell As String        ' address of the "Forecast" banner cell
    FcstText As String
    FcstCol  As Long          ' column the split sits on right now
    Split    As TEdge         ' the white rule drawn between the two banners
End Type

' One cell edge, remembered well enough to put back exactly - including a
' theme colour (these rules are white, theme 0) rather than a flat RGB.


'============================== Entry points ==============================

Public Sub UpdateGAAP_Split()
    RunSplit MAIN_SHEET, True
End Sub

Public Sub UpdateGAAP_Split_Auto()
    RunSplit MAIN_SHEET, False
End Sub

Public Sub UpdateAllTabs_Split()
    RunSplit TARGET_SHEETS, True
End Sub

Public Sub UpdateAllTabs_Split_Auto()
    RunSplit TARGET_SHEETS, False
End Sub


'============================== Driver ====================================

' askUser = False runs silently against last calendar month, for a button.
Private Sub RunSplit(ByVal sheetList As String, ByVal askUser As Boolean)
    Dim names() As String, i As Long
    Dim ws As Worksheet, blk As TBlock
    Dim lastActual As Long, done As String, skipped As String
    Dim scrn As Boolean, alerts As Boolean

    names = Split(sheetList, ",")

    ' Find a readable tab first, so the prompt can show where things stand.
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
        MsgBox "Could not find the " & FY_LABEL & " block on: " & sheetList & vbCrLf & vbCrLf & _
               "Expected the banner in row " & BANNER_ROW & ", month names (Oct, Nov, ...) in row " & _
               MONTH_ROW & ", and a 'Total' column.", vbExclamation, "Actual / forecast split"
        Exit Sub
    End If

    If askUser Then
        lastActual = AskLastActualMonth(ws, blk)
        If lastActual < 0 Then Exit Sub                  ' cancelled
    Else
        lastActual = CalendarToFiscal(Month(DateSerial(Year(Date), Month(Date) - 1, 1)))
    End If

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
                MoveSplit ws, blk, lastActual
                done = done & vbCrLf & "  - " & ws.Name
            End If
        End If
    Next i

CleanUp:
    Application.CutCopyMode = False
    Application.ScreenUpdating = scrn
    Application.DisplayAlerts = alerts

    If Err.Number <> 0 Then
        MsgBox "Stopped: " & Err.Description, vbExclamation, "Actual / forecast split"
    ElseIf askUser Then
        MsgBox "Actual through " & FiscalMonthName(lastActual) & "." & vbCrLf & "Updated:" & done & _
               IIf(Len(skipped) > 0, vbCrLf & vbCrLf & "Skipped:" & skipped, ""), _
               vbInformation, "Actual / forecast split"
    ElseIf Len(skipped) > 0 Then
        MsgBox "Actual through " & FiscalMonthName(lastActual) & ", but skipped:" & skipped, _
               vbExclamation, "Actual / forecast split"
    End If
End Sub


'============================== The move ==================================

Private Sub MoveSplit(ws As Worksheet, blk As TBlock, ByVal lastActual As Long)
    Dim bnr As TBanner
    Dim rng As Range
    Dim c As Long, boundaryCol As Long, actTo As Long
    Dim stashAct As String, stashFcst As String
    Dim blockLeft As TEdge, blockRight As TEdge

    bnr = ReadBanner(ws, blk)

    ' The block's own outer rules, which the format copy below would otherwise
    ' overwrite with whatever the banner cell happened to carry.
    blockLeft = CaptureEdge(ws.Cells(BANNER_ROW, blk.StartCol), xlEdgeLeft)
    blockRight = CaptureEdge(ws.Cells(BANNER_ROW, blk.EndCol), xlEdgeRight)

    ' First column that is not fully closed - the forecast banner starts here.
    For c = blk.StartCol To blk.EndCol
        If blk.LastMonth(c) > 0 Then
            If ColumnStatus(blk.FirstMonth(c), blk.LastMonth(c), lastActual) <> ST_ACTUAL Then
                boundaryCol = c
                Exit For
            End If
        End If
    Next c
    If boundaryCol = 0 Then boundaryCol = blk.EndCol + 1        ' whole year closed

    ' Park a copy of each banner's formatting off to the right first: laying
    ' down one banner can otherwise land on the cell the other one copies from.
    stashAct = StashFormat(ws, bnr.ActCell, 2)
    stashFcst = StashFormat(ws, bnr.FcstCell, 1)

    Set rng = ws.Range(ws.Cells(BANNER_ROW, blk.StartCol), ws.Cells(BANNER_ROW, blk.EndCol))
    rng.UnMerge
    rng.ClearContents

    actTo = MinL(boundaryCol - 1, blk.EndCol)
    If boundaryCol > blk.StartCol Then WriteBanner ws, blk.StartCol, actTo, bnr.ActText, stashAct
    If boundaryCol <= blk.EndCol Then WriteBanner ws, boundaryCol, blk.EndCol, bnr.FcstText, stashFcst

    DropStash ws, stashAct
    DropStash ws, stashFcst

    ' Borders while the cells are still separate, merging last - Excel is
    ' happier that way, and a merged block shows its outer cells' edges.
    RedrawRules ws, blk, bnr, boundaryCol, blockLeft, blockRight
    If boundaryCol > blk.StartCol Then MergeBand ws, blk.StartCol, actTo
    If boundaryCol <= blk.EndCol Then MergeBand ws, boundaryCol, blk.EndCol
End Sub

Private Sub MergeBand(ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long)
    If c2 > c1 Then ws.Range(ws.Cells(BANNER_ROW, c1), ws.Cells(BANNER_ROW, c2)).Merge
End Sub

' Copies a banner cell's formatting to a scratch cell past the right-hand end
' of the sheet, and returns its address ("" when there is nothing to copy).
Private Function StashFormat(ws As Worksheet, ByVal srcAddr As String, ByVal slot As Long) As String
    Dim tmp As Range
    If Len(srcAddr) = 0 Then Exit Function
    Set tmp = ws.Cells(BANNER_ROW, ws.Columns.Count - slot)
    tmp.Clear
    ws.Range(srcAddr).Copy
    tmp.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    StashFormat = tmp.Address(False, False)
End Function

Private Sub DropStash(ws As Worksheet, ByVal addr As String)
    If Len(addr) > 0 Then ws.Range(addr).Clear
End Sub

' Lays one banner over c1..c2, wearing the formatting the sheet's own banner
' cell had.  Merging happens later, in MergeBand.
Private Sub WriteBanner(ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, _
                        ByVal caption As String, ByVal srcAddr As String)
    If Len(srcAddr) > 0 Then
        ws.Range(srcAddr).Copy
        ws.Range(ws.Cells(BANNER_ROW, c1), ws.Cells(BANNER_ROW, c2)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    ws.Cells(BANNER_ROW, c1).Value = caption
End Sub

' Rebuilds the vertical rules in the banner row: the block's own outer rules,
' and the rule between the two banners, drawn on both sides of the boundary
' the way the sheet already does it.  Everything else in the row is interior
' to a merged banner, so it is cleared.
Private Sub RedrawRules(ws As Worksheet, blk As TBlock, bnr As TBanner, ByVal boundaryCol As Long, _
                        blockLeft As TEdge, blockRight As TEdge)
    Dim onSheet As Boolean

    ' --- banner row: the two banners, and the rule between them -----------
    ws.Range(ws.Cells(BANNER_ROW, blk.StartCol), ws.Cells(BANNER_ROW, blk.EndCol)) _
      .Borders(xlInsideVertical).LineStyle = xlLineStyleNone

    ApplyEdge ws.Cells(BANNER_ROW, blk.StartCol), xlEdgeLeft, blockLeft
    ApplyEdge ws.Cells(BANNER_ROW, blk.EndCol), xlEdgeRight, blockRight

    onSheet = (boundaryCol > blk.StartCol And boundaryCol <= blk.EndCol)
    If onSheet Then DrawRule ws, BANNER_ROW, boundaryCol, bnr.Split

    ' --- month row: the split rule carries on down through it -------------
    ' The sheet already rules every quarter and total column, so those lines
    ' stay put; the split only adds one where it lands mid-quarter, and only
    ' a line a previous run added mid-quarter is taken away again.
    If bnr.FcstCol > 0 And bnr.FcstCol <> boundaryCol Then
        If Not NativeRule(blk, bnr.FcstCol) Then ClearRule ws, MONTH_ROW, bnr.FcstCol
    End If
    If onSheet Then
        If Not NativeRule(blk, boundaryCol) Then DrawRule ws, MONTH_ROW, boundaryCol, bnr.Split
    End If
End Sub

' A rule sits on both cells that meet at it, the way the sheet draws them.
Private Sub DrawRule(ws As Worksheet, ByVal r As Long, ByVal c As Long, e As TEdge)
    If Not e.Has Then Exit Sub
    ApplyEdge ws.Cells(r, c), xlEdgeLeft, e
    ApplyEdge ws.Cells(r, c - 1), xlEdgeRight, e
End Sub

Private Sub ClearRule(ws As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim none As TEdge
    ApplyEdge ws.Cells(r, c), xlEdgeLeft, none
    ApplyEdge ws.Cells(r, c - 1), xlEdgeRight, none
End Sub

' True where the sheet itself rules the month row: either side of a quarter
' column, and either side of Total.
Private Function NativeRule(blk As TBlock, ByVal c As Long) As Boolean
    If c <= blk.StartCol Or c > blk.EndCol Then Exit Function
    NativeRule = IsGroupCol(blk, c) Or IsGroupCol(blk, c - 1)
End Function

Private Function IsGroupCol(blk As TBlock, ByVal c As Long) As Boolean
    If c < blk.StartCol Or c > blk.EndCol Then Exit Function
    If blk.LastMonth(c) = 0 Then Exit Function
    IsGroupCol = (blk.LastMonth(c) - blk.FirstMonth(c) >= 2)     ' quarter or total
End Function

Private Function CaptureEdge(cell As Range, ByVal edge As Long) As TEdge
    Dim e As TEdge
    With cell.Borders(edge)
        If .LineStyle = xlLineStyleNone Then
            CaptureEdge = e
            Exit Function
        End If
        e.Has = True
        e.LineStyle = .LineStyle
        e.Weight = .Weight
        e.Colr = .Color
        ' ThemeColor errors on a border that is not theme-coloured.
        On Error Resume Next
        e.ThemeClr = .ThemeColor
        e.Tint = .TintAndShade
        e.ThemeOK = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End With
    CaptureEdge = e
End Function

Private Sub ApplyEdge(cell As Range, ByVal edge As Long, e As TEdge)
    With cell.Borders(edge)
        If Not e.Has Then
            .LineStyle = xlLineStyleNone
            Exit Sub
        End If
        .LineStyle = e.LineStyle
        .Weight = e.Weight
        If e.ThemeOK Then
            .ThemeColor = e.ThemeClr
            .TintAndShade = e.Tint
        Else
            .Color = e.Colr
        End If
    End With
End Sub

' Reads the banners that are on the sheet now: their wording, which cell
' carries their formatting, and the rule drawn between them.
Private Function ReadBanner(ws As Worksheet, blk As TBlock) As TBanner
    Dim bnr As TBanner
    Dim c As Long, txt As String, area As Range

    For c = blk.StartCol To blk.EndCol
        Set area = ws.Cells(BANNER_ROW, c).MergeArea
        txt = Trim$(CStr(area.Cells(1, 1).Value))
        If Len(txt) > 0 Then
            If InStr(1, txt, "Forecast", vbTextCompare) > 0 Then
                If Len(bnr.FcstCell) = 0 Then
                    bnr.FcstCell = area.Cells(1, 1).Address(False, False)
                    bnr.FcstText = txt
                    bnr.FcstCol = area.Cells(1, 1).Column
                End If
            ElseIf InStr(1, txt, "Actual", vbTextCompare) > 0 Then
                If Len(bnr.ActCell) = 0 Then
                    bnr.ActCell = area.Cells(1, 1).Address(False, False)
                    bnr.ActText = txt
                End If
            End If
        End If
    Next c

    If Len(bnr.ActText) = 0 Then bnr.ActText = FY_LABEL & " Actual"
    If Len(bnr.FcstText) = 0 Then bnr.FcstText = FY_LABEL & " Forecast"

    ' A banner with no twin still lends its formatting to the other one.
    If Len(bnr.ActCell) = 0 Then bnr.ActCell = bnr.FcstCell
    If Len(bnr.FcstCell) = 0 Then bnr.FcstCell = bnr.ActCell

    ' The rule between the banners, taken from the forecast banner's left edge
    ' - white on both tabs today, thick on GAAP and medium on Combined.
    If Len(bnr.FcstCell) > 0 Then bnr.Split = CaptureEdge(ws.Range(bnr.FcstCell), xlEdgeLeft)
    If Not bnr.Split.Has And Len(bnr.ActCell) > 0 Then
        bnr.Split = CaptureEdge(ws.Range(bnr.ActCell), xlEdgeRight)
    End If

    ReadBanner = bnr
End Function


'============================== Block discovery ===========================

' Locates the fiscal-year month block: first month column through "Total".
Private Function FindBlock(ws As Worksheet) As TBlock
    Dim blk As TBlock
    Dim lastCol As Long, c As Long, startCol As Long, endCol As Long
    Dim lbl As String, idx As Long

    lastCol = ws.Cells(MONTH_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then Exit Function

    ' Preferred: first column carrying an "FYxx ..." banner above an "Oct".
    For c = 1 To lastCol
        If InStr(1, CStr(ws.Cells(BANNER_ROW, c).MergeArea.Cells(1, 1).Value), FY_LABEL, vbTextCompare) = 1 Then
            If FiscalIndexOf(ws.Cells(MONTH_ROW, c).Value) = 1 Then
                startCol = c
                Exit For
            End If
        End If
    Next c

    ' Fallback: the right-most "Oct" in the month row (prior year sits left).
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
    blk.Found = True
    FindBlock = blk
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
             "A month name (Jun, July ...), a calendar month number" & vbCrLf & _
             "(1 = Jan ... 12 = Dec), or 0 for none." & vbCrLf & vbCrLf & _
             "Marked actual through: " & IIf(currentIdx = 0, "(none)", FiscalMonthName(currentIdx))

    ans = Trim$(InputBox(prompt, "Actual / forecast split", FiscalMonthName(defIdx)))
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
    MsgBox "'" & ans & "' is not a month. Nothing was changed.", vbExclamation, "Actual / forecast split"
    AskLastActualMonth = -1
End Function

' Reads the current banner so the prompt can show where things stand.
Private Function DetectCurrentLastActual(ws As Worksheet, blk As TBlock) As Long
    Dim c As Long, endCol As Long, best As Long, txt As String

    For c = blk.StartCol To blk.EndCol
        txt = CStr(ws.Cells(BANNER_ROW, c).MergeArea.Cells(1, 1).Value)
        If InStr(1, txt, "Actual", vbTextCompare) > 0 And InStr(1, txt, "Forecast", vbTextCompare) = 0 Then
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

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function
