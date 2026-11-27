Attribute VB_Name = "ModuleControleFactures"
Option Explicit

Private Const CONTROL_SHEET As String = "Controle Factures"
Private Const TRACKING_FILE_PATTERN As String = "TRACKING EXPORT*.xlsx"
Private Const MAP_FIRST_ROW As Long = 5
Private Const GROUP_FIRST_ROW As Long = 5

Public Sub InstallerControleFactures()
    ActualiserGroupesFactures
End Sub

Public Sub ActualiserGroupesFactures()
    Dim facture As Workbook
    Dim ws As Worksheet
    Dim tracking As Workbook
    Dim wsTracking As Worksheet
    Dim trackingPath As String
    Dim wasAlreadyOpen As Boolean
    Dim oldGroups As Object
    Dim comboOrder As Collection
    Dim comboData As Object
    Dim comboCounts As Object
    Dim item As Variant
    Dim comboKey As String
    Dim program As String
    Dim sourceLot As String
    Dim invoiceGroup As String
    Dim mapRow As Long
    Dim btn As Button

    On Error GoTo CleanFail
    Set facture = ThisWorkbook
    trackingPath = FindWorkbookPath(facture.Path, TRACKING_FILE_PATTERN, facture.Name)
    If Len(trackingPath) = 0 Then
        Err.Raise vbObjectError + 2100, , _
            "No file matching '" & TRACKING_FILE_PATTERN & "' found in: " & facture.Path
    End If

    wasAlreadyOpen = IsWorkbookOpen(trackingPath)
    Set tracking = OpenReadOnlyWorkbook(trackingPath)
    Set wsTracking = FindTrackingSheet(tracking)
    If wsTracking Is Nothing Then
        Err.Raise vbObjectError + 2101, , "Sheet 'TRACKING EXPORT' not found."
    End If

    Set ws = GetOrCreateControlSheet(facture)
    Set oldGroups = ReadExistingGroups(ws)
    Set comboOrder = New Collection
    Set comboData = CreateObject("Scripting.Dictionary")
    comboData.CompareMode = vbTextCompare
    Set comboCounts = CreateObject("Scripting.Dictionary")
    comboCounts.CompareMode = vbTextCompare

    For Each item In TrackingCombinations(wsTracking)
        program = CStr(item(0))
        sourceLot = CStr(item(1))
        comboKey = MappingKey(program, sourceLot)
        If Not comboCounts.Exists(comboKey) Then
            comboOrder.Add comboKey
            comboData.Add comboKey, Array(program, sourceLot)
            comboCounts.Add comboKey, 0&
        End If
        comboCounts(comboKey) = CLng(comboCounts(comboKey)) + 1
    Next item

    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    ws.Cells.Clear

    ' Left table: LOT-to-Invoice-Group association UI
    ws.Range("A1").Value = "1. Program & LOT Grouping (Editable)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = _
        "Set the same 'Invoice Group' name for LOTs of the same Program to invoice them together, then click 'Refresh Groups'."
    ws.Range("A2:D2").Merge
    ws.Range("A2").WrapText = True
    ws.Range("A4:D4").Value = Array("Program", "LOT", "Rows", "Invoice Group")
    StyleHeader ws.Range("A4:D4")

    ' Right table: Final invoices ready to generate
    ws.Range("F1").Value = "2. Invoices to Generate"
    ws.Range("F1").Font.Bold = True
    ws.Range("F1").Font.Size = 12
    ws.Range("F4:K4").Value = _
        Array("Program", "Invoice Group", "Included LOTs", "Rows", "Action", "Generated File")
    StyleHeader ws.Range("F4:K4")

    Set btn = ws.Buttons.Add(ws.Range("F2").Left, ws.Range("F2").Top + 2, 125, 24)
    btn.Caption = "Refresh Groups"
    btn.OnAction = "ActualiserGroupesFactures"

    Set btn = ws.Buttons.Add(ws.Range("H2").Left, ws.Range("H2").Top + 2, 165, 24)
    btn.Caption = "Generate All Invoices"
    btn.OnAction = "GenererToutesLesFactures"

    mapRow = MAP_FIRST_ROW
    For Each item In comboOrder
        comboKey = CStr(item)
        item = comboData(comboKey)
        program = CStr(item(0))
        sourceLot = CStr(item(1))
        If oldGroups.Exists(comboKey) Then
            invoiceGroup = CStr(oldGroups(comboKey))
        Else
            invoiceGroup = DefaultInvoiceGroup(program, sourceLot)
        End If
        ws.Cells(mapRow, 1).Value = program
        ws.Cells(mapRow, 2).Value = sourceLot
        ws.Cells(mapRow, 3).Value = comboCounts(comboKey)
        ws.Cells(mapRow, 4).Value = invoiceGroup
        ws.Cells(mapRow, 4).Interior.Color = RGB(255, 251, 230)
        ws.Cells(mapRow, 4).Borders.LineStyle = xlContinuous
        ws.Cells(mapRow, 4).Borders.Color = RGB(190, 190, 190)
        mapRow = mapRow + 1
    Next item

    BuildInvoiceGroups ws

    ws.Columns("A:K").AutoFit
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 22
    ws.Columns("H").ColumnWidth = 28
    ws.Columns("J").ColumnWidth = 15
    ws.Columns("K").ColumnWidth = 26
    ws.Rows(2).RowHeight = 32
    ws.Activate

    If Not wasAlreadyOpen And Not tracking Is Nothing Then
        tracking.Close SaveChanges:=False
    End If

    MsgBox "Control sheet refreshed.", vbInformation, "Invoice Control"
    Exit Sub

CleanFail:
    If Not wasAlreadyOpen And Not tracking Is Nothing Then
        On Error Resume Next
        tracking.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "Invoice Control"
End Sub

Public Sub GenererFactureDepuisLigne()
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim program As String
    Dim invoiceGroup As String
    Dim targetLots As Collection

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    rowNumber = ws.Shapes(Application.Caller).TopLeftCell.Row
    program = Trim$(CStr(ws.Cells(rowNumber, 6).Value))
    invoiceGroup = Trim$(CStr(ws.Cells(rowNumber, 7).Value))
    Set targetLots = GetLotsForGroup(ws, program, invoiceGroup)

    GenererUneFacture program, invoiceGroup, targetLots, ThisWorkbook
    ws.Cells(rowNumber, 11).Value = "Generated " & Format(Now, "dd/mm/yyyy hh:nn")
    Exit Sub
Fail:
    MsgBox "Invoice generation failed (" & Err.Number & "): " & Err.Description, _
           vbCritical, "Invoice Control"
End Sub

Public Sub GenererToutesLesFactures()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim doneCount As Long
    Dim program As String
    Dim invoiceGroup As String
    Dim targetLots As Collection

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    For rowNumber = GROUP_FIRST_ROW To lastRow
        program = Trim$(CStr(ws.Cells(rowNumber, 6).Value))
        invoiceGroup = Trim$(CStr(ws.Cells(rowNumber, 7).Value))
        If Len(program) > 0 And Len(invoiceGroup) > 0 Then
            Set targetLots = GetLotsForGroup(ws, program, invoiceGroup)
            GenererUneFacture program, invoiceGroup, targetLots, ThisWorkbook
            ws.Cells(rowNumber, 11).Value = "Generated " & Format(Now, "dd/mm/yyyy hh:nn")
            doneCount = doneCount + 1
        End If
    Next rowNumber
    MsgBox doneCount & " invoice(s) generated.", vbInformation, "Invoice Control"
    Exit Sub
Fail:
    MsgBox "Stopped at row " & rowNumber & " (" & Err.Number & "): " & Err.Description, _
           vbCritical, "Invoice Control"
End Sub

Private Sub BuildInvoiceGroups(ByVal ws As Worksheet)
    Dim groups As Collection
    Dim groupSeen As Object
    Dim rowNumber As Long
    Dim outputRow As Long
    Dim program As String
    Dim invoiceGroup As String
    Dim groupKey As String
    Dim btn As Button

    Set groups = New Collection
    Set groupSeen = CreateObject("Scripting.Dictionary")
    groupSeen.CompareMode = vbTextCompare

    For rowNumber = MAP_FIRST_ROW To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        program = Trim$(CStr(ws.Cells(rowNumber, 1).Value))
        invoiceGroup = Trim$(CStr(ws.Cells(rowNumber, 4).Value))
        If Len(program) > 0 And Len(invoiceGroup) > 0 Then
            groupKey = MappingKey(program, invoiceGroup)
            If Not groupSeen.Exists(groupKey) Then
                groupSeen.Add groupKey, True
                groups.Add Array(program, invoiceGroup)
            End If
        End If
    Next rowNumber

    outputRow = GROUP_FIRST_ROW
    For Each groupKey In groups
        program = CStr(groupKey(0))
        invoiceGroup = CStr(groupKey(1))
        ws.Cells(outputRow, 6).Value = program
        ws.Cells(outputRow, 7).Value = invoiceGroup
        ws.Cells(outputRow, 8).Value = IncludedLots(ws, program, invoiceGroup)
        ws.Cells(outputRow, 9).Value = CountGroupRows(ws, program, invoiceGroup)
        Set btn = ws.Buttons.Add(ws.Cells(outputRow, 10).Left + 2, _
                                 ws.Cells(outputRow, 10).Top + 1, 90, 18)
        btn.Caption = "Generate"
        btn.OnAction = "GenererFactureDepuisLigne"
        outputRow = outputRow + 1
    Next groupKey
End Sub

Private Function GetLotsForGroup(ByVal ws As Worksheet, ByVal program As String, _
                                 ByVal invoiceGroup As String) As Collection
    Dim lots As New Collection
    Dim seen As Object
    Dim rowNumber As Long
    Dim lot As String
    Dim key As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For rowNumber = MAP_FIRST_ROW To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(rowNumber, 1).Value)), program, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(ws.Cells(rowNumber, 4).Value)), invoiceGroup, vbTextCompare) = 0 Then
            lot = Trim$(CStr(ws.Cells(rowNumber, 2).Value))
            key = UCase$(lot)
            If Len(lot) > 0 And Not seen.Exists(key) Then
                seen.Add key, lot
                lots.Add lot
            End If
        End If
    Next rowNumber
    Set GetLotsForGroup = lots
End Function

Private Function IncludedLots(ByVal ws As Worksheet, ByVal program As String, _
                              ByVal invoiceGroup As String) As String
    Dim lots As Collection
    Dim item As Variant
    Dim result As String

    Set lots = GetLotsForGroup(ws, program, invoiceGroup)
    For Each item In lots
        If Len(result) > 0 Then result = result & ", "
        result = result & CStr(item)
    Next item
    IncludedLots = result
End Function

Private Function CountGroupRows(ByVal wsControl As Worksheet, _
                                ByVal program As String, _
                                ByVal invoiceGroup As String) As Long
    Dim rowNumber As Long

    For rowNumber = MAP_FIRST_ROW To wsControl.Cells(wsControl.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(wsControl.Cells(rowNumber, 1).Value)), program, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(wsControl.Cells(rowNumber, 4).Value)), invoiceGroup, vbTextCompare) = 0 Then
            CountGroupRows = CountGroupRows + CLng(wsControl.Cells(rowNumber, 3).Value)
        End If
    Next rowNumber
End Function

Private Function TrackingCombinations(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    Dim seen As Object
    Dim rowNumber As Long
    Dim program As String
    Dim lot As String
    Dim key As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For rowNumber = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        program = Trim$(CStr(ws.Cells(rowNumber, 1).Value2))
        lot = NormalizeInvoiceLot(program, CStr(ws.Cells(rowNumber, 2).Value2))
        If Len(program) > 0 And Len(lot) > 0 Then
            key = MappingKey(program, lot)
            If Not seen.Exists(key) Then
                seen.Add key, True
                result.Add Array(program, lot)
            End If
        End If
    Next rowNumber
    Set TrackingCombinations = result
End Function

Private Function ReadExistingGroups(ByVal ws As Worksheet) As Object
    Dim groups As Object
    Dim rowNumber As Long
    Dim program As String
    Dim lot As String
    Dim invoiceGroup As String
    Dim key As String

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    For rowNumber = MAP_FIRST_ROW To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        program = Trim$(CStr(ws.Cells(rowNumber, 1).Value))
        lot = NormalizeInvoiceLot(program, CStr(ws.Cells(rowNumber, 2).Value))
        invoiceGroup = Trim$(CStr(ws.Cells(rowNumber, 4).Value))
        If Len(program) > 0 And Len(lot) > 0 And Len(invoiceGroup) > 0 Then
            key = MappingKey(program, lot)
            groups(key) = invoiceGroup
        End If
    Next rowNumber
    Set ReadExistingGroups = groups
End Function

Private Function DefaultInvoiceGroup(ByVal program As String, _
                                     ByVal sourceLot As String) As String
    If StrComp(program, "WINGS", vbTextCompare) = 0 And _
       (StrComp(sourceLot, "900", vbTextCompare) = 0 Or _
        StrComp(sourceLot, "1000", vbTextCompare) = 0) Then
        DefaultInvoiceGroup = "900+1000"
    ElseIf StrComp(program, "ATR", vbTextCompare) = 0 And _
           (StrComp(sourceLot, "LOT2", vbTextCompare) = 0 Or _
            StrComp(sourceLot, "LOT3", vbTextCompare) = 0) Then
        DefaultInvoiceGroup = "LOT2+LOT3"
    ElseIf StrComp(program, "AH", vbTextCompare) = 0 And _
           (StrComp(sourceLot, "H160", vbTextCompare) = 0 Or _
            StrComp(sourceLot, "H175", vbTextCompare) = 0) Then
        DefaultInvoiceGroup = "H160+H175"
    Else
        DefaultInvoiceGroup = sourceLot
    End If
End Function

Private Function NormalizeInvoiceLot(ByVal program As String, _
                                     ByVal lot As String) As String
    Dim slashPosition As Long

    lot = Trim$(lot)
    If StrComp(Trim$(program), "WINGS", vbTextCompare) = 0 Then
        slashPosition = InStrRev(lot, "/")
        If slashPosition > 0 Then
            lot = Trim$(Mid$(lot, slashPosition + 1))
        End If
    End If
    NormalizeInvoiceLot = lot
End Function

Private Function MappingKey(ByVal program As String, ByVal lot As String) As String
    MappingKey = UCase$(Trim$(program) & ChrW(1) & Trim$(lot))
End Function

Private Sub StyleHeader(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Interior.Color = RGB(31, 78, 121)
    rng.Font.Color = RGB(255, 255, 255)
End Sub

Private Function GetOrCreateControlSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, CONTROL_SHEET, vbTextCompare) = 0 Then
            Set GetOrCreateControlSheet = ws
            Exit Function
        End If
    Next ws
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = CONTROL_SHEET
    Set GetOrCreateControlSheet = ws
End Function

Private Function FindTrackingSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If UCase$(Left$(Trim$(ws.Name), 15)) = "TRACKING EXPORT" Then
            Set FindTrackingSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function IsWorkbookOpen(ByVal fullPath As String) As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function OpenReadOnlyWorkbook(ByVal path As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            Set OpenReadOnlyWorkbook = wb
            Exit Function
        End If
    Next wb
    Set OpenReadOnlyWorkbook = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=False)
End Function

Private Function FindWorkbookPath(ByVal folderPath As String, _
                                  ByVal filePattern As String, _
                                  ByVal excludeName As String) As String
    Dim candidate As String

    If Len(Trim$(folderPath)) = 0 Then Exit Function
    candidate = Dir(folderPath & Application.PathSeparator & filePattern)
    Do While Len(candidate) > 0
        If Left$(candidate, 2) <> "~$" And _
           StrComp(candidate, excludeName, vbTextCompare) <> 0 Then
            FindWorkbookPath = folderPath & Application.PathSeparator & candidate
            Exit Function
        End If
        candidate = Dir
    Loop
End Function