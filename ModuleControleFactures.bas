Attribute VB_Name = "ModuleControleFactures"
Option Explicit

Private Const CONTROL_SHEET As String = "Controle Factures"
Private Const TRACKING_FILE_PATTERN As String = "TRACKING EXPORT*.xlsx"

Public Sub InstallerControleFactures()
    Dim facture As Workbook
    Dim ws As Worksheet
    Dim tracking As Workbook
    Dim wsTracking As Worksheet
    Dim trackingPath As String
    Dim combinations As Collection
    Dim item As Variant
    Dim rowNumber As Long
    Dim btn As Button

    Set facture = ThisWorkbook
    trackingPath = FindWorkbookPath(facture.path, TRACKING_FILE_PATTERN, facture.Name)
    If Len(trackingPath) = 0 Then
        Err.Raise vbObjectError + 2100, , _
            "No file matching '" & TRACKING_FILE_PATTERN & "' found in: " & facture.path
    End If

    Set tracking = OpenReadOnlyWorkbook(trackingPath)
    Set wsTracking = FindTrackingSheet(tracking)
    If wsTracking Is Nothing Then
        Err.Raise vbObjectError + 2101, , "Sheet 'TRACKING EXPORT' not found."
    End If

    Set combinations = FindTrackingCombinations(wsTracking)
    Set ws = GetOrCreateControlSheet(facture)

    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ws.Cells.Clear

    ws.Range("A1:E1").value = Array("Program", "LOT", "Rows", "Action", "Generated file")
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:E1").Font.Color = RGB(255, 255, 255)
    ws.Range("G1").value = "Commands"
    ws.Range("G1").Font.Bold = True

    Set btn = ws.Buttons.Add(ws.Range("G2").Left, ws.Range("G2").Top, 160, 28)
    btn.Caption = "Generate all invoices"
    btn.OnAction = "GenererToutesLesFactures"

    rowNumber = 2
    For Each item In combinations
        ws.Cells(rowNumber, 1).value = item(0)
        ws.Cells(rowNumber, 2).value = item(1)
        ws.Cells(rowNumber, 3).value = CountMatchingRows(wsTracking, CStr(item(0)), CStr(item(1)))
        Set btn = ws.Buttons.Add(ws.Cells(rowNumber, 4).Left, _
                                 ws.Cells(rowNumber, 4).Top, 100, 20)
        btn.Caption = "Generate"
        btn.OnAction = "GenererFactureDepuisLigne"
        rowNumber = rowNumber + 1
    Next item

    ws.Columns("A:C").AutoFit
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 24
    ws.Activate
    MsgBox combinations.Count & " combinations detected.", vbInformation, "Invoice control"
End Sub

Public Sub GenererFactureDepuisLigne()
    Dim ws As Worksheet
    Dim rowNumber As Long

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    rowNumber = ws.Shapes(Application.Caller).TopLeftCell.Row
    GenererUneFacture CStr(ws.Cells(rowNumber, 1).value), _
                      CStr(ws.Cells(rowNumber, 2).value), ThisWorkbook
    ws.Cells(rowNumber, 5).value = "Generated " & Format(Now, "dd/mm/yyyy hh:nn")
    Exit Sub
Fail:
    MsgBox "Invoice generation failed (" & Err.Number & "): " & Err.Description, _
           vbCritical, "Invoice control"
End Sub

Public Sub GenererToutesLesFactures()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim doneCount As Long

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 2 To lastRow
        If Len(Trim$(CStr(ws.Cells(rowNumber, 1).value))) > 0 Then
            GenererUneFacture CStr(ws.Cells(rowNumber, 1).value), _
                              CStr(ws.Cells(rowNumber, 2).value), ThisWorkbook
            ws.Cells(rowNumber, 5).value = "Generated " & Format(Now, "dd/mm/yyyy hh:nn")
            doneCount = doneCount + 1
        End If
    Next rowNumber
    MsgBox doneCount & " invoices generated.", vbInformation, "Invoice control"
    Exit Sub
Fail:
    MsgBox "Stopped at row " & rowNumber & " (" & Err.Number & "): " & Err.Description, _
           vbCritical, "Invoice control"
End Sub

Private Function FindTrackingCombinations(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    Dim seen As Object
    Dim rowNumber As Long
    Dim program As String
    Dim lot As String
    Dim key As String

    Set seen = CreateObject("Scripting.Dictionary")
    For rowNumber = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        program = Trim$(CStr(ws.Cells(rowNumber, 1).Value2))
        lot = NormalizeInvoiceLot(program, CStr(ws.Cells(rowNumber, 2).Value2))
        If Len(program) > 0 And Len(lot) > 0 Then
            key = UCase$(program & ChrW(1) & lot)
            If Not seen.Exists(key) Then
                seen.Add key, True
                result.Add Array(program, lot)
            End If
        End If
    Next rowNumber
    Set FindTrackingCombinations = result
End Function

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

Private Function CountMatchingRows(ByVal ws As Worksheet, _
                                   ByVal program As String, _
                                   ByVal lot As String) As Long
    Dim rowNumber As Long

    For rowNumber = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If MatchesInvoiceGroup(ws, rowNumber, program, lot) Then
            CountMatchingRows = CountMatchingRows + 1
        End If
    Next rowNumber
End Function

Private Function MatchesInvoiceGroup(ByVal ws As Worksheet, _
                                     ByVal rowNumber As Long, _
                                     ByVal program As String, _
                                     ByVal lot As String) As Boolean
    MatchesInvoiceGroup = _
        StrComp(Trim$(CStr(ws.Cells(rowNumber, 1).Value2)), program, _
                vbTextCompare) = 0 And _
        StrComp(NormalizeInvoiceLot(program, CStr(ws.Cells(rowNumber, 2).Value2)), _
                lot, vbTextCompare) = 0
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

Private Function FindTrackingSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If UCase$(Left$(Trim$(ws.Name), 15)) = "TRACKING EXPORT" Then
            Set FindTrackingSheet = ws
            Exit Function
        End If
    Next ws
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
