Attribute VB_Name = "ModuleGenerationFactures"
Option Explicit

Private Const INPUT_SHEET_KEY As String = "remplir"
Private Const FINAL_SHEET As String = "Facture finale"
Private Const SOURCE_SHEET_KEY As String = "prix de vente"
Private Const CONTROL_SHEET As String = "Controle Factures"
Private Const TRACKING_FILE_PATTERN As String = "TRACKING EXPORT*.xlsx"
Private Const PRICE_FILE_PATTERN As String = "prix*.xlsx"
Private Const OUTPUT_FOLDER As String = "Factures Generated"
Private Const INVOICE_NUMBER_PREFIX As String = "LAT21"
Private Const FINAL_FIRST_ROW As Long = 8

Public Sub GenererUneFacture(ByVal program As String, _
                             ByVal invoiceGroup As String, _
                             ByVal targetLots As Collection, _
                             ByVal template As Workbook)
    Dim tracking As Workbook
    Dim source As Workbook
    Dim wsTracking As Worksheet
    Dim wsSource As Worksheet
    Dim generated As Workbook
    Dim trackingPath As String
    Dim sourcePath As String
    Dim trackingWasOpen As Boolean
    Dim sourceWasOpen As Boolean
    Dim outFolder As String
    Dim baseName As String
    Dim tempPath As String
    Dim finalPath As String
    Dim errNumber As Long
    Dim errText As String
    Dim oldScreen As Boolean
    Dim coisage As String

    oldScreen = Application.ScreenUpdating

    trackingPath = FindWorkbookPath(template.Path, TRACKING_FILE_PATTERN, template.Name)
    sourcePath = FindWorkbookPath(template.Path, PRICE_FILE_PATTERN, template.Name)
    If Len(trackingPath) = 0 Then
        Err.Raise vbObjectError + 2200, , _
            "No file matching '" & TRACKING_FILE_PATTERN & "' found in: " & template.Path
    End If
    If Len(sourcePath) = 0 Then
        Err.Raise vbObjectError + 2201, , _
            "No file matching '" & PRICE_FILE_PATTERN & "' found in: " & template.Path
    End If

    trackingWasOpen = IsWorkbookOpen(trackingPath)
    sourceWasOpen = IsWorkbookOpen(sourcePath)

    Set tracking = OpenReadOnlyWorkbook(trackingPath)
    Set wsTracking = FindTrackingSheet(tracking)
    Set source = OpenReadOnlyWorkbook(sourcePath)
    Set wsSource = FindSheet(source, SOURCE_SHEET_KEY)
    If wsTracking Is Nothing Then
        Err.Raise vbObjectError + 2202, , "Sheet 'TRACKING EXPORT...' not found in " & tracking.Name
    End If
    If wsSource Is Nothing Then
        Err.Raise vbObjectError + 2203, , "Sheet containing '" & SOURCE_SHEET_KEY & _
                                          "' not found in " & source.Name
    End If

    coisage = AskForCoisage(template, program, invoiceGroup)

    outFolder = template.Path & Application.PathSeparator & OUTPUT_FOLDER
    EnsureFolder outFolder

    baseName = SafeFileName(program & "_" & invoiceGroup)
    finalPath = outFolder & Application.PathSeparator & baseName & ".xlsx"
    tempPath = outFolder & Application.PathSeparator & "~temp_" & baseName & _
               FileExtension(template.Name)

    CloseIfOpen finalPath
    CloseIfOpen tempPath
    If Len(Dir(finalPath)) > 0 Then Kill finalPath
    If Len(Dir(tempPath)) > 0 Then Kill tempPath

    On Error GoTo Fail
    Application.ScreenUpdating = False

    template.SaveCopyAs tempPath
    Set generated = Workbooks.Open(tempPath, UpdateLinks:=False)
    FillGeneratedInvoice generated, wsTracking, wsSource, program, invoiceGroup, targetLots, coisage
    DeleteSheetIfExists generated, CONTROL_SHEET

    Application.DisplayAlerts = False
    generated.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    generated.Close SaveChanges:=False
    Set generated = Nothing
    If Len(Dir(tempPath)) > 0 Then Kill tempPath

    If Not trackingWasOpen And Not tracking Is Nothing Then tracking.Close SaveChanges:=False
    If Not sourceWasOpen And Not source Is Nothing Then source.Close SaveChanges:=False

    Application.ScreenUpdating = oldScreen
    Exit Sub

Fail:
    errNumber = Err.Number
    errText = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not generated Is Nothing Then generated.Close SaveChanges:=False
    If Len(Dir(tempPath)) > 0 Then Kill tempPath
    If Not trackingWasOpen And Not tracking Is Nothing Then tracking.Close SaveChanges:=False
    If Not sourceWasOpen And Not source Is Nothing Then source.Close SaveChanges:=False
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise errNumber, , errText
End Sub

Private Sub FillGeneratedInvoice(ByVal facture As Workbook, _
                                 ByVal wsTracking As Worksheet, _
                                 ByVal wsSource As Worksheet, _
                                 ByVal program As String, _
                                 ByVal invoiceGroup As String, _
                                 ByVal targetLots As Collection, _
                                 ByVal coisage As String)
    Dim wsInput As Worksheet
    Dim wsFinal As Worksheet
    Dim prices As Object
    Dim totals As Object
    Dim vusByMsn As Object
    Dim lotByMsn As Object
    Dim occurrences As Object
    Dim msnOrder As Collection
    Dim vus As Collection
    Dim inputRow As Long
    Dim summaryRow As Long
    Dim finalRow As Long
    Dim templateLastRow As Long
    Dim activeLastRow As Long
    Dim msn As String
    Dim vu As String
    Dim msnLot As String
    Dim code As String
    Dim price As Double
    Dim foundPrice As Boolean
    Dim lotComplete As Boolean
    Dim lotCompleteByMsn As Object
    Dim articleCodesByMsn As Object
    Dim articleCodes As Collection
    Dim item As Variant
    Dim inputName As String

    Set wsInput = FindSheet(facture, INPUT_SHEET_KEY)
    Set wsFinal = FindSheet(facture, FINAL_SHEET)
    If wsInput Is Nothing Then
        Err.Raise vbObjectError + 2204, , "Input sheet (name containing '" & _
                                          INPUT_SHEET_KEY & "') not found in the template."
    End If
    If wsFinal Is Nothing Then
        Err.Raise vbObjectError + 2205, , "Sheet '" & FINAL_SHEET & "' not found in the template."
    End If
    inputName = wsInput.Name

    Set prices = CreateObject("Scripting.Dictionary")
    prices.CompareMode = vbTextCompare
    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare
    Set vusByMsn = CreateObject("Scripting.Dictionary")
    vusByMsn.CompareMode = vbTextCompare
    Set lotByMsn = CreateObject("Scripting.Dictionary")
    lotByMsn.CompareMode = vbTextCompare
    Set occurrences = CreateObject("Scripting.Dictionary")
    occurrences.CompareMode = vbTextCompare
    Set lotCompleteByMsn = CreateObject("Scripting.Dictionary")
    lotCompleteByMsn.CompareMode = vbTextCompare
    Set articleCodesByMsn = CreateObject("Scripting.Dictionary")
    articleCodesByMsn.CompareMode = vbTextCompare
    Set msnOrder = New Collection

    LoadPrices wsSource, prices
    templateLastRow = DetectFinalTableLastRow(wsFinal)
    ClearInvoiceData wsInput, wsFinal, templateLastRow

    inputRow = 2
    For Each item In GetMatchingRows(wsTracking, program, targetLots)
        msnLot = NormalizeInvoiceLot(program, CStr(wsTracking.Cells(CLng(item), 2).Value2))
        msn = Trim$(CStr(wsTracking.Cells(CLng(item), 3).Value2))
        vu = Trim$(CStr(wsTracking.Cells(CLng(item), 4).Value2))
        code = NormalizeCode(wsTracking.Cells(CLng(item), 6).Value2)
        If Len(msn) > 0 And Len(vu) > 0 And Len(code) > 0 Then
            If Not totals.Exists(msn) Then
                totals.Add msn, 0#
                occurrences.Add msn, 0
                lotCompleteByMsn.Add msn, False
                lotByMsn.Add msn, msnLot
                Set vus = New Collection
                vusByMsn.Add msn, vus
                Set articleCodes = New Collection
                articleCodesByMsn.Add msn, articleCodes
                msnOrder.Add msn
            Else
                Set vus = vusByMsn(msn)
                Set articleCodes = articleCodesByMsn(msn)
            End If
            occurrences(msn) = CLng(occurrences(msn)) + 1

            SetCell wsInput.Cells(inputRow, 1), msn & "-" & occurrences(msn), False
            SetCell wsInput.Cells(inputRow, 2), msn, False
            SetCell wsInput.Cells(inputRow, 3), vu, False

            AddUnique vus, vu
            AddUnique articleCodes, code
            price = LookupPrice(code, prices, foundPrice, lotComplete)
            If foundPrice Then
                If lotComplete Then
                    If Not CBool(lotCompleteByMsn(msn)) Then
                        totals(msn) = 0#
                        lotCompleteByMsn(msn) = True
                    End If
                    totals(msn) = CDbl(totals(msn)) + price
                ElseIf Not CBool(lotCompleteByMsn(msn)) Then
                    totals(msn) = CDbl(totals(msn)) + price
                End If
            End If
            inputRow = inputRow + 1
        End If
    Next item

    activeLastRow = AdjustFinalTableRows(wsFinal, templateLastRow, msnOrder.Count)

    summaryRow = 2
    finalRow = FINAL_FIRST_ROW
    For Each item In msnOrder
        SetCell wsInput.Cells(summaryRow, 5), _
                BuildReference(program, CStr(lotByMsn(CStr(item))), vusByMsn(CStr(item))), False
        SetCell wsInput.Cells(summaryRow, 6), CStr(item), False
        SetCell wsInput.Cells(summaryRow, 7), Abs(CDbl(totals(CStr(item)))), False

        SetCell wsFinal.Cells(finalRow, 2), "='" & inputName & "'!E" & summaryRow, True
        SetCell wsFinal.Cells(finalRow, 9), "='" & inputName & "'!F" & summaryRow, True
        SetCell wsFinal.Cells(finalRow, 10), _
                CollectionAsLines(articleCodesByMsn(CStr(item))), False
        SetCell wsFinal.Cells(finalRow, 12), "='" & inputName & "'!G" & summaryRow, True
        SetCell wsFinal.Cells(finalRow, 13), "=L" & finalRow, True

        summaryRow = summaryRow + 1
        finalRow = finalRow + 1
    Next item

    If finalRow <= activeLastRow Then
        wsFinal.Rows(finalRow & ":" & activeLastRow).Hidden = True
    End If

    ClearRangeSafe wsInput.Range("I:OI")
    wsInput.Range("I:OI").ClearContents
    SetCell wsFinal.Cells(4, 3), coisage, False
    SetCell wsFinal.Cells(3, 2), _
            "N° Facture : " & BuildInvoiceNumber(coisage, wsFinal.Cells(4, 14).Value), _
            False
End Sub

Private Function DetectFinalTableLastRow(ByVal wsFinal As Worksheet) As Long
    Dim r As Long
    Dim maxScan As Long
    Dim rowText As String
    Dim colMFormula As String
    Dim colBText As String

    maxScan = wsFinal.UsedRange.Row + wsFinal.UsedRange.Rows.Count + 5
    If maxScan < FINAL_FIRST_ROW + 1 Then maxScan = FINAL_FIRST_ROW + 30

    For r = FINAL_FIRST_ROW To maxScan
        colMFormula = UCase$(Trim$(CStr(wsFinal.Cells(r, 13).Formula)))
        colBText = UCase$(Trim$(CStr(wsFinal.Cells(r, 2).Value2)))
        rowText = UCase$(Trim$(CStr(wsFinal.Cells(r, 10).Value2) & " " & _
                               CStr(wsFinal.Cells(r, 11).Value2) & " " & _
                               CStr(wsFinal.Cells(r, 12).Value2)))

        If Left$(colMFormula, 5) = "=SUM(" Or Left$(colMFormula, 7) = "=SOMME(" Or _
           InStr(1, rowText, "TOTAL", vbTextCompare) > 0 Or _
           InStr(1, colBText, "TOTAL", vbTextCompare) > 0 Then
            DetectFinalTableLastRow = r - 1
            Exit Function
        End If
    Next r

    DetectFinalTableLastRow = Application.WorksheetFunction.Max( _
        FINAL_FIRST_ROW, _
        wsFinal.Cells(wsFinal.Rows.Count, 2).End(xlUp).Row, _
        wsFinal.Cells(wsFinal.Rows.Count, 9).End(xlUp).Row, _
        wsFinal.Cells(wsFinal.Rows.Count, 12).End(xlUp).Row)
End Function

Private Function AdjustFinalTableRows(ByVal wsFinal As Worksheet, _
                                      ByVal templateLastRow As Long, _
                                      ByVal neededCount As Long) As Long
    Dim capacity As Long
    Dim extraRows As Long
    Dim insertAt As Long
    Dim newLastRow As Long
    Dim lastDataRow As Long

    If templateLastRow < FINAL_FIRST_ROW Then templateLastRow = FINAL_FIRST_ROW
    wsFinal.Rows(FINAL_FIRST_ROW & ":" & templateLastRow).Hidden = False

    capacity = templateLastRow - FINAL_FIRST_ROW + 1
    newLastRow = templateLastRow

    If neededCount > capacity Then
        extraRows = neededCount - capacity
        insertAt = templateLastRow + 1
        newLastRow = templateLastRow + extraRows

        wsFinal.Rows(insertAt).Resize(extraRows).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        wsFinal.Rows(templateLastRow).Copy
        wsFinal.Rows(insertAt & ":" & newLastRow).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        wsFinal.Rows(insertAt & ":" & newLastRow).RowHeight = wsFinal.Rows(templateLastRow).RowHeight

        ClearRangeSafe wsFinal.Range("B" & insertAt & ":B" & newLastRow)
        ClearRangeSafe wsFinal.Range("I" & insertAt & ":J" & newLastRow)
        ClearRangeSafe wsFinal.Range("L" & insertAt & ":M" & newLastRow)
    End If

    If neededCount > 0 Then
        lastDataRow = FINAL_FIRST_ROW + neededCount - 1
    Else
        lastDataRow = FINAL_FIRST_ROW
    End If

    UpdateFooterSumFormulas wsFinal, newLastRow, lastDataRow
    AdjustFinalTableRows = newLastRow
End Function

Private Sub UpdateFooterSumFormulas(ByVal wsFinal As Worksheet, _
                                    ByVal newLastRow As Long, _
                                    ByVal lastDataRow As Long)
    Dim footerRow As Long
    Dim col As Long
    Dim fText As String
    Dim colLetter As String

    For footerRow = newLastRow + 1 To newLastRow + 15
        For col = 1 To 15
            fText = Trim$(CStr(wsFinal.Cells(footerRow, col).Formula))
            If UCase$(Left$(fText, 5)) = "=SUM(" Or UCase$(Left$(fText, 7)) = "=SOMME(" Then
                colLetter = ColumnIndexToLetter(col)
                If col = 14 Then colLetter = "M"
                wsFinal.Cells(footerRow, col).Formula = _
                    "=SUM(" & colLetter & FINAL_FIRST_ROW & ":" & colLetter & lastDataRow & ")"
            End If
        Next col
    Next footerRow
End Sub

Private Function ColumnIndexToLetter(ByVal colIndex As Long) As String
    Dim parts As Variant
    parts = Split(Cells(1, colIndex).Address(True, False), "$")
    ColumnIndexToLetter = parts(0)
End Function

Private Function GetMatchingRows(ByVal ws As Worksheet, _
                                 ByVal program As String, _
                                 ByVal targetLots As Collection) As Collection
    Dim result As New Collection
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim normalizedLot As String
    Dim targetLot As Variant
    Dim isMatch As Boolean

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowNumber, 1).Value2)), program, vbTextCompare) = 0 Then
            normalizedLot = NormalizeInvoiceLot(program, CStr(ws.Cells(rowNumber, 2).Value2))
            isMatch = False
            For Each targetLot In targetLots
                If StrComp(normalizedLot, CStr(targetLot), vbTextCompare) = 0 Then
                    isMatch = True
                    Exit For
                End If
            Next targetLot
            If isMatch Then result.Add rowNumber
        End If
    Next rowNumber
    Set GetMatchingRows = result
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

Private Sub ClearInvoiceData(ByVal wsInput As Worksheet, _
                             ByVal wsFinal As Worksheet, _
                             ByVal templateLastRow As Long)
    Dim lastInputRow As Long

    lastInputRow = wsInput.UsedRange.Row + wsInput.UsedRange.Rows.Count - 1
    If lastInputRow < 2 Then lastInputRow = 2
    ClearRangeSafe wsInput.Range("A2:G" & lastInputRow)

    If templateLastRow < FINAL_FIRST_ROW Then templateLastRow = FINAL_FIRST_ROW
    wsFinal.Rows(FINAL_FIRST_ROW & ":" & templateLastRow).Hidden = False
    ClearRangeSafe wsFinal.Range("B" & FINAL_FIRST_ROW & ":B" & templateLastRow)
    ClearRangeSafe wsFinal.Range("I" & FINAL_FIRST_ROW & ":I" & templateLastRow)
    ClearRangeSafe wsFinal.Range("J" & FINAL_FIRST_ROW & ":J" & templateLastRow)
    ClearRangeSafe wsFinal.Range("L" & FINAL_FIRST_ROW & ":M" & templateLastRow)
End Sub

Private Sub LoadPrices(ByVal ws As Worksheet, ByVal prices As Object)
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim code As String
    Dim values As Variant

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 2 To lastRow
        code = NormalizeCode(ws.Cells(rowNumber, 1).Value2)
        If Len(code) > 0 Then
            If Not prices.Exists(code) Then
                'values(0) = billing price (I), values(1) = per VU (H),
                'values(2) = Nature (F), values(3) = Remarque (J).
                values = Array(ws.Cells(rowNumber, 9).Value2, _
                               ws.Cells(rowNumber, 8).Value2, _
                               ws.Cells(rowNumber, 6).Value2, _
                               ws.Cells(rowNumber, 10).Value2)
                prices.Add code, values
            End If
        End If
    Next rowNumber
End Sub

Private Function NormalizeCode(ByVal value As Variant) As String
    Dim code As String

    If IsError(value) Or IsEmpty(value) Then Exit Function
    code = UCase$(Trim$(Replace(CStr(value), ChrW(160), " ")))
    If Left$(code, 2) = "IS" Then NormalizeCode = code
End Function

Private Function LookupPrice(ByVal code As String, ByVal prices As Object, _
                             ByRef found As Boolean, _
                             ByRef lotComplete As Boolean) As Double
    Dim values As Variant

    found = False
    lotComplete = False
    If Not prices.Exists(code) Then Exit Function
    values = prices(code)
    lotComplete = IsLotComplete(values(2), values(3))
    If IsNumericValue(values(0)) Then
        LookupPrice = Abs(CDbl(values(0)))
        found = True
    ElseIf IsNumericValue(values(1)) Then
        LookupPrice = Abs(CDbl(values(1)))
        found = True
    End If
End Function

Private Function IsLotComplete(ByVal nature As Variant, _
                               ByVal remarque As Variant) As Boolean
    Dim text As String

    text = UCase$(Trim$(CStr(nature)) & " " & Trim$(CStr(remarque)))
    IsLotComplete = InStr(1, text, "LOT", vbTextCompare) > 0 And _
                    InStr(1, text, "COMPLET", vbTextCompare) > 0
End Function

Private Function IsNumericValue(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    IsNumericValue = IsNumeric(value)
End Function

Private Function AskForCoisage(ByVal template As Workbook, _
                               ByVal program As String, _
                               ByVal invoiceGroup As String) As String
    Dim defaultValue As String
    Dim answer As String

    On Error Resume Next
    defaultValue = CStr(FindSheet(template, FINAL_SHEET).Cells(4, 3).Value)
    On Error GoTo 0
    answer = InputBox( _
        "Saisir le coisage / packing list pour [" & program & " - " & invoiceGroup & "] :", _
        "Informations facture", defaultValue)
    If Len(Trim$(answer)) = 0 Then
        Err.Raise vbObjectError + 2206, , _
            "La génération est annulée : le coisage est obligatoire."
    End If
    AskForCoisage = Trim$(answer)
End Function

Private Function BuildInvoiceNumber(ByVal coisage As String, _
                                    ByVal shippedDate As Variant) As String
    Dim parts As Variant
    Dim sequenceText As String
    Dim sequenceNumber As Long
    Dim shippedDay As String

    If Not IsDate(shippedDate) Then
        Err.Raise vbObjectError + 2207, , _
            "La date d'expédition en N4 est invalide."
    End If

    parts = Split(Trim$(coisage), "-")
    If UBound(parts) < 2 Then
        Err.Raise vbObjectError + 2208, , _
            "Format de coisage invalide. Exemple attendu : CAS26-11-025ID."
    End If

    sequenceText = DigitsOnly(CStr(parts(2)))
    If Len(sequenceText) = 0 Then
        Err.Raise vbObjectError + 2209, , _
            "Aucun numéro de séquence trouvé dans le coisage : " & coisage
    End If
    sequenceNumber = CLng(sequenceText)
    shippedDay = Format$(CDate(shippedDate), "dd")
    BuildInvoiceNumber = INVOICE_NUMBER_PREFIX & shippedDay & sequenceNumber
End Function

Private Function DigitsOnly(ByVal value As String) As String
    Dim index As Long
    Dim character As String

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If character >= "0" And character <= "9" Then
            DigitsOnly = DigitsOnly & character
        End If
    Next index
End Function

Private Function CollectionAsLines(ByVal values As Collection) As String
    Dim item As Variant

    For Each item In values
        If Len(CollectionAsLines) > 0 Then CollectionAsLines = _
            CollectionAsLines & vbLf
        CollectionAsLines = CollectionAsLines & CStr(item)
    Next item
End Function

Private Function BuildReference(ByVal program As String, ByVal lot As String, _
                                ByVal vus As Collection) As String
    Dim item As Variant
    Dim text As String

    For Each item In vus
        If Len(text) > 0 Then text = text & "-"
        text = text & CStr(item)
    Next item
    If InStr(1, program, "-CCP", vbTextCompare) > 0 Then
        BuildReference = Replace(program, "-CCP", "", 1, -1, vbTextCompare) & _
                         "-" & text & "-CCP " & UCase$(lot)
    Else
        BuildReference = program & "-" & text & "-" & UCase$(lot)
    End If
End Function

Private Sub AddUnique(ByVal values As Collection, ByVal value As String)
    On Error Resume Next
    values.Add value, UCase$(value)
    On Error GoTo 0
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function SafeFileName(ByVal value As String) As String
    Dim invalid As Variant

    For Each invalid In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        value = Replace(value, CStr(invalid), "_")
    Next invalid
    SafeFileName = Trim$(value)
End Function

Private Function FileExtension(ByVal fileName As String) As String
    Dim p As Long

    p = InStrRev(fileName, ".")
    If p > 0 Then FileExtension = Mid$(fileName, p) Else FileExtension = ".xlsx"
End Function

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            If wb.Worksheets.Count > 1 Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
            End If
            Exit Sub
        End If
    Next ws
End Sub

Private Sub CloseIfOpen(ByVal fullPath As String)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub

Private Function FindTrackingSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If UCase$(Left$(Trim$(ws.Name), 15)) = "TRACKING EXPORT" Then
            Set FindTrackingSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function FindSheet(ByVal wb As Workbook, ByVal key As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), Trim$(key), vbTextCompare) = 0 Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, key, vbTextCompare) > 0 Then
            Set FindSheet = ws
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

Private Sub ClearRangeSafe(ByVal rng As Range)
    Dim target As Range
    Dim c As Range

    On Error Resume Next
    rng.ClearContents
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    Set target = Intersect(rng, rng.Worksheet.UsedRange)
    If target Is Nothing Then Exit Sub
    For Each c In target.Cells
        If c.MergeCells Then
            c.MergeArea.ClearContents
        Else
            c.ClearContents
        End If
    Next c
End Sub

Private Sub SetCell(ByVal c As Range, ByVal v As Variant, ByVal asFormula As Boolean)
    If c.MergeCells Then
        If c.Address <> c.MergeArea.Cells(1, 1).Address Then Exit Sub
    End If
    If asFormula Then
        c.Formula = v
    Else
        c.Value = v
    End If
End Sub