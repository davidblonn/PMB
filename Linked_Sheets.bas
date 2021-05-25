Attribute VB_Name = "Linked_Sheets"
Option Compare Database
Option Explicit

Dim linkS As newSheet               'persistent object for Linked Sheets spreadsheet
                                    'remains open for editing through all necessary subs
                                    'as determined by iteration and value of (s)
Dim b As Boolean                    'to validate files and folder
Dim pasteLastRow As Long            'store last row values

Dim LinkedSheet As String           'Linked Sheets file path
Dim excelOpen As String             'linked sheet file name
Dim newFiles As String              'new workbook file path
Dim cProg As New clsLblProg         'progress bar
Dim txtCurrent As String

Public Function OpenLinkedSheet(b As Boolean) As newSheet
'returns a NewSheet object that opens LINKEDSHEETS spreadsheet for editing
Set OpenLinkedSheet = New newSheet
With OpenLinkedSheet
        .fType = "file"
    Set .xlApp = CreateObject("Excel.Application")
        .sourceFile = Forms!frmFilePicker!lblLinkedSheets
        .sheetName = GetFilenameFromPath(.sourceFile)
        .ValidateFileFolder (b)
            If b = False Then Exit Function
    Set .xlBook = .xlApp.Workbooks.Open(.sourceFile, ReadOnly:=False)
        .xlApp.Visible = False
        .xlApp.DisplayAlerts = False
        Debug.Print .sourceFile
End With

End Function

Public Sub PBICLinkedSheets(s As Integer)
'clears LINKEDSHEETS PBIC_TAC_ERC worksheet
'copies in new data
'updates link between excel worksheet and access linked table
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='PBIC_TAC_ERC'")) Then
        DoCmd.DeleteObject acTable, "PBIC_TAC_ERC"
    End If
    
Dim PbicTacErc As newSheet
b = True

    newFiles = Forms!frmFilePicker!lblPBIC
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

    Set PbicTacErc = New newSheet
    
    With PbicTacErc
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "PBIC_TAC_ERC"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets(.sheetName)
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:B500").Clear
            .xlPasteSheet.range("A2:B" & .lastRow).Value = _
                .xlCopySheet.range("A2:B" & .lastRow).Value
            .DataTransfer linkS
            .xlBook.Close savechanges:=False
    End With
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        Set linkS = Nothing
    Else
        s = s - 1
    End If

    Set PbicTacErc = Nothing

End Sub

Public Sub PSDsLinkedSheets(s As Integer, Optional cProg As clsLblProg)
'clears LINKEDSHEETS PSDs worksheet
'copies in new data
'updates link between excel worksheet and access linked table
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='PSDs'")) Then
        DoCmd.DeleteObject acTable, "PSDs"
    End If

Dim Psds As newSheet
Dim newFiles As String
Dim i As Integer
Dim n As Integer
Dim recordCount As Long
Dim z As String

    If s = 5 Then
        cProg.Initialize Forms!frmFilePicker!lblBack, _
            Forms!frmFilePicker!lblFront, Forms!frmFilePicker!lblCaption, _
            Forms!frmFilePicker!lblCurrent
        cProg.Max = s
    End If
    
    b = True
    
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)
    
    If Not TestIfOpen(excelOpen) Then
        Set linkS = OpenLinkedSheet(b)
            If b = False Then Exit Sub
    End If
    
    For i = 0 To 3
        Select Case i
            Case 0
                newFiles = Forms!frmFilePicker!lblTurnIn
                txtCurrent = "Turn-In PSDs"
            Case 1
                newFiles = Forms!frmFilePicker!lblOutgoing
                txtCurrent = "Outgoing PSDs"
            Case 2
                newFiles = Forms!frmFilePicker!lblIncoming
                txtCurrent = "Incoming PSDs"
            Case 3
                newFiles = Forms!frmFilePicker!lblOpenVetting
                txtCurrent = "Open Vetting PSDs"
        End Select
        
        Set Psds = New newSheet
    
        With Psds
                .fType = "file"
                .sourceFile = newFiles
                .sheetName = "PSDs"
                .ValidateFileFolder (b)
                    If b = False Then Exit Sub
            Set .xlBook = Workbooks.Open(.sourceFile)
            Set .xlCopySheet = .xlBook.Worksheets(.sheetName)
            Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
'top 2 rows not records, (- 2) for total record count
                .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row - 2
                    If i = 0 Then
'clear all old records
                        .xlPasteSheet.range("A2:AH3000").Clear
'set paste range & copy ranges, value = value
                        .xlPasteSheet.range("A2:AH" & .lastRow).Value = _
                        .xlCopySheet.range("A3:AH" & .lastRow + 2).Value
                        pasteLastRow = 2
                    Else
                        .xlPasteSheet.range("A" & pasteLastRow & _
                                            ":AH" & pasteLastRow + .lastRow - 1).Value = _
                        .xlCopySheet.range("A3:AH" & .lastRow + 2).Value
                    End If
                    Debug.Print pasteLastRow
                    Debug.Print (pasteLastRow + .lastRow - 1)
                    For n = pasteLastRow To (pasteLastRow + .lastRow - 1)
                        z = .xlPasteSheet.range("Q" & n).Value
                            If Len(z) < 9 And Not (Len(z) = 0) Then
                                Do Until Len(z) = 9
                                    z = "0" & z
                                Loop
                                    .xlPasteSheet.range("Q" & n).Value = "'" & z
                            End If
                    Next n
'all rows (+ 1) to get next empty row
                pasteLastRow = .xlPasteSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
                
                    If i = 3 Then
                        .SetFieldNames pasteLastRow
                        .DataTransfer linkS
                    End If
                .xlBook.Close savechanges:=False
        End With
        
        Set Psds = Nothing
        s = s - 1
        cProg.Increment txtCurrent
    Next i
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        cProg.Increment
        Set linkS = Nothing
        Set cProg = Nothing
    Else
        s = s - 1
    End If

End Sub

Public Sub SB700LinkedSheets(s As Integer)
'clears LINKEDSHEETS SB700 worksheet
'copies in new data
'updates link between excel worksheet and access linked table
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='SB700'")) Then
        DoCmd.DeleteObject acTable, "SB700"
    End If

Dim SB700 As newSheet
b = True

    newFiles = Forms!frmFilePicker!lblSB700
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

    Set SB700 = New newSheet
    
    With SB700
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "SB700"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets(.sheetName)
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:D1000").Clear
            .xlPasteSheet.range("A2:D" & .lastRow).Value = _
                .xlCopySheet.range("A2:D" & .lastRow).Value
            .SetFieldNames pasteLastRow
            .DataTransfer linkS
            .xlBook.Close savechanges:=False
    End With
    
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        Set linkS = Nothing
    Else
        s = s - 1
    End If
    
    Set SB700 = Nothing

End Sub

Public Sub WVARNGLinkedSheets(s As Integer, Optional cProg As clsLblProg)
'clears LINKEDSHEETS NIINDetail and PBAuth worksheets
'copies in new data
'breaksdown MATCAT code into (4) separate codes for 1st (4) positions
'concatenates PBIC-TAC-ERC codes for 2 possible codes for each row
'updates link between excel worksheets and access linked tables
'updates MATCAT and PBICTACERC tables if option is chosen on form
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='PBAuth'")) Then
        DoCmd.DeleteObject acTable, "PBAuth"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='NIINDetail'")) Then
        DoCmd.DeleteObject acTable, "NIINDetail"
    End If

Dim Wvarng As newSheet
b = True

    If s = 5 Then
        cProg.Initialize Forms!frmFilePicker!lblBack, _
            Forms!frmFilePicker!lblFront, Forms!frmFilePicker!lblCaption, _
            Forms!frmFilePicker!lblCurrent
        If Forms!frmFilePicker!optPBICMATCAT = True Then
            cProg.Max = s + 1
            s = s + 1
        Else
            cProg.Max = s
        End If
    End If

    newFiles = Forms!frmFilePicker!lblWVARNG_All
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

Dim regEx As New RegExp                         'PBIC-TAC-ERC code
Dim n As Long                                   'iterator for deleting rows
Dim e As Integer                                'loop through boolean columns
Dim p As range                                  'change "-" to "0"
Dim q As Long
Dim z As String
Dim cellTest As String                          'PBIC code to regEx test
Dim codeOne As String                           'first found code
Dim multiCode As String                         'second found code
Dim conCat As String                            'concatenated 1st code
Dim multiConCat As String                       'concatenated 2nd code
Dim i As Long                                   'row for regEx test
Dim pattern As String                           'regEx test
Dim Matcat As String                            'full MATCAT code
Dim matPos As String                            'each (4) code positions
Dim uicCat As String                            'unitUIC
Dim niinCat As String                           'item NIIN number

'*************************************************************
'********************NIINDETAIL*******************************
    Set Wvarng = New newSheet
    
    With Wvarng
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "NIINDetail"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets(1)
            .xlCopySheet.range("B:G, I:I, Q:Q, T:U, AF:AF, AJ:AN").Delete
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:AN10000").Clear
            .xlPasteSheet.range("A2:AN" & .lastRow - 1).Value = _
                .xlCopySheet.range("A3:AN" & .lastRow).Value
            .xlPasteSheet.Columns("N").NumberFormat = "@"
            .xlPasteSheet.Columns("V:W").NumberFormat = "$#,##0.00"
            pasteLastRow = .lastRow
        cProg.Increment "NIINDetail Codes"
        s = s - 1 '*************************************************************(6)
'perform operations on all rows of worksheet BOTTOM-UP
        With .xlPasteSheet
        
            For n = pasteLastRow To 2 Step -1
'delete rows based on OLD UICs
                z = .range("N" & n)
                If Len(z) < 9 And Not (Len(z) = 0) Then
                    Do Until Len(z) = 9
                        z = "0" & z
                    Loop
                        .range("N" & n).Value = "'" & z
                End If
                    
'concatenates UIC and NIIN to create unique row key to match to MATCAT table
                uicCat = .range("A" & n)
                niinCat = .range("N" & n)
                    .range("AC" & n).Value = uicCat & niinCat
                        
'parse out each MATCAT code by position
                Matcat = .range("U" & n)
                If Len(Matcat) >= 2 Then
                    matPos = Left(Matcat, 1)
                        .range("Y" & n).Value = matPos
                    matPos = Mid(Matcat, 2, 1)
                        .range("Z" & n).Value = matPos
                    matPos = Mid(Matcat, 3, 1)
                        .range("AA" & n).Value = matPos
                    matPos = Mid(Matcat, 4, 1)
                        .range("AB" & n).Value = matPos
                End If
                Matcat = ""
            Next n
            cProg.Increment "NIINDetail Save"
            s = s - 1 '***********************************************************(5)
            .range("A2").EntireRow.Insert xlDown
            .range("A2").EntireRow.Insert xlDown
            .range("A2").EntireRow.Insert xlDown
            .range("Y2:AB4").Value = "text"
            .range("N2:N4").Value = "text"
        End With
            
            .DataTransfer linkS
        
        Set .xlCopySheet = Nothing
        Set .xlPasteSheet = Nothing
            .sheetName = ""
    End With
    s = s - 1 '*******************************************************************(4)
    cProg.Increment "PBAuth Codes"

'******************************************************
'*******************PB AUTH****************************
    With Wvarng
            .sheetName = "PBAuth"
        Set .xlCopySheet = .xlBook.Worksheets(2)
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:AA10000").Clear
            .xlPasteSheet.range("A2:AA" & .lastRow - 1).Value = _
                .xlCopySheet.range("A3:AA" & .lastRow).Value
            .xlPasteSheet.Columns("K").NumberFormat = "@"
            .xlPasteSheet.Columns("W:X").NumberFormat = "$#,##0.00"
        
        With .xlPasteSheet
        
            pasteLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
'change all "-" to "0" or delete cell for proper number formatting
            For Each p In .range("F2:H" & pasteLastRow)
                p.NumberFormat = "0"
                If p.Value = "-" Then
                    p.Value = "0"
                End If
            Next p
            
            For Each p In .range("L2:L" & pasteLastRow)
                If Not IsNumeric(p.Value) Then
                    p.Clear
                End If
            Next p
            
'set regEx expression to (1) alpha-numeric char at beginning of string
            With regEx
                .Global = True
                '.Multiline = True
                .IgnoreCase = True
                .pattern = "(^[a-zA-Z0-9]{1})"
            End With
'perform operations on all rows of worksheet TOP-DOWN
            For i = 2 To pasteLastRow
            
                z = .range("N" & n)
                If Len(z) < 9 And Not (Len(z) = 0) Then
                    Do Until Len(z) = 9
                        z = "0" & z
                    Loop
                        .range("N" & n).Value = "'" & z
                End If
            
'concatenate UIC and LIN for unique row key
                Dim linCat As String
                    uicCat = .range("A" & i).Value
                    linCat = .range("D" & i).Value
                    .range("AF" & i).Value = uicCat & linCat
'reset all string variables
                    conCat = ""
                    multiConCat = ""
                    cellTest = ""
                    multiCode = ""
                    codeOne = ""
'*****PBIC search*****
                        cellTest = .range("U" & i)
'if PBIC is multiple, get both
                            If Len(cellTest) > 1 Then
                                codeOne = Mid(cellTest, 5, 1)
                                multiCode = Right(cellTest, 1)
'captures both codes together ie... "8, 3"
                                .range("AC" & i).Value = codeOne & ", " & multiCode
'test, then add 1st PBIC code to conCat
                                    If regEx.test(codeOne) Then
'"0" is added to match codes in table
                                        conCat = "0" & codeOne
                                    End If
'test, then add second PBIC to multiConCat
                                    If regEx.test(multiCode) Then
                                        multiConCat = "0" & multiCode
                                    End If
                            Else
'if PBIC is one char, test, then add to conCat
                                If regEx.test(cellTest) Then
'multiConCat matches conCat...performed to accomodate (1) PBIC code, (2) TAC code situations
                                    conCat = conCat & "0" & cellTest
                                    multiConCat = conCat
                                Else
'if no PBIC is found, assign space holder for TAC search
'"0/" is added to match codes in table
                                    conCat = "0/"
                                    multiConCat = conCat
                                End If
                            End If
                            
'*****TAC*****
                        cellTest = .range("T" & i)
'if TAC is multiple, get both
                            If Len(cellTest) > 1 Then
                                codeOne = Mid(cellTest, 5, 1)
                                multiCode = Right(cellTest, 1)
                                .range("AD" & i).Value = codeOne & ", " & multiCode
'test, then add 1st TAC code to conCat
                                    If regEx.test(codeOne) Then
                                        conCat = conCat & "0" & codeOne
                                    End If
'test, then add second TAC to multiCode
                                    If regEx.test(multiCode) Then
                                        multiConCat = multiConCat & "0" & multiCode
                                    Else
'if failed test, but PBIC was earlier assigned
                                        If Len(multiConCat) = 2 Then
                                            multiConCat = multiConCat & "0" & codeOne
                                        End If
                                    End If
                            Else
'if TAC is one char, test, then add to conCat
                                If regEx.test(cellTest) Then
                                    conCat = conCat & "0" & cellTest
'if multiConCat is not the same as conCat, add one char TAC, else clear multiConCat
'situation occurs when there is only (1) PBIC and (1) TAC code
                                    If multiConCat <> Left(conCat, 2) Then
                                        multiConCat = multiConCat & "0" & cellTest
                                    Else
                                        multiConCat = ""
                                    End If
                                End If
                            End If
                                
                        cellTest = .range("K" & i)
'test, then add ERC code to conCat
                            If regEx.test(cellTest) Then
                                If Len(conCat) > 1 And conCat <> "0/" Then
                                    conCat = conCat & cellTest
                                End If
'if multiConCat has not been cleared, add ERC to it
                                    If Len(multiConCat) > 1 And multiConCat <> "0/" Then
                                        multiConCat = multiConCat & cellTest
                                    End If
                            End If
'fill cell with first concatenated code
                                If Len(conCat) > 1 And conCat <> "0/" Then
                                    .range("AB" & i).Value = conCat
                                End If
'fill cell with second, if verified code
                                If Len(multiConCat) > 1 And multiConCat <> "0/" Then
                                    .range("AE" & i).Value = multiConCat
                                End If
            Next i
            s = s - 1 '***********************************************************(3)
            cProg.Increment "PBAuth Save"
            .range("A2").EntireRow.Insert xlDown
            .range("A2").EntireRow.Insert xlDown
            .range("A2").EntireRow.Insert xlDown
            .range("K2:K4").Value = "text"
        End With '.xlPasteSheet
            .SetFieldNames pasteLastRow
            .DataTransfer linkS
            .xlBook.Close savechanges:=False
    End With 'Wvarng
    Set Wvarng = Nothing
'matches new data from NIINDetail and PBAuth tables to MATCAT and PBICTACERC tables
'if option on form has been selected
'this operation is always performed when all LINKEDSHEETS calls are made
        If Forms!frmFilePicker!optPBICMATCAT = True Then
            MatchedPBICTACERC s '**************************************************(2)
                cProg.Increment "Matched Tables SQL"
            MatchedMATCAT s, cProg '***********************************************(1)
            Set linkS = Nothing
            Set cProg = Nothing
            Exit Sub
        End If
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        cProg.Increment
        Set linkS = Nothing
        Set cProg = Nothing
    Else
        s = s - 1
    End If
    
End Sub

Public Sub UnitsLinkedSheets(s As Integer)
'clears LINKEDSHEETS Unit worksheet
'copies in new data
'updates link between excel worksheets and access linked table
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='Units'")) Then
        DoCmd.DeleteObject acTable, "Units"
    End If

Dim Units As newSheet
b = True

    newFiles = Forms!frmFilePicker!lblUnits
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

    Set Units = New newSheet
    
    With Units
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "Units"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets(.sheetName)
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "C").End(xlUp).Row
            .xlPasteSheet.range("A2:F500").Clear
            .xlPasteSheet.range("A2:F" & .lastRow).Value = _
                .xlCopySheet.range("A2:F" & .lastRow).Value
            .DataTransfer linkS
            .xlBook.Close savechanges:=False
    End With
    
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        Set linkS = Nothing
    Else
        s = s - 1
    End If
    
    Set Units = Nothing

End Sub

Public Sub ActRegLinkedSheets(s As Integer)
'clears LINKEDSHEETS Activity Register worksheet
'copies in new data
'updates link between excel worksheets and access linked tables
DoCmd.SetWarnings False

Dim ActReg As newSheet
b = True

    newFiles = Forms!frmFilePicker!lblActivityReg
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

    Set ActReg = New newSheet
    
    With ActReg
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "NewAR"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets(1)
            .xlCopySheet.range("B:B, F:G, J:J, L:M, P:P, S:Y, AB:AB, AD:AD").Delete
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:N10000").Clear
            .xlPasteSheet.range("C2:C" & .lastRow).NumberFormat = "@"
            .xlPasteSheet.range("A2:N" & .lastRow).Value = _
                .xlCopySheet.range("A2:N" & .lastRow).Value
                .xlPasteSheet.Columns("C").NumberFormat = "@"
                .xlPasteSheet.Columns("L").NumberFormat = "@"
            .xlBook.Close savechanges:=False
        pasteLastRow = .xlPasteSheet.Cells(Rows.Count, "A").End(xlUp).Row
            .SetFieldNames pasteLastRow
            .DataTransfer linkS
    End With
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        Set linkS = Nothing
    Else
        s = s - 1
    End If
    
    Set ActReg = Nothing

End Sub

Public Sub MATCATLinkedSheets(s As Integer)
'clears LINKEDSHEETS MATCAT worksheet
'copies in new data
'updates link between excel worksheets and access linked tables
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='MATCAT'")) Then
        DoCmd.DeleteObject acTable, "MATCAT"
    End If

Dim Matcat As newSheet
b = True

    newFiles = Forms!frmFilePicker!lblMATCAT
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

    Set Matcat = New newSheet
    
    With Matcat
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "MATCAT"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets(.sheetName)
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:AC10000").Clear
            .xlPasteSheet.range("A2:C" & .lastRow).Value = _
                .xlCopySheet.range("A2:C" & .lastRow).Value
            .DataTransfer linkS
            .xlBook.Close savechanges:=False
    End With
    
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        Set linkS = Nothing
    Else
        s = s - 1
    End If
    
    Set Matcat = Nothing

End Sub

Public Sub MVMT_TYPELinkedSheets(s As Integer)
'clears LINKEDSHEETS MVMT_TYPE worksheet
'copies in new data
'updates link between excel worksheets and access linked tables
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='MVMT_TYPE'")) Then
        DoCmd.DeleteObject acTable, "MVMT_TYPE"
    End If

Dim Mvmttype As newSheet
b = True

    newFiles = Forms!frmFilePicker!lblMVMT_TYPE
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

    Set Mvmttype = New newSheet
    
    With Mvmttype
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "MVMT_TYPE"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets(.sheetName)
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:AB1000").Clear
            .xlPasteSheet.range("A5:B" & .lastRow + 3).Value = _
                .xlCopySheet.range("A2:B" & .lastRow).Value
            .xlPasteSheet.Columns("A").NumberFormat = "@"
            .DataTransfer linkS
            .xlBook.Close savechanges:=False
    End With
    
'run SQL to UPDATE field in table
'*****should be moved to building activity register worksheet*********
        DoCmd.RunSQL "UPDATE tblActRegAppend " & _
                "INNER JOIN MVMT_TYPE " & _
                "ON tblActRegAppend.MVMT_TYPE = MVMT_TYPE.MVMT_Code " & _
                "SET [tblActRegAppend].[MVMT_Text] = [MVMT_TYPE].[MVMT-Text];"
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        Set linkS = Nothing
    Else
        s = s - 1
    End If
    
    Set Mvmttype = Nothing

End Sub

Public Sub ASIOE_CMILinkedSheets(s As Integer)
'clears LINKEDSHEETS PBIC_TAC_ERC worksheet
'copies in new data
'updates link between excel worksheet and access linked table
DoCmd.SetWarnings False

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='ASIOE_CMI'")) Then
        DoCmd.DeleteObject acTable, "ASIOE_CMI"
    End If
    
Dim Asioecmi As newSheet
b = True

    newFiles = Forms!frmFilePicker!lblASIOE
    LinkedSheet = Forms!frmFilePicker!lblLinkedSheets
    excelOpen = GetFilenameFromPath(LinkedSheet)

        If Not TestIfOpen(excelOpen) Then
                Set linkS = OpenLinkedSheet(b)
                    If b = False Then Exit Sub
        End If

    Set Asioecmi = New newSheet
    
    With Asioecmi
            .fType = "file"
            .sourceFile = newFiles
            .sheetName = "ASIOE_CMI"
            .ValidateFileFolder (b)
                If b = False Then Exit Sub
        Set .xlBook = Workbooks.Open(.sourceFile)
        Set .xlCopySheet = .xlBook.Worksheets("Sheet1")
            .xlCopySheet.Rows("1:2").Delete shift:=xlShiftUp
        Set .xlPasteSheet = linkS.xlBook.Worksheets(.sheetName)
            .lastRow = .xlCopySheet.Cells(Rows.Count, "A").End(xlUp).Row
            .xlPasteSheet.range("A2:G" & .lastRow + 500).Clear
            .xlPasteSheet.range("A2:G" & .lastRow).Value = _
                .xlCopySheet.range("A2:G" & .lastRow).Value
            
            'With .xlPasteSheet
            '    .range("A1").Value = "Prm_LIN"
            '    .range("B1").Value = "Prm_Nomen"
            '    .range("C1").Value = "Prm_Qty"
            '    .range("D1").Value = "Type"
            '    .range("E1").Value = "Asc_LIN"
            '    .range("F1").Value = "Asc_Nomen"
            '    .range("G1").Value = "Asc_Qty"
            'End With
            
            .DataTransfer linkS
            .xlBook.Close savechanges:=False
    End With
    
    If s = 1 Then
        SaveExcel linkS.xlBook, "linkedSheet", , , linkS
        Set linkS = Nothing
    Else
        s = s - 1
    End If

    Set Asioecmi = Nothing

End Sub

Public Sub MatchedPBICTACERC(s As Integer)
'deletes all records from current table
'adds new records from PBAuth table joined on 1st and 2nd PBIC-TAC-ERC codes
DoCmd.SetWarnings False

Dim sql As String
'run DELETE * SQL
        DoCmd.RunSQL "DELETE * FROM tblMatchedPBICTACERC;"
'records from PBAuth with matching FIRST PBIC-TAC-ERC codes
        sql = "INSERT INTO tblMatchedPBICTACERC (UIC_LIN, FoundPTE, MatchedPTE, [Desc], UIC, LIN, [Multi-PBIC], [Multi-TAC]) " & _
                    "SELECT UIC_LIN, FoundPTE, [Matl Ind#], Description, PBUIC, PBAuthLin, [Multi-PBIC], [Multi-TAC] " & _
                    "FROM PBIC_TAC_ERC " & _
                    "INNER JOIN PBAuth " & _
                    "ON PBIC_TAC_ERC.[Matl Ind#] = PBAuth.FoundPTE;"

        DoCmd.RunSQL sql
'records from PBAuth with matching SECOND PBIC-TAC-ERC codes
        sql = "INSERT INTO tblMatchedPBICTACERC (UIC_LIN, Found2ndPTE, Matched2ndPTE, [Desc_2], UIC, LIN, [Multi-PBIC], [Multi-TAC]) " & _
                    "SELECT UIC_LIN, SecondFoundPTE, [Matl Ind#], Description, PBUIC, PBAuthLin, [Multi-PBIC], [Multi-TAC] " & _
                    "FROM PBIC_TAC_ERC " & _
                    "INNER JOIN PBAuth " & _
                    "ON PBIC_TAC_ERC.[Matl Ind#] = PBAuth.SecondFoundPTE;"

        DoCmd.RunSQL sql
        s = s - 1

End Sub

Public Sub MatchedMATCAT(s As Integer, Optional cProg As clsLblProg)
'deletes all records from current table
'adds new records from NIINDetail table joined on 1st, 2nd, 3rd, & 4th MATCAT codes
DoCmd.SetWarnings False

Dim sql As String
'run DELETE * SQL
        DoCmd.RunSQL "DELETE * FROM tblMatchedMATCAT;"
'records from NIINDetail with matching FIRST MATCAT code
        sql = "INSERT INTO tblMatchedMATCAT (UIC_NIIN, Code_Position, MATCAT_Pos, Found_Code, [Desc]) " & _
                "SELECT UIC_NIIN, POS, Mat_Pos_1, POS_CODE, POS_DESC " & _
                "FROM MATCAT " & _
                "INNER JOIN NIINDetail " & _
                "ON MATCAT.POS_CODE = NIINDetail.MAT_Pos_1 " & _
                "WHERE (MATCAT.POS = 1);"
        DoCmd.RunSQL sql
'matching 2nd
        sql = "INSERT INTO tblMatchedMATCAT (UIC_NIIN, Code_Position, MATCAT_Pos, Found_Code, [Desc]) " & _
                "SELECT UIC_NIIN, POS, MAT_Pos_2, POS_CODE, POS_DESC " & _
                "FROM MATCAT " & _
                "INNER JOIN NIINDetail " & _
                "ON MATCAT.POS_CODE = NIINDetail.MAT_Pos_2 " & _
                "WHERE (MATCAT.POS = 2);"
        DoCmd.RunSQL sql
'matching 3rd
        sql = "INSERT INTO tblMatchedMATCAT (UIC_NIIN, Code_Position, MATCAT_Pos, Found_Code, [Desc]) " & _
                "SELECT UIC_NIIN, POS, MAT_Pos_3, POS_CODE, POS_DESC " & _
                "FROM MATCAT " & _
                "INNER JOIN NIINDetail " & _
                "ON MATCAT.POS_CODE = NIINDetail.MAT_Pos_3 " & _
                "WHERE (MATCAT.POS = 3);"
        DoCmd.RunSQL sql
'matching 4th
        sql = "INSERT INTO tblMatchedMATCAT (UIC_NIIN, Code_Position, MATCAT_Pos, Found_Code, [Desc]) " & _
                "SELECT UIC_NIIN, POS, MAT_Pos_4, POS_CODE, POS_DESC " & _
                "FROM MATCAT " & _
                "INNER JOIN NIINDetail " & _
                "ON MATCAT.POS_CODE = NIINDetail.MAT_Pos_4 " & _
                "WHERE (MATCAT.POS = 4);"
        DoCmd.RunSQL sql
        
        If Forms!frmFilePicker!optPBICMATCAT = True Then
            Forms!frmFilePicker!optPBICMATCAT = False
        End If
        
        If s = 1 Then
            SaveExcel linkS.xlBook, "linkedSheet", , , linkS
            cProg.Increment
            Set linkS = Nothing
            Set cProg = Nothing
            s = s - 1
        End If

End Sub
