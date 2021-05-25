Attribute VB_Name = "Office_Reports"
Option Compare Database
Option Explicit

Public newReport As Form_frmOffReports
Public curForm As Form

Public Sub ReportsInitialize(openForm As Form)
    
    Set curForm = openForm
    Set newReport = New Form_frmOffReports
    
    With newReport
        .optUG = unitSelect.optUG
        .strIN = unitSelect.sql
        .sheetName = curForm.comboReportPicker          'report chosen
        .filePath = curForm.lblReportFolder             'save folder
        .fileName = .filePath & "\" & _
                    .sheetName                          'full file path
        .objType = "psdReport"                          'report type
        .rsSQL = SetSQL                                 'full sql
'create recordset of results
        Set .rsReport = CurrentDb.OpenRecordset(.rsSQL, dbOpenDynaset)
        Set .xlBook = BuildSpreadsheet
    End With
    
    SaveExcel newReport.xlBook, newReport.objType, newReport
    
    Set curForm = Nothing

End Sub

Public Sub StaffInitialize(openForm As Form)
    
    Set curForm = openForm
    Set newReport = New Form_frmOffReports

Dim xlSheetNames As Variant
    xlSheetNames = Array("In_State_LT", "Outgoing_LT", "Incoming_LT", "Open_Vet", _
                        "Turn_In", "Await_PU", "Staff_Report")
                        
Dim i As Integer                'iterate thru array
'loop through array
        For i = LBound(xlSheetNames) To UBound(xlSheetNames)
'build getReport object
             With newReport
                    .sheetName = xlSheetNames(i)            'current excel sheet
                    .fileName = curForm.lblStaffReport      'report template
                    .filePath = curForm.lblReportFolder     'save path
                    .objType = "staffReport"                'report type
                
                    If xlSheetNames(i) = "Staff_Report" Then
'last sheet executed
'creates recordset from temp table of previous report
                            Set .rsReport = _
                                CurrentDb.OpenRecordset("tblCompareStaffReport", dbOpenDynaset)
                            Set .xlBook = BuildSpreadsheet         'append to excel worksheet
                            SaveExcel newReport.xlBook, newReport.objType, newReport  'save report template with new name
                            Set curForm = Nothing
                    Else
'set SQL
                            .rsSQL = SetSQL
'create recordset
                            Set .rsReport = _
                                CurrentDb.OpenRecordset(.rsSQL, dbOpenDynaset)
                            Set .xlBook = BuildSpreadsheet         'append to excel worksheet
                    End If
            End With
            
        Next i

End Sub

Public Function SetSQL() As String

Dim andSQL As String            'sql segment
Dim toFrom As String            'sql segment
'selection based on sheet name for excel workbook
'chosen either from from COMBOBOX or assigned during STAFF REPORT
    Select Case newReport.sheetName
'(4) COMBOBOX options
        Case "Outgoing LT", "Incoming LT", "Turn-In", "Open Vetting"
'inner select assigns sql segments to specific report type
            Select Case newReport.sheetName
                Case "Outgoing LT"
                    andSQL = "AND (Type = 'Lateral Transfer') " & _
                            "AND (ToState <> 'WV') "
                    toFrom = "From"
                Case "Incoming LT"
                    andSQL = "AND (Type = 'Lateral Transfer') " & _
                            "AND (FromState <> 'WV') " & _
                            "AND (VetStatus = 'DST Approved' OR VetStatus = 'DST Directed') "
                    toFrom = "To"
                Case "Turn-In"
                    andSQL = "AND (Type = 'Turn-In') " & _
                            "AND (VetStatus = 'DST Approved' OR VetStatus = 'DST Directed') "
                    toFrom = "From"
                Case "Open Vetting"
                    andSQL = "AND (Type = 'Lateral Transfer' OR Type = 'Turn-In') " & _
                            "AND (VetLevel NOT LIKE 'A*') " & _
                            "AND (VetStatus = 'Vetting Open') "
                    toFrom = "From"
            End Select
'full SQL statement is concatenated
            SetSQL = "SELECT * " & _
                        "FROM PSDs " & _
                        "INNER JOIN PBAuth " & _
                        "ON (PSDs." & toFrom & "UIC = PBAuth.PBUIC) AND (PSDs." & toFrom & "PBLIN = PBAuth.PBAuthLin) " & _
                        "WHERE (PSDs." & toFrom & "UIC in (" & newReport.strIN & ")) " & _
                        "AND (PSDStatus = 'Open' OR PSDStatus = 'Expired') " & andSQL & _
                        "ORDER BY PSDID"
'(5) of (7) STAFF REPORT worksheet names
'*****this function is not called for "Staff_Report" sheet name*****
        Case "In_State_LT", "Outgoing_LT", "Incoming_LT", "Open_Vet", "Turn_In"
'inner select assigns SQL AND statements to specific sheets
            Select Case newReport.sheetName
                Case "In_State_LT"
                    andSQL = "AND (Type = 'Lateral Transfer') " & _
                            "AND (FromState = 'WV' AND ToState = 'WV') " & _
                            "AND (FromUIC <> 'W7N7AA') " & _
                            "AND (VetStatus = 'DST Approved' OR VetStatus = 'DST Directed')"
                Case "Outgoing_LT"
                    andSQL = "AND (Type = 'Lateral Transfer') " & _
                            "AND (ToState <> 'WV')"
                Case "Incoming_LT"
                    andSQL = "AND (Type = 'Lateral Transfer') " & _
                            "AND (FromState <> 'WV') " & _
                            "AND (VetStatus = 'DST Approved' OR VetStatus = 'DST Directed')"
                Case "Open_Vet"
                    andSQL = "AND (Type = 'Lateral Transfer' OR Type = 'Turn-In') " & _
                            "AND (FromState = 'WV') " & _
                            "AND (VetStatus = 'Vetting Open')"
                Case "Turn_In"
                    andSQL = "AND (Type = 'Turn-In') " & _
                            "AND (VetStatus = 'DST Approved' OR VetStatus = 'DST Directed')"
            End Select
'full SQL statement is concatenated
            SetSQL = "SELECT * " & _
                        "FROM PSDs " & _
                        "WHERE (PSDStatus = 'Open' OR PSDStatus = 'Expired') " & andSQL
'(1) of (7) STAFF REPORT worksheet names
        Case "Await_PU"
'assign SQL statement
            SetSQL = "SELECT * " & _
                        "FROM PBAuth " & _
                        "WHERE (PBDueOut <> 0) " & _
                        "AND (PBUIC = 'W7N7AA')"
            
    End Select

End Function

Public Function BuildSpreadsheet() As Excel.Workbook

'On Error GoTo Err_General

Dim xlApp As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim rangeStr As String
    
    Set xlApp = Excel.Application
'uses object type property
    If newReport.objType = "staffReport" Then
'copies recordset to STAFF REPORT template file
            Set BuildSpreadsheet = xlApp.Workbooks.Open(newReport.fileName)
            Set xlSheet = BuildSpreadsheet.Worksheets(newReport.sheetName)
            Worksheets(newReport.sheetName).Activate
            newReport.rsReport.MoveFirst
                
                If newReport.sheetName = "Staff_Report" Then
'recordset is data from previous report
'copied into cells to be refernced for comparison by excel worksheet formulas
                    With xlSheet
                        .range("C1:C33").ClearContents
                        With newReport.rsReport
                            rangeStr = "T4!" & Nz(!Date, "") & "!" & _
                                "C1!" & Nz(!InStateTotal, "") & "!" & _
                                "C5!" & Nz(![InState$], "") & "!" & _
                                "C6!" & Nz(!OutgoingTotal, "") & "!" & _
                                "C10!" & Nz(![Outgoings$], "") & "!" & _
                                "C11!" & Nz(!TITotal, "") & "!" & _
                                "C15!" & Nz(![TI$], "") & "!" & _
                                "C16!" & Nz(!OpenVetTotal, "") & "!" & _
                                "C20!" & Nz(![OpenVet$], "") & "!" & _
                                "C21!" & Nz(![54Total], "") & "!" & _
                                "C22!" & Nz(![54$], "") & "!" & _
                                "C23!" & Nz(![54Vehicles], "") & "!" & _
                                "C24!" & Nz(!IncomingTotal, "") & "!" & _
                                "C28!" & Nz(![Incoming$], "")
                            MultiRangeColumn xlSheet, rangeStr, "values"
                        End With
                        .range("L1") = Format(Now, "mm/dd/yyyy HH:mm:ss")
                    End With
                Else
'clears data from referenced sheet name, copies in new data
'excel worksheet template contains absolute formulas for data processing
                        xlSheet.range("A3:AH2500").ClearContents
                        xlSheet.range("A3").CopyFromRecordset newReport.rsReport
                End If
'save, close and quit template
            BuildSpreadsheet.Save
            BuildSpreadsheet.Close True
            xlApp.quit
    
    Else
'creates new report based on COMBOBOX selection
            Set BuildSpreadsheet = xlApp.Workbooks.Add
            Set xlSheet = BuildSpreadsheet.Worksheets(1)
            
            With xlSheet
                .name = newReport.sheetName
                .Cells.Font.name = "Calibri"
                .Cells.Font.Size = 11
                
                rangeStr = "A:A,L:M!14! C!16! D:D,F:F!7! E:E,G:G,N:N!10! H:I,P:V!5! J!12"
                MultiRangeColumn xlSheet, rangeStr, "width"
                
'build column headers
                rangeStr = "A1!PSD ID! B1!From PB LIN! C1!NIIN Nomen! D1!From State! " & _
                    "E1!From UIC! F1!To State! G1!To UIC! H1!Val! I1!Req! " & _
                    "J1!Vetting Level! K1!Cond Code! L1!Unit Price! M1!Extended Price! " & _
                    "N1!Suspense! O1!Catalog LIN! P1!Auth! Q1!Auth +1! R1!Auth +2! " & _
                    "S1!+/-! T1!PB On Hand! U1!PB Due In! V1!PB Due Out! W1!DARPL! " & _
                    "X1!Pass Thru! Y1!Status! Z1!Type! AA1!Notes"
                MultiRangeColumn xlSheet, rangeStr, "values"
            Dim e As Integer            'recordset iterator
                e = 2                   'row 2 on worksheet
'copy recordset into worksheet
                Do While Not newReport.rsReport.EOF
                    With newReport.rsReport
                        rangeStr = "A" & e & "!" & Nz(!PSDID, "") & "!" & _
                            "B" & e & "!" & Nz(!FromPBLIN, "") & "!" & _
                            "C" & e & "!" & Nz(!NiinNomen, "") & "!" & _
                            "D" & e & "!" & Nz(!FromState, "") & "!" & _
                            "E" & e & "!" & Nz(!FromUIC, "") & "!" & _
                            "F" & e & "!" & Nz(!ToState, 0) & "!" & _
                            "G" & e & "!" & Nz(!ToUIC, 0) & "!" & _
                            "H" & e & "!" & Nz(!Validated, 0) & "!" & _
                            "I" & e & "!" & Nz(!Requested, "") & "!" & _
                            "J" & e & "!" & Nz(!VetLevel, "") & "!" & _
                            "K" & e & "!" & Nz(!CondCode, "") & "!" & _
                            "L" & e & "!" & Nz(![Unit$], "") & "!" & _
                            "M" & e & "!" & Nz(![Ext$], "") & "!" & _
                            "N" & e & "!" & Nz(!Suspense, "") & "!" & _
                            "O" & e & "!" & Nz(!CatLIN, "") & "!" & _
                            "P" & e & "!" & Nz(!CurAuth, "") & "!" & _
                            "Q" & e & "!" & Nz(!Auth1, "") & "!" & _
                            "R" & e & "!" & Nz(!Auth2, "") & "!" & _
                            "S" & e & "!" & Nz(!PlusMinus, "") & "!" & _
                            "T" & e & "!" & Nz(!PBOnHand, "") & "!" & _
                            "U" & e & "!" & Nz(!PBDueIn, "") & "!" & _
                            "V" & e & "!" & Nz(!PBDueOut, "") & "!" & _
                            "W" & e & "!" & Nz(!Darpl, "") & "!" & _
                            "X" & e & "!" & Nz(!IsPassThru, "") & "!" & _
                            "Y" & e & "!" & Nz(!PSDStatus, "") & "!Z" & e & "!" & Nz(!Type, "")
                        MultiRangeColumn xlSheet, rangeStr, "values"
                    End With
                        
                        If (newReport.sheetName = "Open Vetting") Then
'additional record for specific sheet name
                            .range("AB1").Value = "Vet Lvl Assign Date"
                            .range("AB" & e).Value = Nz(newReport.rsReport![VetLevelAssignDate], "")
                        End If
'apply conditional formatting by row values
                    'orange
                    If (.range("Y" & e) = "Expired") Then Rows(e).Interior.Color = RGB(255, 230, 153)
                    'green
                    If (.range("S" & e) > .range("I" & e)) Then .range("P" & e & ":V" & e).Interior.Color = RGB(109, 255, 109)
                    'red
                    If (.range("S" & e) < .range("I" & e)) Then .range("P" & e & ":V" & e).Interior.Color = RGB(255, 109, 109)
                    'grey
                    If (.range("S" & e) = "0") Then .range("P" & e & ":V" & e).Interior.Color = RGB(221, 221, 221)
'iterate record, row
                    e = e + 1
                    newReport.rsReport.MoveNext
                Loop
'format worksheet
                .range("A1:Z1").AutoFilter
                .range("A1:AB1").Interior.Color = RGB(224, 224, 224)
                .range("A1:AA1").Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                .range("A1:AA1").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                .range("A1:AA1").Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                .range("A1:AA1").HorizontalAlignment = xlLeft
                .range("A1:AA1").WrapText = True
                .Rows(1).RowHeight = 30
                
            ActiveWindow.FreezePanes = False
            Rows("2:2").Select
            ActiveWindow.FreezePanes = True
'delete bottom row where duplicate records have been copied
            Dim PSD As Long
            Dim lastRow As Long
            lastRow = Cells(Rows.Count, "A").End(xlUp).Row
'traverse thru rows TOP_DOWN
                For PSD = 2 To lastRow
                     If (.range("A" & PSD) = .range("A" & (PSD + 1))) Then Rows(PSD + 1).Delete
                Next PSD
            End With
    End If
    
Exit_General:
'clear objects
    Set xlApp = Nothing
    Set xlSheet = Nothing
    Exit Function
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General
End Function


