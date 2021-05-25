Attribute VB_Name = "LIN_Search"
Option Compare Database
Option Explicit

Public linSearch As Form_frmLINSearch
Public curForm As Form

Dim l As Boolean
Dim tdf As DAO.TableDef
Dim db As DAO.Database

Public Sub LINInitialize(openForm As Form)
    
    l = True
    Set linSearch = New Form_frmLINSearch
    Set curForm = openForm
    Set db = CurrentDb
    
        linSearch.optGroup = curForm.optLinNiin
        linSearch.optYr = curForm.optYrCompare
        l = FormLINs(l)
            If l = False Then Exit Sub
        BuildLINTable
        
        Set curForm = Nothing
        
End Sub

Public Sub LINFinish(openForm As Form)
    
    Set curForm = openForm
    
    If CurrentProject.AllForms("frmRemoveLIN").IsLoaded Then
        DeleteSelected
    End If
    
    ConsolidatePBNIINTables
    Set linSearch.rsResult = RecordsetMatch(0)
        If linSearch.rsResult.recordCount = 0 Then
            MsgBox "No data selected", vbInformation + vbOKOnly, "No data exported"
            Exit Sub
        End If
    BuildSpreadsheet
    CloseRecordsets

End Sub

Public Function FormLINs(l As Boolean) As Boolean
 
'verify TEXT input
    If IsNull(curForm.lookUpLIN) And IsNull(curForm.lookUpLIN2) Then
        MsgBox "ONE LIN or NIIN is Required to Run Report"
        l = False
        Exit Function
    End If
    
    If Not IsNull(curForm.lookUpLIN) Then
'text box (1), validate TEXT input is correct LIN or NIIN format
        ValidateLIN curForm.lookUpLIN, l
        If l = False Then
            MsgBox "Invalid Number - Field 1"
            Exit Function
        Else
'assign validated number to variable
            linSearch.lin1 = curForm.lookUpLIN
        End If
    End If

    If Not IsNull(curForm.lookUpLIN2) Then
'text box (2), validate TEXT input is correct LIN or NIIN format
        ValidateLIN curForm.lookUpLIN2, l
        If l = False Then
            MsgBox "Invalid Number - Field 2"
            Exit Function
        Else
'assign validated number to variable
            linSearch.lin2 = curForm.lookUpLIN2
        End If
    End If

    FormLINs = l

End Function

Public Function ValidateLIN(frmLIN As String, l As Boolean) As Boolean
'uses OPTION GROUP to determine which regular expression to use
'validates text input matches requirement
'returns FALSE if string doesn't match any of the 3 options
Debug.Print frmLIN
l = False
Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    With regExObj
        .IgnoreCase = True
'if LIN OPTION GROUP is selected
Debug.Print linSearch.optGroup
        If linSearch.optGroup = 1 Then
'LIN = (6) aplha-numeric char string
            .pattern = "[a-zA-Z0-9]{6}"
            l = .test(frmLIN)
            ValidateLIN = l
        Else
'if NIIN OPTION GROUP is selected
            If Len(frmLIN) <= 11 Then
'NIINa = (9) numeric char string
                .pattern = "[0-9]{9}"
                l = .test(frmLIN)
                ValidateLIN = l
            Else
'NSN = (13) alpha-numeric char string
                .pattern = "[0-9]{6}[a-zA-Z]{1}[0-9]{6}"
                l = .test(frmLIN)
                ValidateLIN = l
            End If
        End If
    End With

End Function

Public Sub BuildLINTable()
'create temp table to insert LIN and NIIN values
'uses OPTION GROUP to begin inserting records into table
'multiple sql statements to search for all associated LINs

'delete temp table if it exists
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempLINGroup'")) Then
        DoCmd.DeleteObject acTable, "tempLINGroup"
    End If
'create table and field names
Set tdf = db.CreateTableDef("tempLINGroup")
    tdf.Fields.Append tdf.CreateField("LIN", dbText)
    tdf.Fields.Append tdf.CreateField("NIIN", dbText)
'build new temp table
    db.TableDefs.Append tdf
'delete temp table if it exists
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempGroupBy'")) Then
        DoCmd.DeleteObject acTable, "tempGroupBy"
    End If
'create table and field names
Set tdf = db.CreateTableDef("tempGroupBy")
    tdf.Fields.Append tdf.CreateField("LIN", dbText)
    tdf.Fields.Append tdf.CreateField("NIIN", dbText)
    tdf.Fields.Append tdf.CreateField("Nomen", dbText)
'build new temp table
    db.TableDefs.Append tdf
'delete temp table if it exists
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempGroupByDistinctLIN'")) Then
        DoCmd.DeleteObject acTable, "tempGroupByDistinctLIN"
    End If

    If linSearch.optGroup = 1 Then
'insert LIN records from text box on form
        DoCmd.RunSQL "INSERT INTO tempLINGroup (LIN) VALUES ('" & linSearch.lin1 & "');"
        DoCmd.RunSQL "INSERT INTO tempLINGroup (LIN) VALUES ('" & linSearch.lin2 & "');"
    Else
'insert NIIN records from text box on form
        DoCmd.RunSQL "INSERT INTO tempLINGroup (NIIN) VALUES ('" & linSearch.lin1 & "');"
        DoCmd.RunSQL "INSERT INTO tempLINGroup (NIIN) VALUES ('" & linSearch.lin2 & "');"
'run INSERT INTO SQL to fill LIN fields that match NIIN fields
'all further searches are performed with LINs
        DoCmd.RunSQL "INSERT INTO tempLINGroup (LIN) " & _
            "SELECT NIINDetail.NiinCatLIN " & _
            "FROM tempLINGroup " & _
            "INNER JOIN NIINDetail " & _
            "ON tempLINGroup.NIIN = NIINDetail.NIIN;"
    End If
'run SQL to get all authorized substitutes from SB700 table
        DoCmd.RunSQL "INSERT INTO tempLINGroup (LIN) " & _
            "SELECT SB700.AppH1SubLIN " & _
            "FROM SB700 " & _
            "INNER JOIN tempLINGroup ON SB700.AppH1AuthLIN = tempLINGroup.LIN;"
'run SQL to get primary LINs with tempLINGroup as substitutes
        DoCmd.RunSQL "INSERT INTO tempLINGroup (LIN) " & _
            "SELECT SB700.AppH1AuthLIN " & _
            "FROM SB700 " & _
            "INNER JOIN tempLINGroup ON SB700.AppH1SubLIN = tempLINGroup.LIN;"
'run SQL to get all LINs being used as subs where tempLINGroup is PB LIN
        DoCmd.RunSQL "INSERT INTO tempLINGroup (LIN) " & _
            "SELECT NIINDetail.NiinCatLIN " & _
            "FROM tempLINGroup " & _
            "INNER JOIN NIINDetail ON tempLINGroup.LIN = NIINDetail.NiinPBLIN;"
'run SQL to join all associated LINs with fields from NIINDetail table
'allows for multiple NIIN associated with a single LIN
'becomes ROWSOURCE for lstPickLINGroup
'*****a valid LIN that doesn't have an associated piece of equipment*****
'*****on any PB's here in the state will be deleted by the***************
'*****INNER JOIN statement in this SQL***********************************
        DoCmd.RunSQL "INSERT INTO tempGroupBy (LIN, NIIN, Nomen) " & _
            "SELECT tempLINGroup.LIN, NIINDetail.NIIN, NIINDetail.NiinCatLINNomen " & _
            "FROM NIINDetail " & _
            "INNER JOIN tempLINGroup ON NIINDetail.NiinCatLIN = tempLINGroup.LIN " & _
            "GROUP BY tempLINGroup.LIN, NIINDetail.NIIN, NIINDetail.NiinCatLINNomen;"
'if no results returned, run same SQL but matching LINs to PBAuth table
        If DCount("*", "tempGroupBy") = 0 Then
            DoCmd.RunSQL "INSERT INTO tempGroupBy (LIN, NIIN, Nomen) " & _
                    "SELECT tempLINGroup.LIN, PBAuth.NIIN, PBAuth.PBAuthNomen " & _
                    "FROM PBAuth " & _
                    "INNER JOIN tempLINGroup ON PBAuth.PBAuthLIN = tempLINGroup.LIN " & _
                    "GROUP BY tempLINGroup.LIN, PBAuth.NIIN, PBAuth.PBAuthNomen;"
        End If
'run SQL to create table of distinct records
'becomes ROWSOURCE for lstPickLINGroupBy
        DoCmd.RunSQL "SELECT tempGroupBy.LIN, " & _
            "First(tempGroupBy.NIIN) AS NIIN, First(tempGroupBy.Nomen) AS Nomen " & _
            "INTO tempGroupByDistinctLIN " & _
            "FROM tempGroupBy " & _
            "GROUP BY tempGroupBy.LIN;"
            
Set tdf = Nothing
    
End Sub

Public Sub DeleteSelected()

Dim strIN As String         'concatenates LISTBOX selections
Dim i As Integer            'LISTBOX iterater
Dim lstOne As Access.ListBox
Dim lstTwo As Access.ListBox
    Set lstOne = curForm.lstPickLINGroupBy
    Set lstTwo = curForm.lstPickLINGroup
        
'build string by looping through the unit LIN LISTBOX
        For i = 0 To lstOne.ListCount - 1
            If lstOne.Selected(i) Then
                strIN = strIN & "'" & lstOne.Column(0, i) & "',"
            End If
        Next i
            If Len(strIN) <> 0 Then
'if LIN selections were made, assign to object string
'***Left(strIN, Len(strIN) - 1) is used to remove comma at end of string
                linSearch.deleteLRecords = Left(strIN, Len(strIN) - 1)
            End If
    strIN = ""
'build string by looping through the unit NIIN LISTBOX
        For i = 0 To lstTwo.ListCount - 1
'adds items not selected
            If Not lstTwo.Selected(i) Then
                strIN = strIN & "'" & lstTwo.Column(1, i) & "',"
            End If
        Next i
            If Len(strIN) <> 0 Then
'if NIIN selections were made, assign to object string
                linSearch.deleteNRecords = Left(strIN, Len(strIN) - 1)
            End If
'clear LIST BOX values
    lstTwo.RowSourceType = "Value List"
    lstTwo.RowSource = ""
    lstOne.RowSourceType = "Value List"
    lstOne.RowSource = ""

End Sub

Public Sub ConsolidatePBNIINTables()
'create deletion strings
'creates temp tables to match LINs to PBAuth and NIINDetail tables
'creates temp table to JOIN PBAuth, NIINDetail and Units table info
'deletes any LIN and/or NIIN selections made on LISTBOXES
Dim strLDelete As String
Dim strNDelete As String
Dim strLNDelete As String
'to delete LINs in PBAuthLin field
    strLDelete = "DELETE * " & _
                "FROM tempLINConsolidated " & _
                "WHERE tempLINConsolidated.PBAuthLin in (" & linSearch.deleteLRecords & ");"
'to delete NIINs NOT in IIN field
    strNDelete = "DELETE * " & _
                "FROM tempLINConsolidated " & _
                "WHERE tempLINConsolidated.NIIN NOT in (" & linSearch.deleteNRecords & ");"
'to delete LINs in NiinCatLIN field
    strLNDelete = "DELETE * " & _
                "FROM tempLINConsolidated " & _
                "WHERE tempLINConsolidated.NiinCatLIN in (" & linSearch.deleteLRecords & ");"
'delete temp table if exists
        If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempMatchPBAuth'")) Then
            DoCmd.DeleteObject acTable, "tempMatchPBAuth"
        End If
'run SQL SELECT INTO temp table PBAuth fields WHERE LIN = PBAuth
    db.Execute "SELECT DISTINCT PBUIC, PBAuthLin, PBAuthNomen, " & _
            "CurAuth, Auth1, Auth2, PercentFill, " & _
            "PlusMinus, Darpl, PBOnHand, PBDueIn, " & _
            "PBDueOut, IsCTA, Left([PBUIC],4) AS [Expr1], UIC_LIN " & _
            "INTO tempMatchPBAuth " & _
            "FROM tempGroupBY " & _
            "INNER JOIN PBAuth ON tempGroupBY.LIN = PBAuth.PBAuthLin;"
'delete temp table if exists
        If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempMatchNIIN'")) Then
            DoCmd.DeleteObject acTable, "tempMatchNIIN"
        End If
'run SQL SELECT INTO temp table NIINDetail fields WHERE LIN = NIINDetail
    db.Execute "SELECT DISTINCT NiinUIC, NiinPBLIN, NiinDueIn, NiinDueOut, " & _
            "NiinCatLIN, NIINDetail.NIIN, NiinNomen, NiinOH, Left([NiinUIC],4) AS [Expr1], UIC_NIIN " & _
            "INTO tempMatchNIIN " & _
            "FROM tempGroupBY " & _
            "INNER JOIN NIINDetail ON tempGroupBY.LIN = NIINDetail.NiinCatLIN;"
'delete temp table if exists
        If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempLINConsolidated'")) Then
            DoCmd.DeleteObject acTable, "tempLINConsolidated"
        End If
'run SQL SELECT INTO temp table JOIN PBAuthLin = NiinPBLIN & UIC = UIC & Units = UIC
    db.Execute "SELECT tempMatchPBAuth.PBUIC, tempMatchPBAuth.PBAuthLin, " & _
            "tempMatchPBAuth.PBAuthNomen, tempMatchNIIN.NiinPBLIN, tempMatchNIIN.NiinDueIn, " & _
            "tempMatchNIIN.NiinDueOut, tempMatchNIIN.NiinCatLIN, tempMatchNIIN.NIIN, " & _
            "tempMatchNIIN.NiinNomen, Left([PBUIC],4) AS Expr1, tempMatchNIIN.NiinOH, " & _
            "Units.Unit, tempMatchPBAuth.CurAuth, tempMatchPBAuth.Auth1, tempMatchPBAuth.Auth2, " & _
            "tempMatchPBAuth.PlusMinus, tempMatchPBAuth.PBOnHand, tempMatchPBAuth.PBDueIn, " & _
            "tempMatchPBAuth.PBDueOut, tempMatchPBAuth.UIC_LIN, tempMatchNIIN.UIC_NIIN " & _
            "INTO tempLINConsolidated " & _
            "FROM Units " & _
            "INNER JOIN (tempMatchPBAuth " & _
            "LEFT JOIN tempMatchNIIN " & _
            "ON (tempMatchPBAuth.PBAuthLin = tempMatchNIIN.NiinPBLIN) " & _
            "AND (tempMatchPBAuth.PBUIC = tempMatchNIIN.NiinUIC)) " & _
            "ON Units.UIC = tempMatchPBAuth.PBUIC;"
'if deletion strings in object have been filled with LISTBOX selections
        If Len(linSearch.deleteLRecords) > 0 Then
'run DELETE * SQL WHERE LIN = PBAuth
            DoCmd.RunSQL strLDelete
        End If

        If Len(linSearch.deleteNRecords) > 0 Then
'run DELETE * SQL WHERE NIIN NOT = NIINDetail
            DoCmd.RunSQL strNDelete
            If Len(linSearch.deleteLRecords) > 0 Then
'run DELETE * SQL WHERE LIN = NIINDetail
                DoCmd.RunSQL strLNDelete
            End If
        End If

End Sub

Public Function RecordsetMatch(m As Integer) As DAO.Recordset
Dim sql As String

Select Case m
    Case 0 'LIN SQL
        sql = "SELECT Expr1, PBUIC, Unit, PBAuthLin, " & _
                "PBAuthNomen, CurAuth, Auth1, Auth2, PlusMinus, " & _
                "PBOnHand, PBDueIn, PBDueOut, " & _
                "NiinCatLIN, NiinOH, NiinDueIn, NiinDueOut, " & _
                "NIIN, NiinNomen, UIC_LIN, UIC_NIIN " & _
                "FROM tempLINConsolidated " & _
                "ORDER BY PBUIC"
        Set RecordsetMatch = db.OpenRecordset(sql, dbOpenDynaset)
    Case 1 'MATCAT
        sql = "SELECT DISTINCT PBUIC, Unit, NiinPBLIN, NIIN, NiinNomen, NiinOH, " & _
                "Code_Position, MATCAT_Pos, Found_Code, Desc, " & _
                "tempLINConsolidated.UIC_NIIN " & _
                "INTO tempMatcat " & _
                "FROM tempLINConsolidated " & _
                "INNER JOIN tblMatchedMATCAT " & _
                "ON tempLINConsolidated.UIC_NIIN = tblMatchedMATCAT.UIC_NIIN " & _
                "ORDER BY Unit, NiinPBLIN, Code_Position;"
        DoCmd.RunSQL sql
        Set RecordsetMatch = CurrentDb.OpenRecordset("tempMatcat", dbOpenDynaset)
    Case 2 'PBIC
        sql = "SELECT DISTINCT PBUIC, Unit, PBAuthLin, PBAuthNomen, CurAuth, " & _
                "MatchedPTE, Matched2ndPTE, Desc, Desc_2, tempLINConsolidated.UIC_LIN, " & _
                "[Multi-PBIC], [Multi-TAC] " & _
                "INTO tempPbic " & _
                "FROM tempLINConsolidated " & _
                "INNER JOIN tblMatchedPBICTACERC " & _
                "ON tempLINConsolidated.UIC_LIN = tblMatchedPBICTACERC.UIC_LIN " & _
                "ORDER BY Unit, PBAuthLin;"
        DoCmd.RunSQL sql
        Set RecordsetMatch = CurrentDb.OpenRecordset("tempPbic", dbOpenDynaset)
    Case 3 'FromPSDs
        sql = "SELECT DISTINCT PBUIC, Unit, NiinCatLIN, CatLIN, NIIN, SourceNiin, " & _
                "tempLINConsolidated.NiinNomen, NiinDueIn, NiinDueOut, FromUIC, ToUIC, " & _
                "PSDID, Type, PSDStatus, ApprovedOn, Requested, PBAuthLin " & _
                "INTO tempFromPsds " & _
                "FROM tempLINConsolidated " & _
                "INNER JOIN PSDs " & _
                "ON (PSDs.FromPBLIN = tempLINConsolidated.PBAuthLIN) " & _
                "AND (tempLINConsolidated.NIIN = PSDs.SourceNiin) " & _
                "AND (tempLINConsolidated.PBUIC = PSDs.FromUIC) " & _
                "WHERE (tempLINConsolidated.NiinDueIn > 0) " & _
                "OR (tempLINConsolidated.NiinDueOut > 0);"
        DoCmd.RunSQL sql
        Set RecordsetMatch = CurrentDb.OpenRecordset("tempFromPsds", dbOpenDynaset)
    Case 4 'ToPSDs
        sql = "SELECT DISTINCT PBUIC, Unit, NiinCatLIN, CatLIN, NIIN, SourceNiin, " & _
                "tempLINConsolidated.NiinNomen, NiinDueIn, NiinDueOut, FromUIC, ToUIC, " & _
                "PSDID, Type, PSDStatus, ApprovedOn, Requested, PBAuthLin " & _
                "INTO tempToPsds " & _
                "FROM tempLINConsolidated " & _
                "INNER JOIN PSDs " & _
                "ON (PSDs.FromPBLIN = tempLINConsolidated.PBAuthLIN) " & _
                "AND (tempLINConsolidated.NIIN = PSDs.SourceNiin) " & _
                "AND (tempLINConsolidated.PBUIC = PSDs.ToUIC) " & _
                "WHERE (tempLINConsolidated.NiinDueIn > 0) " & _
                "OR (tempLINConsolidated.NiinDueOut > 0);"
        DoCmd.RunSQL sql
        Set RecordsetMatch = CurrentDb.OpenRecordset("tempToPsds", dbOpenDynaset)
    Case 5 'Primary ASIOE
        sql = "SELECT DISTINCT ASIOE_CMI.*, tempLINConsolidated.PBUIC " & _
                "INTO tempPrmAsioe " & _
                "FROM (tempLINConsolidated " & _
                "INNER JOIN ASIOE_CMI ON tempLINConsolidated.PBAuthLin = ASIOE_CMI.Prm_LIN) " & _
                "INNER JOIN PBAuth ON (PBAuth.PBAuthLin = ASIOE_CMI.Prm_LIN) " & _
                "AND (tempLINConsolidated.PBUIC = PBAuth.PBUIC);"
        DoCmd.RunSQL sql
        Set RecordsetMatch = CurrentDb.OpenRecordset("tempPrmAsioe", dbOpenDynaset)
    Case 6 'Associated ASIOE
        sql = "SELECT DISTINCT tempLINConsolidated.PBUIC, ASIOE_CMI.* " & _
                "INTO tempAscAsioe " & _
                "FROM (tempLINConsolidated " & _
                "INNER JOIN PBAuth ON tempLINConsolidated.PBUIC = PBAuth.PBUIC) " & _
                "INNER JOIN ASIOE_CMI ON (PBAuth.PBAuthLin = ASIOE_CMI.Prm_LIN) " & _
                "AND (tempLINConsolidated.NiinCatLIN = ASIOE_CMI.Asc_LIN);"
        DoCmd.RunSQL sql
        Set RecordsetMatch = CurrentDb.OpenRecordset("tempAscAsioe", dbOpenDynaset)
End Select
End Function

Public Sub BuildSpreadsheet()
'creates excel spreadsheet report
'copies multiple recordsets into spreadsheet with layered loops to compare records
'
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet              'worksheet PB_Auth_&_NIIN_Detail
    Dim xlSheet2 As Excel.Worksheet             'worksheet MATCAT_Info
    Dim xlSheet3 As Excel.Worksheet             'worksheet PBIC-TAC-ERC_Info
    Dim xlSheet4 As Excel.Worksheet             'worksheet PSD_Info
    Dim xlSheet5 As Excel.Worksheet
    Dim rangeStr As String
    Dim e As Long                               'rsResult iterator
    Dim x As Long                               'rsMATCAT iterator
    Dim k As Long                               'rsPBIC iterator
    Dim l As Long                               'rsFromPSDs & rsToPSDs iterator
    Dim m As Long                               'rsPrmAsioe & rsAscAsioe iterator
        
        Set xlApp = Excel.Application
        xlApp.Visible = False
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
        Set xlSheet2 = xlBook.Worksheets(2)
        Set xlSheet3 = xlBook.Worksheets(3)
        Set xlSheet4 = xlBook.Sheets.Add(After:=xlSheet)
        Set xlSheet5 = xlBook.Sheets.Add(After:=xlSheet3)
        With linSearch
            Set .rsMATCAT = RecordsetMatch(1)
            Set .rsPBIC = RecordsetMatch(2)
            Set .rsFromPSDs = RecordsetMatch(3)
            Set .rsToPSDs = RecordsetMatch(4)
            Set .rsPrmAsioe = RecordsetMatch(5)
            Set .rsAscAsioe = RecordsetMatch(6)
        End With
        
            xlSheet.name = "PB_Auth_&_NIIN_Detail"
            xlSheet2.name = "MATCAT_Info"
            xlSheet3.name = "PBIC-TAC-ERC_Info"
            xlSheet4.name = "PSD_Info"
            xlSheet5.name = "ASIOE_CMI"
            
            With xlSheet
                .Cells.Font.name = "Calibri"
                .Cells.Font.Size = 11
                
                rangeStr = "A!6! D!14! E!12! F!8! G:L!7! T:T,P:P!10! O!9! U!25! V!40"
                MultiRangeColumn xlSheet, rangeStr, "width"
                
                rangeStr = "F1!L1! A1!C1! A2!B2! C2!U2"
                MultiRangeColumn xlSheet, rangeStr, "merge"
                
                rangeStr = "F3!=SUM(F5:F150)! G3!=SUM(G5:G150)! H3!=SUM(H5:H150)! " & _
                    "I3!=SUM(I5:I150)! J3!=SUM(J5:J150)! K3!=SUM(K5:K150)! L3!=SUM(L5:L150)"
                MultiRangeColumn xlSheet, rangeStr, "formulas"
                
                .Columns("T").NumberFormat = "@" 'text format
'build column headers

                rangeStr = "A4!BN! B4!UIC! C4!PB Auth LIN! D4!Unit! E4!PB Nomen! F4!Cur Auth! " & _
                            "G4!Auth +1! H4!Auth +2! I4!+/-! J4!PB On Hand! K4!PB Due In! L4!PB Due Out! " & _
                            "M4!SET Incoming! M4!SET Outgoing! O4!Projected On Hand! P4!NIIN Catalog LIN! " & _
                            "Q4!NIIN On Hand! R4!NIIN Due In! S4!NIIN Due Out! T4!NIIN! U4!NIIN Nomen! " & _
                            "V4!Notes! A1!Prepared On:! D1!" & Date & "! C3!Links PBIC! T3!Links MATCAT! " & _
                            "B3!Links ASIOE! A2!Reason/Notes"
                MultiRangeColumn xlSheet, rangeStr, "values"
        
                e = 5 'rsResult
                x = 5 'rsMATCAT
                k = 5 'rsPBIC
                l = 5 'rsFromPSDs & rsToPSDs
                m = 5 'rsPrmAsioe & rsAscAsioe
'loop through all records in rsResult
                Do While Not linSearch.rsResult.EOF
                xlSheet.Activate
                    With linSearch.rsResult
                        rangeStr = "A" & e & "!" & !Expr1 & "!" & _
                                    "B" & e & "!" & !PBUIC & "!" & _
                                    "C" & e & "!" & !PBAuthLin & "!" & _
                                    "D" & e & "!" & !Unit & "!" & _
                                    "E" & e & "!" & !PBAuthNomen & "!" & _
                                    "F" & e & "!" & !CurAuth & "!" & _
                                    "G" & e & "!" & !Auth1 & "!" & _
                                    "H" & e & "!" & !Auth2 & "!" & _
                                    "I" & e & "!" & !PlusMinus & "!" & _
                                    "J" & e & "!" & !PBOnHand & "!" & _
                                    "K" & e & "!" & !PBDueIn & "!" & _
                                    "L" & e & "!" & !PBDueOut & "!" & _
                                    "P" & e & "!" & !NIINCatLIN & "!" & _
                                    "Q" & e & "!" & !NiinOH & "!" & _
                                    "R" & e & "!" & !NiinDueIn & "!" & _
                                    "S" & e & "!" & !NiinDueOut & "!" & _
                                    "T" & e & "!" & !NIIN & "!" & _
                                    "U" & e & "!" & !NiinNomen
                    End With
                MultiRangeColumn xlSheet, rangeStr, "values"

'**with CURRENT record*********************************
'**MATCAT**********************************************
                    If linSearch.rsMATCAT.recordCount <> 0 Then
                            linSearch.rsMATCAT.MoveFirst
'loop through rsMATCAT
                            Do While Not linSearch.rsMATCAT.EOF
'compare MATCAT record to CURRENT record
                                If linSearch.rsMATCAT!UIC_NIIN = linSearch.rsResult!UIC_NIIN Then
                        'p_rsMATCAT.FindFirst "UIC_NIIN = '" & p_rsResult!UIC_NIIN & "'"
                            'If Not (p_rsMATCAT.NoMatch) Then
                                    With xlSheet2
                                    xlSheet2.Activate
                                        With linSearch.rsMATCAT
                                            rangeStr = "A" & x & "!" & !PBUIC & "!" & _
                                                        "B" & x & "!" & !Unit & "!" & _
                                                        "C" & x & "!" & !NiinPBLIN & "!" & _
                                                        "D" & x & "!" & !NIIN & "!" & _
                                                        "E" & x & "!" & !NiinNomen & "!" & _
                                                        "F" & x & "!" & !NiinOH & "!" & _
                                                        "G" & x & "!" & !Code_Position & "!" & _
                                                        "H" & x & "!" & !Found_Code & "!" & _
                                                        "I" & x & "!" & !Desc & "!" & _
                                                        "J" & x & "!" & !UIC_NIIN
                                        End With
                                    MultiRangeColumn xlSheet2, rangeStr, "values"

                                        .Hyperlinks.Add Cells(x, 4), Address:="", SubAddress:="'" & xlSheet.name & "'!T" & e
                                    End With
                                        If linSearch.rsMATCAT!Code_Position = 1 Then
                                            xlSheet.Activate
                                            xlSheet.Hyperlinks.Add Cells(e, 20), Address:="", SubAddress:="'" & xlSheet2.name & "'!D" & x
                                        End If
'delete MATCAT record after it has been matched
                                    linSearch.rsMATCAT.Delete
'iterate x for every matched record
                                    x = x + 1
                                End If
                                linSearch.rsMATCAT.MoveNext
                            Loop
                    End If
'**PBIC-TAC-ERC**************************************
                    If linSearch.rsPBIC.recordCount <> 0 Then
                            linSearch.rsPBIC.MoveFirst
'loop through rsPBIC
                            Do While Not linSearch.rsPBIC.EOF
'compare PBIC record to CURRENT record
                                If linSearch.rsPBIC!UIC_LIN = linSearch.rsResult!UIC_LIN Then
                                
                                    With xlSheet3
                                    xlSheet3.Activate
                                        .Columns("F:I").NumberFormat = "@"
                                        With linSearch.rsPBIC
                                            rangeStr = "A" & k & "!" & !PBUIC & "!" & _
                                                        "B" & k & "!" & !Unit & "!" & _
                                                        "C" & k & "!" & !PBAuthLin & "!" & _
                                                        "D" & k & "!" & !PBAuthNomen & "!" & _
                                                        "E" & k & "!" & !CurAuth & "!" & _
                                                        "F" & k & "!" & !MatchedPTE & "!" & _
                                                        "G" & k & "!" & !Matched2ndPTE & "!" & _
                                                        "H" & k & "!" & ![Multi-PBIC] & "!" & _
                                                        "I" & k & "!" & ![Multi-TAC] & "!" & _
                                                        "J" & k & "!" & !Desc & "!" & _
                                                        "K" & k & "!" & !Desc_2 & "!" & _
                                                        "L" & k & "!" & !UIC_LIN
                                        End With
                                    MultiRangeColumn xlSheet3, rangeStr, "values"
                                        
                                        .Hyperlinks.Add Cells(k, 3), Address:="", SubAddress:="'" & xlSheet.name & "'!C" & e
                                    End With
                                        xlSheet.Activate
                                        xlSheet.Hyperlinks.Add Cells(e, 3), Address:="", SubAddress:="'" & xlSheet3.name & "'!C" & k
'delete PBIC record after it has been matched
                                    linSearch.rsPBIC.Delete
'iterate k for every matched record
                                    k = k + 1
                                End If
                                linSearch.rsPBIC.MoveNext
                            Loop
                    End If
'**From PSDs*********************************************
                    If linSearch.rsFromPSDs.recordCount <> 0 Then
                            linSearch.rsFromPSDs.MoveFirst
'loop through rsFromPSDs
                            Do While Not linSearch.rsFromPSDs.EOF
'compare FromPSD record to CURRENT record
                                If linSearch.rsFromPSDs!PBUIC = linSearch.rsResult!PBUIC _
                                    And linSearch.rsFromPSDs!NIIN = linSearch.rsResult!NIIN Then
                                    
                                    With xlSheet4
                                    xlSheet4.Activate
                                        '.Columns("F:I").NumberFormat = "@"
                                        With linSearch.rsFromPSDs
                                            rangeStr = "A" & l & "!" & !PBUIC & "!" & _
                                                        "B" & l & "!" & !Unit & "!" & _
                                                        "C" & l & "!" & !NIINCatLIN & "!" & _
                                                        "D" & l & "!" & !NIIN & "!" & _
                                                        "E" & l & "!" & !NiinNomen & "!" & _
                                                        "G" & l & "!" & !Requested & "!" & _
                                                        "H" & l & "!" & !FromUIC & "!" & _
                                                        "I" & l & "!" & !ToUIC & "!" & _
                                                        "J" & l & "!" & !PSDID & "!" & _
                                                        "K" & l & "!" & !PSDStatus & "!" & _
                                                        "L" & l & "!" & !Type & "!" & _
                                                        "M" & l & "!" & !ApprovedOn
                                        End With
                                    MultiRangeColumn xlSheet4, rangeStr, "values"

                                        .Hyperlinks.Add Cells(l, 7), Address:="", SubAddress:="'" & xlSheet.name & "'!L" & e
                                    End With
                                        xlSheet.Activate
                                        xlSheet.Hyperlinks.Add Cells(e, 12), Address:="", SubAddress:="'" & xlSheet4.name & "'!G" & l
'delete From PSD record after it has been matched
                                    linSearch.rsFromPSDs.Delete
'iterate l for every matched record
                                    l = l + 1
                                End If
                                linSearch.rsFromPSDs.MoveNext
                            Loop
                    End If
'**To PSDs*********************************************
                    If linSearch.rsToPSDs.recordCount <> 0 Then
                            linSearch.rsToPSDs.MoveFirst
'loop through rsToPSDs
                            Do While Not linSearch.rsToPSDs.EOF
'compare ToPSD record to CURRENT record
                                If linSearch.rsToPSDs!PBUIC = linSearch.rsResult!PBUIC _
                                    And linSearch.rsToPSDs!NIIN = linSearch.rsResult!NIIN Then
                                    
                                    With xlSheet4
                                    xlSheet4.Activate
                                        '.Columns("F:I").NumberFormat = "@"
                                        With linSearch.rsToPSDs
                                            rangeStr = "A" & l & "!" & !PBUIC & "!" & _
                                                        "B" & l & "!" & !Unit & "!" & _
                                                        "C" & l & "!" & !NIINCatLIN & "!" & _
                                                        "D" & l & "!" & !NIIN & "!" & _
                                                        "E" & l & "!" & !NiinNomen & "!" & _
                                                        "F" & l & "!" & !Requested & "!" & _
                                                        "H" & l & "!" & !FromUIC & "!" & _
                                                        "I" & l & "!" & !ToUIC & "!" & _
                                                        "J" & l & "!" & !PSDID & "!" & _
                                                        "K" & l & "!" & !PSDStatus & "!" & _
                                                        "L" & l & "!" & !Type & "!" & _
                                                        "M" & l & "!" & !ApprovedOn
                                        End With
                                    MultiRangeColumn xlSheet4, rangeStr, "values"

                                        .Hyperlinks.Add Cells(l, 6), Address:="", SubAddress:="'" & xlSheet.name & "'!K" & e
                                    End With
                                        xlSheet.Activate
                                        xlSheet.Hyperlinks.Add Cells(e, 11), Address:="", SubAddress:="'" & xlSheet4.name & "'!F" & l
'delete ToPSD record after it has been matched
                                    linSearch.rsToPSDs.Delete
'iterate l for every matched record
                                    l = l + 1
                                End If
                                linSearch.rsToPSDs.MoveNext
                            Loop
                    End If
                    
                    If linSearch.rsPrmAsioe.recordCount <> 0 Then
                        linSearch.rsPrmAsioe.MoveFirst
                        
                            Do While Not linSearch.rsPrmAsioe.EOF
                            
                                If linSearch.rsPrmAsioe!Prm_LIN = linSearch.rsResult!PBAuthLin And _
                                    linSearch.rsPrmAsioe!PBUIC = linSearch.rsResult!PBUIC Then
                                
                                    With xlSheet5
                                    xlSheet5.Activate
                                        With linSearch.rsPrmAsioe
                                            rangeStr = "A" & m & "!" & !Prm_LIN & "!" & _
                                                        "B" & m & "!" & !Prm_Nomen & "!" & _
                                                        "C" & m & "!" & !Prm_Qty & "!" & _
                                                        "D" & m & "!" & !Type & "!" & _
                                                        "E" & m & "!" & !Asc_LIN & "!" & _
                                                        "F" & m & "!" & !Asc_Nomen & "!" & _
                                                        "G" & m & "!" & !Asc_Qty & "!" & _
                                                        "K" & m & "!" & !PBUIC
                                        End With
                                    MultiRangeColumn xlSheet5, rangeStr, "values"
                                    
                                        .Hyperlinks.Add Cells(m, 11), Address:="", SubAddress:="'" & xlSheet.name & "'!B" & e
                                    End With
                                    xlSheet.Activate
                                    xlSheet.Hyperlinks.Add Cells(e, 2), Address:="", SubAddress:="'" & xlSheet5.name & "'!K" & m
                                    linSearch.rsPrmAsioe.Delete
                                    m = m + 1
                                End If
                                linSearch.rsPrmAsioe.MoveNext
                            Loop
                    End If
                    
                    If linSearch.rsAscAsioe.recordCount <> 0 Then
                        linSearch.rsAscAsioe.MoveFirst
                        
                            Do While Not linSearch.rsAscAsioe.EOF
                            
                                If linSearch.rsAscAsioe!Asc_LIN = linSearch.rsResult!NIINCatLIN And _
                                    linSearch.rsAscAsioe!PBUIC = linSearch.rsResult!PBUIC Then
                                
                                    With xlSheet5
                                    xlSheet5.Activate
                                        With linSearch.rsAscAsioe
                                            rangeStr = "H" & m & "!" & !Prm_LIN & "!" & _
                                                        "I" & m & "!" & !Prm_Nomen & "!" & _
                                                        "J" & m & "!" & !Prm_Qty & "!" & _
                                                        "D" & m & "!" & !Type & "!" & _
                                                        "E" & m & "!" & !Asc_LIN & "!" & _
                                                        "F" & m & "!" & !Asc_Nomen & "!" & _
                                                        "G" & m & "!" & !Asc_Qty & "!" & _
                                                        "K" & m & "!" & !PBUIC
                                        End With
                                        MultiRangeColumn xlSheet5, rangeStr, "values"

                                        .Hyperlinks.Add Cells(m, 11), Address:="", SubAddress:="'" & xlSheet.name & "'!B" & e
                                    End With
                                    xlSheet.Activate
                                    If Not (range("B" & e).Hyperlinks.Count > 0) Then
                                        xlSheet.Hyperlinks.Add Cells(e, 2), Address:="", SubAddress:="'" & xlSheet5.name & "'!K" & m
                                    End If
                                    linSearch.rsAscAsioe.Delete
                                    m = m + 1
                                End If
                                linSearch.rsAscAsioe.MoveNext
                            Loop
                    End If
                e = e + 1
                linSearch.rsResult.MoveNext
                Loop
'*****PB_Auth_&_NIIN_Detail*************************************************
            xlSheet.Activate
                Dim f As Long                   'traverses up rows
                Dim g As Long                   'traverses down rows
                Dim h As Long                   'counts f-iteration loops
                Dim lastRow As Long
                Dim yrSql As String
                
                yrSql = "Excess & Shortage for "
                
                lastRow = Cells(Rows.Count, "A").End(xlUp).Row
'looks 1 to 5 rows above current row
'if cells are equal, changes PB quantities to zero...removing duplicate counts
                For f = 6 To lastRow
                    h = 1
                    Do
                        If (.range("B" & f) = .range("B" & (f - h))) And (.range("C" & f) = .range("C" & (f - h))) _
                            Then .range("F" & f & ":L" & f).Value = "0"
                        h = h + 1
                    Loop Until h = 5
                Next f
'compares cells to find unit changes and places borders around them
                For g = 5 To lastRow
                    If (.range("A" & g) <> .range("A" & (g + 1))) Then
                        .range("A" & g & ":U" & g).Borders(xlEdgeBottom).Weight _
                        = XlBorderWeight.xlMedium
                    End If
                    
                    If linSearch.optYr = 2 Then
                        .range("I" & g).Value = (.range("J" & g) - .range("G" & g))
                    End If
                    If linSearch.optYr = 3 Then
                        .range("I" & g).Value = (.range("J" & g) - .range("H" & g))
                    End If
                Next g
                
                If linSearch.optYr = 1 Then
                    yrSql = yrSql & "CURRENT year Authorizations"
                ElseIf linSearch.optYr = 2 Then
                    yrSql = yrSql & "NEXT year Authorizations"
                ElseIf linSearch.optYr = 3 Then
                    yrSql = yrSql & "2 YEARS OUT Authorizations"
                End If
                
                .range("F1").Value = yrSql
                
                .range("A4:U4").AutoFilter
                .range("A4:V4").Interior.Color = RGB(224, 224, 224)
                .range("A4:V4").Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                .range("A4:V4").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                .range("A4:V4").Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                .range("A4:V4").HorizontalAlignment = xlLeft
                .range("F1").HorizontalAlignment = xlCenter
                .range("A4:V4").WrapText = True
                .range("C3, T3, B3").WrapText = True
                .range("C3, T3, B3, F1").Font.Italic = True
                .range("C3, T3, B3").Font.Size = 9
                .Rows("3:4").RowHeight = 30
                .range("F3:L3").Font.Bold = True
                
                lastRow = 0
                                   
            End With
'*****MATCAT_Info***************************************************
            With xlSheet2
            xlSheet2.Activate
            
                rangeStr = "A!10! B:B,E:E!25! C!9! D:D,F:F!12! I!40"
                MultiRangeColumn xlSheet2, rangeStr, "width"
                
                rangeStr = "A4!UIC! B4!Unit! C4!PBLIN! D4!NIIN! E4!Nomen! F4!On Hand! " & _
                            "G4!Pos! H4!Code! I4!Text"
                MultiRangeColumn xlSheet2, rangeStr, "values"
                
                .Columns("F:I").NumberFormat = "@"
                .range("A4:I4").AutoFilter
                
                lastRow = Cells(.Rows.Count, "A").End(xlUp).Row
                
Dim z As Long                   'iterates to bottom row of shading block
Dim y As Long                   'flips from even to odd, shades on even
                z = lastRow
                y = 2
'compares cells BOTTOM-UP to find changes in unit/item to shade blocks of rows
                For f = lastRow To 5 Step -1
'compares UIC_NIIN concatenation
                        If (.range("J" & f) <> .range("J" & (f - 1))) Then
'divide by 2 to see if y is even or odd
                            If y Mod 2 = 0 Then
'shades rows on every-other occurence
                                .range("A" & f & ":I" & z).Interior.Color = RGB(224, 224, 224)
                            End If
                                z = f - 1       'assigned row value ablove last row shaded
                                y = y + 1       'add +1, changes next Mod result
                        Else
'removes duplicate matching values
                            .range("A" & f & ":B" & f).Value = ""
                            .range("D" & f & ":E" & f).Value = ""
'compares PB LIN
'situation where catalog LIN is under a different PB LIN
                            If (.range("C" & f) <> .range("C" & (f - 1))) Then
                                If y Mod 2 = 0 Then
                                    .range("A" & f & ":I" & z).Interior.Color = RGB(224, 224, 224)
                                End If
                                    z = f - 1
                                    y = y + 1
                            Else
                                .range("C" & f & ":F" & f).Value = ""
                            End If
                        End If
                Next f
'remove UIC_NIIN concatenation column
                .Columns("J").EntireColumn.Delete
            End With
'*****PBIC-TAC-ERC***********************************************
            With xlSheet3
            xlSheet3.Activate
            
                rangeStr = "A!10! B:B,D:D!25! C!9! F!15! H:I!11! J:K!36"
                MultiRangeColumn xlSheet3, rangeStr, "width"
                
                rangeStr = "A4!UIC! B4!Unit! C4!LIN! D4!Nomen! E4!Cur Auth! F4!PBIC-TAC-ERC! G4!2nd Code! " & _
                            "H4!Multi-PBIC! I4!Multi-TAC! J4!Description! K4!2nd Desc"
                MultiRangeColumn xlSheet3, rangeStr, "values"
                
                .Columns("F:I").NumberFormat = "@"
                .Columns("L").EntireColumn.Delete
                .range("A4:K4").AutoFilter
                
                lastRow = Cells(.Rows.Count, "A").End(xlUp).Row
'combines data from 2 rows were multiple codes are for same item
'traverse through all rows BOTTOM-UP
                For f = lastRow To 5 Step -1
'compare all relevant cells in row
                        If (.range("A" & f) = .range("A" & (f - 1))) _
                            And (.range("H" & f) = .range("H" & (f - 1))) _
                            And (.range("I" & f) = .range("I" & (f - 1))) Then
'code (1)
                                If (.range("F" & (f - 1))) = "" Then    'if top cell is empty
                                    .range("F" & f).Copy                'copy from bottom cell
                                    .range("F" & (f - 1)).PasteSpecial  'paste to top cell
                                End If
'code (2)
                                If (.range("G" & (f - 1))) = "" Then
                                    .range("G" & f).Copy
                                    .range("G" & (f - 1)).PasteSpecial
                                End If
'code description
                                If (.range("J" & (f - 1))) = "" Then
                                    .range("J" & f).Copy
                                    .range("J" & (f - 1)).PasteSpecial
                                End If
                            .Rows(f).EntireRow.Delete                   'deletes bottom row
                        End If
                Next f
                
            End With
'*****PSD_Info***************************************************
            With xlSheet4
            xlSheet4.Activate
            
                rangeStr = "A:A,D:D,H:H,I:I!10! B:B,E:E!25! C!9! J!20! L:M!15"
                MultiRangeColumn xlSheet4, rangeStr, "width"
                
                rangeStr = "A4!UIC! B4!Unit! C4!LIN! D4!NIIN! E4!Nomen! F4!Due In! G4!Due Out! " & _
                            "H4!From UIC! I4!To UIC! J4!PSD ID! K4!Status! L4!Type! M4!Approved On"
                MultiRangeColumn xlSheet4, rangeStr, "values"
                
                .Columns("M").NumberFormat = "mm/dd/yyyy"
                .range("A4:M4").AutoFilter
                
                lastRow = Cells(.Rows.Count, "A").End(xlUp).Row
                
                rangeStr = "F3!=SUM(F5:F" & lastRow & ")! G3!=SUM(G5:G" & lastRow & ")"
                MultiRangeColumn xlSheet4, rangeStr, "formulas"
                
            End With
'*****ASIOE_Info**************************************************
            With xlSheet5
                xlSheet5.Activate
                
                rangeStr = "A:A,E:E,H:H!15! B:B,F:F,I:I!25! C:C,G:G!12! D!10! J!9"
                MultiRangeColumn xlSheet5, rangeStr, "width"
                
                rangeStr = "A2!PB LIN as Primary LIN to System! G2!NIIN Catalog LIN as Associated LIN in System! " & _
                            "A4!LIN as Primary! B4!Nomenclature! C4!Primary Qty! D4!Type! E4!Associated LIN! " & _
                            "F4!Nomenclature! G4!Associated Qty! H4!Part of LIN! I4!Nomenclature! " & _
                            "J4!Qty! K4!UIC"
                MultiRangeColumn xlSheet5, rangeStr, "values"
                
                .range("A4:K4").AutoFilter
                
                lastRow = Cells(.Rows.Count, "D").End(xlUp).Row
                z = lastRow
                y = 2
                
                For f = lastRow To 5 Step -1
                     If (.range("K" & f) <> .range("K" & (f - 1))) Or _
                        (.range("A" & f) <> .range("A" & (f - 1))) Then

                        If y Mod 2 = 0 Then
                            .range("A" & f & ":K" & z).Interior.Color = RGB(224, 224, 224)
                        End If
                        
                            z = f - 1
                            y = y + 1
                            
                    Else
                        If Len(.range("A" & f)) > 0 Then
                            .range("A" & f & ":D" & f & ", K" & f).Value = ""
                            .range("K" & f).Hyperlinks.Delete
                            
                        Else
                            .range("D" & f & ":F" & f).Value = ""

                            If (.range("C" & f) <> .range("C" & (f - 1))) Then
                                If y Mod 2 = 0 Then
                                    .range("A" & f & ":K" & z).Interior.Color = RGB(224, 224, 224)
                                End If
                                    z = f - 1
                                    y = y + 1
                            Else
                                .range("C" & f & ":F" & f).Value = ""
                            End If
                        End If
                    End If
                Next f
            End With
            
    xlSheet.Activate
            
    xlApp.Visible = True
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Sub

Public Sub CloseRecordsets()

With linSearch
    .rsResult.Close
    .rsMATCAT.Close
    .rsPBIC.Close
    .rsFromPSDs.Close
    .rsToPSDs.Close
End With

End Sub
