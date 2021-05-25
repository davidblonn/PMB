Attribute VB_Name = "Radio_Sets"
Option Compare Database
Option Explicit

Public Sub RadioSets()

'builds a spreadsheet that matches the current radio equipment
'at units against the authorized quantity of different
'radio systems and the different components in each system

'************************************************************************************
'                  RETRIEVE RADIO SYSTEM AND RECEIVER/TRANSMITTER DATA
'************************************************************************************

Dim db As DAO.Database
Dim rs As DAO.Recordset     'radio system configurations
Dim rs2 As DAO.Recordset    'equipment on hand
Dim sql As String           'sql SELECT statements
    Set db = CurrentDb

DoCmd.SetWarnings False
'delete temp table if exists
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempRadioGroup'")) Then
        DoCmd.DeleteObject acTable, "tempRadioGroup"
    End If
'delete temp table if exists
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempRTGroup'")) Then
        DoCmd.DeleteObject acTable, "tempRTGroup"
    End If
'match all authorized radio data from units to radio configurations table
'calculate multiples on system equipment
'create temporary table
sql = "SELECT PBAuth.PBUIC, PBAuth.PBAuthLIN, PBAuth.CurAuth, " & _
            "PBAuth.Auth1, PBAuth.Auth2, Left(PBUIC, 4) AS Expr1, Units.Unit, tblRadioLINs.Req_RTs, " & _
            "tblRadioLINs.Veh_AMP_ADP, tblRadioLINs.PWR_AMP, tblRadioLINs.H250_Handset, tblRadioLINs.MANPACK_Antenna, " & _
            "tblRadioLINs.Ruck_Sack, tblRadioLINs.Batt_Box, tblRadioLINs.W2_Cable, tblRadioLINs.W4_Cable, " & _
            "tblRadioLINs.Locking_Bar, [Req_RTs]*[CurAuth] AS RTSum, [Req_RTs]*[Auth1] AS NextRTSum, [Req_RTs]*[Auth2] AS RTSumSecond, " & _
            "[Veh_AMP_ADP]*[CurAuth] AS VehSum, [PWR_AMP]*[CurAuth] AS PwrSum, " & _
            "[H250_Handset]*[CurAuth] AS HandSum, [MANPACK_Antenna]*[CurAuth] AS PackSum, [Ruck_Sack]*[CurAuth] AS RuckSum, " & _
            "[Batt_Box]*[CurAuth] AS BattSum, [W2_Cable]*[CurAuth] AS WTwoSum, [W4_Cable]*[CurAuth] AS WFourSum, " & _
            "[Locking_Bar]*[CurAuth] AS BarSum, " & _
            "Left(PBUIC, 4) AS BN " & _
            "INTO tempRadioGroup " & _
            "FROM Units " & _
            "INNER JOIN (tblRadioLINs INNER JOIN PBAuth ON tblRadioLINs.LIN = PBAuth.PBAuthLin) ON Units.UIC = PBAuth.PBUIC;"
'run SELECT INTO statement
            db.Execute sql
            DoCmd.OpenTable "tempRadioGroup"

'match all current RT LINs at units to those LINs in stored RT table
'create temporary table
sql = "SELECT NIINDetail.NiinUIC, NIINDetail.NiinCatLIN, NIINDetail.NiinOH, " & _
            "NIINDetail.NiinDueIn, NIINDetail.NiinDueOut, Left(NIINDetail.NiinUIC, 4) AS Expr1, NiinPBLIN, " & _
            "Left(NiinUIC, 4) AS BN, Units.Unit " & _
            "INTO tempRTGroup " & _
            "FROM Units " & _
            "INNER JOIN (tblRTLINs INNER JOIN NIINDetail ON tblRTLINs.RT_LIN = NIINDetail.NiinCatLIN) ON Units.UIC = NIINDetail.NiinUIC;"
'run SELECT INTO statement
            db.Execute sql
            DoCmd.OpenTable "tempRTGroup"

'calculate authorized sum of all system equipment GROUP BY unit UIC prefix
sql = "SELECT BN, Sum(RTSum) AS RTTotal, Sum(NextRTSum) AS NextRTTotal, Sum(RTSumSecond) AS RTSecond, First(Unit) AS Name, " & _
        "Sum(VehSum) AS SVehSum, Sum(PwrSum) AS SPwrSum, Sum(HandSum) AS SHandSum, Sum(PackSum) AS SPackSum, " & _
        "Sum(RuckSum) AS SRuckSum, Sum(BattSum) AS SBattSum, Sum(WTwoSum) AS W2Sum, Sum(WFourSum) AS W4Sum, " & _
        "Sum(BarSum) AS SBarSum, Sum(CurAuth) AS VRCAuth, Sum(Auth1) AS VRCAuth1, Sum(Auth2) AS VRCAuth2 " & _
        "FROM tempRadioGroup " & _
        "GROUP BY BN"
'create recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

'calculate on hand sum of all RTs GROUP BY unit UIC prefix
sql = "SELECT BN, Sum(NiinOH) AS OHSum, Sum(NiinDueIn) AS DISum, Sum(NiinDueOut) AS DOSum, First(Unit) AS Name " & _
        "FROM tempRTGroup " & _
        "GROUP BY BN"
'create recordset
    Set rs2 = db.OpenRecordset(sql, dbOpenSnapshot)
'delete temp tables
    DoCmd.Close acTable, "tempRadioGroup"
    DoCmd.Close acTable, "tempRTGroup"

'************************************************************************************
'                               BUILD SPREADSHEET
'************************************************************************************
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim rangeStr As String
    Dim e As Integer                'iterates rs2

    Set xlApp = Excel.Application
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    With xlSheet
        .name = "RT Count"
        .Cells.Font.name = "Calibri"
        .Cells.Font.Size = 11
        rangeStr = "A!20! B:R!10!"
        MultiRangeColumn xlSheet, rangeStr, "width"
        'build column headers
        rangeStr = "A1!Prepared on:! D1" & Date & "!" & "A2!Reason/Notes:! " & _
            "O3!ALL Class IX Amounts are based on Current Auth! A4!Unit! B4!UIC! " & _
            "C4!Auth'd VRCs! D4!Auth +1! E4!Auth +2! G4!Current RTs Auth! H4!Auth +1! " & _
            "I4!Auth +2! K4!RT On Hand! L4!RT Due In! M4!RT Due Out! O4!Veh AMP ADP! " & _
            "P4!PWR AMP! Q4!H250 Handset! R4!MANPACK Antenna! S4!Ruck Sack! T4!Batt Box! " & _
            "U4!W2 Cable! V4!W4 Cable! W4!Locking Bar! Y4! !"
        MultiRangeColumn xlSheet, rangeStr, "values"
        
        rangeStr = "C3!=SUM(C5:C150)! D3!=SUM(D5:D150)! E3!=SUM(E5:E150)! " & _
            "G3!=SUM(G5:G150)! H3!=SUM(H5:H150)! I3!=SUM(I5:I150)! " & _
            "K3!=SUM(J5:J150)! L3!=SUM(K5:K150)! M3!=SUM(L5:L150)"
        MultiRangeColumn xlSheet, rangeStr, "formulas"
               
        rangeStr = "A1!C1! A2!B2! C2!T2"
        MultiRangeColumn xlSheet, rangeStr, "merge"
        
        e = 5
        
'copy in recordset on radio equipment on hand
        Do While Not rs2.EOF
            rangeStr = "K" & e & "!" & Nz(rs2!OHSum, 0) & "!" & _
                        "L" & e & "!" & Nz(rs2!DISum, 0) & "!" & _
                        "M" & e & "!" & Nz(rs2!DOSum, 0)

'match UIC prefix to copy in authorized radio system configurations
                Do While Not rs.EOF
                    If rs!BN.Value = rs2!BN.Value Then
                        With rs
                            rangeStr = "A" & e & "!" & Nz(rs!name, "") & "!" & _
                                "B" & e & "!" & Nz(!BN, "") & "!" & _
                                "C" & e & "!" & Nz(!VRCAuth, "") & "!" & _
                                "D" & e & "!" & Nz(!VRCAuth1, "") & "!" & _
                                "E" & e & "!" & Nz(!VRCAuth2, "") & "!" & _
                                "G" & e & "!" & Nz(!RTTotal, "") & "!" & _
                                "H" & e & "!" & Nz(!NextRTTotal, "") & "!" & _
                                "I" & e & "!" & Nz(!RTSecond, "") & "!" & _
                                "O" & e & "!" & Nz(!SVehSum, "") & "!" & _
                                "P" & e & "!" & Nz(!SPwrSum, "") & "!" & _
                                "Q" & e & "!" & Nz(!SHandSum, "") & "!" & _
                                "R" & e & "!" & Nz(!SPackSum, "") & "!" & _
                                "S" & e & "!" & Nz(!SRuckSum, "") & "!" & _
                                "T" & e & "!" & Nz(!SBattSum, "") & "!" & _
                                "U" & e & "!" & Nz(!W2Sum, "") & "!" & _
                                "V" & e & "!" & Nz(!W4Sum, "") & "!" & _
                                "W" & e & "!" & Nz(!SBarSum, "")
                        End With
                        MultiRangeColumn xlSheet, rangeStr, "values"
                        Exit Do
'when no matching authorization
                    Else
                        .range("A" & e).Value = Nz(rs2!name, "")
                        .range("B" & e).Value = Nz(rs2!BN, "")
                    End If
                        rs.MoveNext
                
                Loop
'return to rs first record
                rs.MoveFirst
'iterate e
            e = e + 1
            rs2.MoveNext
        Loop
        
    Dim g As Integer        'traverse rows
    Dim lastRow As Long
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
'red fill for units showing less on hand than currently authorized
        For g = 3 To lastRow
            If (.range("K" & g) < .range("C" & g)) Then .range("K" & g).Interior.Color = RGB(255, 133, 133)
        Next g
        
        .range("A4:W4").Interior.Color = RGB(224, 224, 224)
        .range("A4:W4").Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
        .range("A4:W4").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        .range("A4:W4").Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        .range("A4:W4").HorizontalAlignment = xlLeft
        .range("A4:W4").VerticalAlignment = xlTop
        .range("A4:W4").WrapText = True
        .Rows(4).RowHeight = 30
            
    End With
        
    Set rs = Nothing            'to be reused
    Set xlSheet = Nothing       'reused for 2nd sheet
     
'match all authorized radio data from units on radio configurations table
'calculate multiples on system equipment
'GROUP BY radio system type
    sql = "SELECT First(PBUIC) AS UIC, First(Name) AS Unit, First(PBAuthLin) AS Lin, " & _
        "First(PBAuthNomen) AS Nomen, Sum(CurAuth) AS SumAuth, Sum(Auth1) AS SumAuth1, Sum(Auth2) AS SumAuth2, " & _
        "Sum([Req_RTs]*[CurAuth]) AS CurRT, Sum([Req_RTs]*[Auth1]) AS RTAuth1, Sum([Req_RTs]*[Auth2]) AS RTAuth2, " & _
        "Left(PBUIC, 4) AS BN " & _
        "FROM tblRadioLINs INNER JOIN PBAuth ON tblRadioLINs.LIN = PBAuth.PBAuthLin " & _
        "GROUP BY Lin, Left(PBUIC, 4)"
'create recordset
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    
    Set xlSheet = xlBook.Worksheets(2)
    
'copy recordset data to worksheet 2
    With xlSheet
        .name = "Auth VRCs"
        rangeStr = "B1!UIC! C1!Unit! D1!Auth'd VRC! E1!NOMEN! F1!Auth VRC! " & _
            "G1!Auth VRC +1! H1!Auth VRC +2! I1!Auth RT! J1!Auth RT +1! " & _
            "K1!Auth RT +2! L1!BN"
        MultiRangeColumn xlSheet, rangeStr, "values"
        
        .range("B2").CopyFromRecordset rs
        .Cells.EntireColumn.AutoFit
        .range("F:K").ColumnWidth = 13.5
        .range("B1:L1").AutoFilter
    End With
'clear objects
    rs.Close
    rs2.Close
    Set rs = Nothing
    Set rs2 = Nothing

    xlApp.Visible = True

    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
'delete temp tables
    DoCmd.DeleteObject acTable, "tempRadioGroup"
    DoCmd.DeleteObject acTable, "tempRTGroup"

    DoCmd.SetWarnings True

End Sub
