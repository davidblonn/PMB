VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private p_xlApp As Excel.Application
Private p_xlBook As Workbook            'copy from
Private p_xlCopySheet As Worksheet      'copy from
Private p_xlPasteSheet As Worksheet     'paste to
Private p_fType As String
Private p_sheetName As String
Private p_sourceFile As String
Private p_lastRow As Long

Public Function ValidateFileFolder(b As Boolean) As Boolean
    
    FileFolderCheck p_sourceFile, p_fType, b, p_sheetName
        If b = False Then Exit Function

End Function

Public Sub DataTransfer(linkS As newSheet)

Dim tableName As String
Dim sql As String

    tableName = p_sheetName & "!"
    sql = "DELETE * FROM " & p_sheetName & " WHERE "
    
    Select Case p_sheetName
        Case "PBIC_TAC_ERC"
            sql = sql & "Description IS NULL;"
        Case "PSDs"
            sql = sql & "PSDID IS NULL;"
        Case "SB700"
            sql = sql & "AppH1AuthLIN IS NULL;"
        Case "NIINDetail"
            sql = sql & "NiinUIC IS NULL;"
        Case "PBAuth"
            sql = sql & "PBUIC IS NULL;"
        Case "Units"
            sql = sql & "UIC IS NULL;"
        Case "MATCAT"
            sql = sql & "POS_CODE IS NULL;"
        Case "MVMT_TYPE"
            sql = sql & "MVMT_Code IS NULL;"
        Case "ASIOE_CMI"
            sql = sql & "Prm_LIN IS NULL;"
        Case "NewAR"
            sql = Left(sql, 14) & "tblActRegAppend WHERE UIC IS NULL;"
            DoCmd.RunSavedImportExport ("LinkedSheetsNewAR")
            DoCmd.RunSQL sql
            Exit Sub
    End Select
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, _
            p_sheetName, linkS.sourceFile, True, tableName
        CurrentDb.TableDefs.Refresh
        DoCmd.RunSQL sql

End Sub

Public Sub SetFieldNames(pasteLastRow As Long)

Select Case p_sheetName

    Case "PSDs"
        With p_xlPasteSheet
            Dim r As range
            Dim d As String
            For Each r In .range("U2:V" & pasteLastRow)
                d = Left(r.Value, 11)
                r.Value = d
                r.NumberFormat = "mm/dd/yyyy"
            Next r
            
            .range("A1").Value = "IsPassThru"
            .range("B1").Value = "PSDID"
            .range("C1").Value = "FromState"
            .range("D1").Value = "FromUIC"
            .range("E1").Value = "ToState"
            .range("F1").Value = "ToUIC"
            .range("G1").Value = "Validated"
            .range("H1").Value = "Requested"
            .range("I1").Value = "CondCode"
            .range("J1").Value = "FromPBLIN"
            .range("K1").Value = "ToPBLIN"
            .range("L1").Value = "SupplyClass"
            .range("M1").Value = "AAC"
            .range("N1").Value = "Unit$"
            .range("O1").Value = "Ext$"
            .range("P1").Value = "NiinNomen"
            .range("Q1").Value = "SourceNiin"
            .range("R1").Value = "VetLevel"
            .range("S1").Value = "Suspense"
            .range("T1").Value = "CatLIN"
            .range("U1").Value = "CreatedOn"
            .range("V1").Value = "ApprovedOn"
            .range("W1").Value = "PSDStatus"
            .range("X1").Value = "Gain_DOC"
            .range("Y1").Value = "DOC_NUM"
            .range("Z1").Value = "Lose_DOC"
            .range("AA1").Value = "CloseDate"
            .range("AB1").Value = "DaysSince"
            .range("AC1").Value = "DaysUntil"
            .range("AD1").Value = "DirectiveSent"
            .range("AE1").Value = "VetLevelAssignDate"
            .range("AF1").Value = "VetStatus"
            .range("AG1").Value = "CreatedBy"
            .range("AH1").Value = "Type"
        End With
        
    Case "NIINDetail"
        With p_xlPasteSheet
            .range("A1").Value = "NiinUIC"
            .range("B1").Value = "NiinPBLIN"
            .range("C1").Value = "PrimaryPBLIN"
            .range("D1").Value = "ASIOE"
            .range("E1").Value = "CMI"
            .range("F1").Value = "ERC_P"
            .range("G1").Value = "ERC_T"
            .range("H1").Value = "ERC_A-C"
            .range("I1").Value = "NiinOH"
            .range("J1").Value = "NiinDueIn"
            .range("K1").Value = "NiinDueOut"
            .range("L1").Value = "NiinCatLIN"
            .range("M1").Value = "NiinCatLINNomen"
            .range("N1").Value = "NIIN"
            .range("O1").Value = "NiinNomen"
            .range("P1").Value = "Match"
            .range("Q1").Value = "SCMC"
            .range("R1").Value = "TAC"
            .range("S1").Value = "PBIC"
            .range("T1").Value = "ARC"
            .range("U1").Value = "MATCAT"
            .range("V1").Value = "UnitPrice"
            .range("W1").Value = "ExtPrice"
            .range("X1").Value = "MDLStatus"
            .range("Y1").Value = "MAT_Pos_1"
            .range("Z1").Value = "MAT_Pos_2"
            .range("AA1").Value = "MAT_Pos_3"
            .range("AB1").Value = "MAT_Pos_4"
            .range("AC1").Value = "UIC_NIIN"
        End With
        
    Case "PBAuth"
        With p_xlPasteSheet
            .range("A1").Value = "PBUIC"
            .range("B1").Value = "Name"
            .range("C1").Value = "Location"
            .range("D1").Value = "PBAuthLin"
            .range("E1").Value = "PBAuthNomen"
            .range("F1").Value = "CurAuth"
            .range("G1").Value = "Auth1"
            .range("H1").Value = "Auth2"
            .range("I1").Value = "PercentFill"
            .range("J1").Value = "PlusMinus"
            .range("K1").Value = "All ERCs"
            .range("L1").Value = "Darpl"
            .range("M1").Value = "PB LIN Major Capability"
            .range("N1").Value = "PBOnHand"
            .range("O1").Value = "PBDueIn"
            .range("P1").Value = "PBDueOut"
            .range("Q1").Value = "MATCAT"
            .range("R1").Value = "ARC"
            .range("S1").Value = "Catalog LIN"
            .range("T1").Value = "TAC"
            .range("U1").Value = "PBIC"
            .range("V1").Value = "NIIN"
            .range("W1").Value = "Unit Price"
            .range("X1").Value = "Extended Price"
            .range("Y1").Value = "IsCTA"
            .range("Z1").Value = "SCMC"
            .range("AA1").Value = "MDL Status"
            .range("AB1").Value = "FoundPTE"
            .range("AC1").Value = "Multi-PBIC"
            .range("AD1").Value = "Multi-TAC"
            .range("AE1").Value = "SecondFoundPTE"
            .range("AF1").Value = "UIC_LIN"
        End With
        
    Case "NewAR"
        With p_xlPasteSheet
            .range("A1").Value = "UIC"
            .range("B1").Value = "DOC_NUM"
            .range("C1").Value = "NSN"
            .range("D1").Value = "LIN"
            .range("E1").Value = "NOMEN"
            .range("F1").Value = "DOC_ID_CD"
            .range("G1").Value = "UI"
            .range("H1").Value = "UIC_GAIN"
            .range("I1").Value = "GAIN_LOSE_"
            .range("J1").Value = "MVMT_TYPE"
            .range("K1").Value = "DATE_CLOSE"
            '.range("L1").Value = "TRANS_DATE"
            .range("L1").Value = "SERIAL_NUM"
            .range("M1").Value = "LastUpdate"
            .range("N1").Value = "FORM_NUM"
            '.range("P1").Value = "TRANS_YEAR"
            '.range("P2:P" & pasteLastRow).Formula = "=MID(B2,7,1)"
        End With
        
    Case "SB700"
        With p_xlPasteSheet
            .range("A1").Value = "AppH1AuthLIN"
            .range("B1").Value = "AppH1AuthNomen"
            .range("C1").Value = "AppH1SubLIN"
            .range("D1").Value = "AppH1SubNomen"
        End With
    
    End Select

End Sub

Public Property Get xlApp() As Excel.Application
    Set xlApp = p_xlApp
End Property
Public Property Set xlApp(xlApp As Excel.Application)
    Set p_xlApp = xlApp
End Property

Public Property Get xlBook() As Workbook
    Set xlBook = p_xlBook
End Property
Public Property Set xlBook(xlBook As Workbook)
    Set p_xlBook = xlBook
End Property

Public Property Get xlCopySheet() As Worksheet
    Set xlCopySheet = p_xlCopySheet
End Property
Public Property Set xlCopySheet(xlCopySheet As Worksheet)
    Set p_xlCopySheet = xlCopySheet
End Property

Public Property Get xlPasteSheet() As Worksheet
    Set xlPasteSheet = p_xlPasteSheet
End Property
Public Property Set xlPasteSheet(xlPasteSheet As Worksheet)
    Set p_xlPasteSheet = xlPasteSheet
End Property

Public Property Get fType() As Variant
    fType = p_fType
End Property
Public Property Let fType(ByVal vNewValue As Variant)
    p_fType = vNewValue
End Property

Public Property Get sheetName() As Variant
    sheetName = p_sheetName
End Property
Public Property Let sheetName(ByVal vNewValue As Variant)
    p_sheetName = vNewValue
End Property

Public Property Get sourceFile() As Variant
    sourceFile = p_sourceFile
End Property
Public Property Let sourceFile(ByVal vNewValue As Variant)
    p_sourceFile = vNewValue
End Property

Public Property Get lastRow() As Variant
    lastRow = p_lastRow
End Property
Public Property Let lastRow(ByVal vNewValue As Variant)
    p_lastRow = vNewValue
End Property


