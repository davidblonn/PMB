VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLINSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private p_lin1 As String                'form text box 1
Private p_lin2 As String                'form text box 2
Private p_optGroup As Integer           'OPTION GROUP value
Private p_optYr As Integer
Private p_deleteLRecords As String      'LINs selected for deletion
Private p_deleteNRecords As String      'NIINs selected for deletion
Private p_rsResult As DAO.Recordset     'sql from tempLINConsolidated
Private p_rsMATCAT As DAO.Recordset     'matches temp table to MATCAT table
Private p_rsPBIC As DAO.Recordset       'matches temp table to PBIC table
Private p_rsFromPSDs As DAO.Recordset   'matches temp table to FromUIC PSDs
Private p_rsToPSDs As DAO.Recordset     'matches temp table to ToUIC PSDs
Private p_rsPrmAsioe As DAO.Recordset
Private p_rsAscAsioe As DAO.Recordset

Private Sub btnQryLin1_Click()
'gets text box input from form
'builds temp table to add found LINs
'uses temp queries to search for all associated LINs
'displays temp queries on FORM LISTBOXES for deletion
'On Error GoTo Err_General
DoCmd.SetWarnings False

    LINInitialize Screen.ActiveForm
    
    If DCount("*", "tempGroupBy") = 0 Then
        MsgBox "ReCheck LIN(s). ZERO Results Returned"
        Exit Sub
    End If
    
    If MsgBox("Open LIN and NIIN Lists to Remove Items?", vbYesNo _
            + vbQuestion) = vbYes Then
                OpenRemoveLIN
    Else
        btnRestrictLINGroupBy_Click
    End If

Exit_General:
    DoCmd.SetWarnings False
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    Resume Exit_General
End Sub

Private Sub OpenRemoveLIN()

    DoCmd.openForm "frmRemoveLIN"
    Forms!frmRemoveLIN.Move Left:=7000, Top:=250

    With Forms!frmRemoveLIN!lstPickLINGroup
            .RowSourceType = "Table/Query"
            .RowSource = "tempGroupBy"
            .Requery
        End With

        With Forms!frmRemoveLIN!lstPickLINGroupBy
            .RowSourceType = "Table/Query"
            .RowSource = "tempGroupByDistinctLIN"
            .Requery
        End With
End Sub

Private Sub btnRestrictLINGroupBy_Click()

    ReportFinish
    
End Sub

Private Sub ReportFinish()

DoCmd.SetWarnings False

'build LINGroup object
    LINFinish Screen.ActiveForm
    
    Forms!frmLINSearch.SetFocus
    
    If CurrentProject.AllForms("frmRemoveLIN").IsLoaded Then
        Forms!frmRemoveLIN!lstPickLINGroup.RowSourceType = "Value List"
        Forms!frmRemoveLIN!lstPickLINGroup.RowSource = ""
        Forms!frmRemoveLIN!lstPickLINGroupBy.RowSourceType = "Value List"
        Forms!frmRemoveLIN!lstPickLINGroupBy.RowSource = ""
        DoCmd.Close acForm, "frmRemoveLIN"
    End If
    
    GoTo Exit_General

Exit_General:
'delete temp tables
'clear objects
Set linSearch = Nothing
    DoCmd.SetWarnings True
    btnClearLIN_Click
    Set linSearch = Nothing
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    Resume Exit_General

End Sub

Private Sub btnClearLIN_Click()
'used to remove all text box input and temp tables created
'used when LISTBOXES have been filled with LIN/NIIN info, but user
'no longer wants to continue with that info
DoCmd.SetWarnings False
'clear LISTBOX source

'delete temp tables
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempLINGroup'")) Then
        DoCmd.DeleteObject acTable, "tempLINGroup"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempGroupBy'")) Then
        DoCmd.DeleteObject acTable, "tempGroupBy"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempGroupByDistinctLIN'")) Then
        DoCmd.DeleteObject acTable, "tempGroupByDistinctLIN"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempMatchPBAuth'")) Then
            DoCmd.DeleteObject acTable, "tempMatchPBAuth"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempMatchNIIN'")) Then
        DoCmd.DeleteObject acTable, "tempMatchNIIN"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempLINConsolidated'")) Then
        DoCmd.DeleteObject acTable, "tempLINConsolidated"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempMatcat'")) Then
        DoCmd.DeleteObject acTable, "tempMatcat"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempPbic'")) Then
        DoCmd.DeleteObject acTable, "tempPbic"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempToPsds'")) Then
        DoCmd.DeleteObject acTable, "tempToPsds"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempFromPsds'")) Then
        DoCmd.DeleteObject acTable, "tempFromPsds"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempPrmAsioe'")) Then
        DoCmd.DeleteObject acTable, "tempPrmAsioe"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempAscAsioe'")) Then
        DoCmd.DeleteObject acTable, "tempAscAsioe"
    End If
'clear form text boxes
    Forms!frmLINSearch!lookUpLIN.Value = Null
    Forms!frmLINSearch!lookUpLIN2.Value = Null
    
DoCmd.SetWarnings True
End Sub

Private Sub Command10_Click()
'closes form
DoCmd.Close acForm, "frmLINSearch"
End Sub

Private Sub btnRadioSets_Click()
'calls Radio_Sets Module to open standard report
'on radio sets and receiver/transmitters
RadioSets

End Sub

Public Property Get lin1() As Variant
    lin1 = p_lin1
End Property
Public Property Let lin1(ByVal vNewValue As Variant)
    p_lin1 = vNewValue
End Property


Public Property Get lin2() As Variant
    lin2 = p_lin2
End Property
Public Property Let lin2(ByVal vNewValue As Variant)
    p_lin2 = vNewValue
End Property


Public Property Get optGroup() As Variant
    optGroup = p_optGroup
End Property
Public Property Let optGroup(ByVal vNewValue As Variant)
    p_optGroup = vNewValue
End Property


Public Property Get optYr() As Variant
    optYr = p_optYr
End Property
Public Property Let optYr(ByVal vNewValue As Variant)
    p_optYr = vNewValue
End Property


Public Property Get deleteLRecords() As Variant
    deleteLRecords = p_deleteLRecords
End Property
Public Property Let deleteLRecords(ByVal vNewValue As Variant)
    p_deleteLRecords = vNewValue
End Property


Public Property Get deleteNRecords() As Variant
    deleteNRecords = p_deleteNRecords
End Property
Public Property Let deleteNRecords(ByVal vNewValue As Variant)
    p_deleteNRecords = vNewValue
End Property


Public Property Get rsResult() As DAO.Recordset
    Set rsResult = p_rsResult
End Property
Public Property Set rsResult(rs As DAO.Recordset)
    Set p_rsResult = rs
End Property


Public Property Get rsMATCAT() As DAO.Recordset
    Set rsMATCAT = p_rsMATCAT
End Property
Public Property Set rsMATCAT(rs As DAO.Recordset)
    Set p_rsMATCAT = rs
End Property


Public Property Get rsPBIC() As DAO.Recordset
    Set rsPBIC = p_rsPBIC
End Property
Public Property Set rsPBIC(rs As DAO.Recordset)
    Set p_rsPBIC = rs
End Property


Public Property Get rsFromPSDs() As DAO.Recordset
    Set rsFromPSDs = p_rsFromPSDs
End Property
Public Property Set rsFromPSDs(rs As DAO.Recordset)
    Set p_rsFromPSDs = rs
End Property


Public Property Get rsToPSDs() As DAO.Recordset
    Set rsToPSDs = p_rsToPSDs
End Property
Public Property Set rsToPSDs(rs As DAO.Recordset)
    Set p_rsToPSDs = rs
End Property


Public Property Get rsPrmAsioe() As DAO.Recordset
    Set rsPrmAsioe = p_rsPrmAsioe
End Property
Public Property Set rsPrmAsioe(rs As DAO.Recordset)
    Set p_rsPrmAsioe = rs
End Property


Public Property Get rsAscAsioe() As DAO.Recordset
    Set rsAscAsioe = p_rsAscAsioe
End Property
Public Property Set rsAscAsioe(rs As DAO.Recordset)
    Set p_rsAscAsioe = rs
End Property
