VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRemoveLIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub btnRunRemoveLIN_Click()

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
    ClearLIN
    Set linSearch = Nothing
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    Resume Exit_General

End Sub

Private Sub ClearLIN()
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
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempFromPsds'")) Then
        DoCmd.DeleteObject acTable, "tempFromPsds"
    End If
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempToPsds'")) Then
            DoCmd.DeleteObject acTable, "tempToPsds"
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

Private Sub cmdClose_Click()
    Forms!frmLINSearch.SetFocus
    DoCmd.Close acForm, "frmRemoveLIN"
End Sub
