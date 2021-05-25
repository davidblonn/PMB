VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFilePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim s As Integer                'countdown to save
Dim cProg As New clsLblProg         'progress bar

Private Sub btnLinkedSheets_Click()
'calls SUBs of all files being updated into LINKEDSHEETS excel spreadsheet
'based on OPTION BUTTON selections
'run SQL to UPDATE database tables based on new linked table items
'On Error GoTo Err_General
DoCmd.SetWarnings False

CloseExcel

s = 19
    With cProg
        .Initialize Forms!frmFilePicker!lblBack, _
            Forms!frmFilePicker!lblFront, Forms!frmFilePicker!lblCaption, _
            Forms!frmFilePicker!lblCurrent
        
        If Forms!frmFilePicker!optUpdate = True Then
            If Forms!frmFilePicker!optBlueFiles = True Then
                .Max = s
            Else
                .Max = s - 5
            End If
        Else
            .Max = s - 6
        End If
    End With
    
    If Forms!frmFilePicker!optPBICMATCAT = True Then
            Forms!frmFilePicker!optPBICMATCAT = False
        End If

    If Forms!frmFilePicker!optUpdate = True Then
        ActRegLinkedSheets s
            cProg.Increment "Activity Register"
    Else
        s = s - 1                       '(17)
    End If

    If Forms!frmFilePicker!optBlueFiles = True Then
        PBICLinkedSheets s        'PBIC-TAC-ERC File
            cProg.Increment "PBIC_TAC_ERC"
        SB700LinkedSheets s       'SB 700-20 File
            cProg.Increment "SB700-20"
        UnitsLinkedSheets s       'Units File
            cProg.Increment "Units"
        MATCATLinkedSheets s      'MATCAT File
            cProg.Increment "MATCAT"
        MVMT_TYPELinkedSheets s   'Movement Type File
            cProg.Increment "Movement Type"
        ASIOE_CMILinkedSheets s
            cProg.Increment "ASIOE CMI"
    Else
        s = s - 5                       '(12)
    End If
    
    PSDsLinkedSheets s, cProg     'All (4) PSD Files

    WVARNGLinkedSheets s, cProg   'WVARNG_All File
    
    MatchedPBICTACERC s             'UPDATES table
        cProg.Increment "Matched Tables SQL"
's will equal (1) here and trigger save in next sub
    MatchedMATCAT s, cProg          'UPDATES table

    'MsgBox "Spreadsheet Transfer Complete"

Exit_General:
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General
End Sub

Private Sub btnActivityReg_Click()
'choose the file to be used in updating the LINKEDSHEETS spreadsheet
Dim label As String
label = "ActReg"
FileUpdate label        'choose file and UPDATE table with file path
End Sub

Private Sub btnIncoming_Click()
Dim label As String
label = "PSD3"
FileUpdate label
End Sub

Private Sub btnLinkedExcelFile_Click()
Dim label As String
label = "LinkedSheet"
FileUpdate label
End Sub

Private Sub btnOpenVetting_Click()
Dim label As String
label = "PSD4"
FileUpdate label
End Sub

Private Sub btnOutgoing_Click()
Dim label As String
label = "PSD2"
FileUpdate label
End Sub

Private Sub btnPBIC_Click()
Dim label As String
label = "PBIC"
FileUpdate label
End Sub

Private Sub btnMATCAT_Click()
Dim label As String
label = "MATCAT"
FileUpdate label
End Sub

Private Sub btnMVMT_TYPE_Click()
Dim label As String
label = "MVMT_TYPE"
FileUpdate label
End Sub

Private Sub btnASIOE_CMI_Click()
Dim label As String
label = "ASIOE_CMI"
FileUpdate label
End Sub

Private Sub btnSB700_Click()
Dim label As String
label = "SB700"
FileUpdate label
End Sub

Private Sub btnTurnIn_Click()
Dim label As String
label = "PSD1"
FileUpdate label
End Sub

Private Sub btnUnits_Click()
Dim label As String
label = "Units"
FileUpdate label
End Sub

Private Sub btnWVARNG_All_Click()
Dim label As String
label = "UnitLIN"
FileUpdate label
End Sub

Public Sub FileUpdate(label As String)
'uses windows FILE DIALOG to select new file to link to
On Error GoTo Err_General
DoCmd.SetWarnings False

Dim filePath As String              'file path
Dim diaFile As FileDialog

Set diaFile = Application.FileDialog(msoFileDialogFilePicker)
With diaFile
    .Filters.Add "Excel Files", "*.xlsx"
    .AllowMultiSelect = False
    .Show
End With
'assign new file path
filePath = diaFile.SelectedItems(1)
'run SQL to UPDATE table with last chosen file
DoCmd.RunSQL "UPDATE tblFilePicker SET " & label & _
        " = '" & filePath & "'"
Me.Refresh

Exit_General:
    Set diaFile = Nothing
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General
End Sub

Private Sub Command10_Click()
'close form
DoCmd.Close acForm, "frmFilePicker"
End Sub

Private Sub btnShowButtons_Click()
'makes visible command buttons to run single file updates
    Forms!frmFilePicker!btnSoloPBIC.Visible = True
    Forms!frmFilePicker!btnSoloPSDs.Visible = True
    Forms!frmFilePicker!btnSoloPBAuth.Visible = True
    Forms!frmFilePicker!btnSoloActReg.Visible = True
    Forms!frmFilePicker!btnSoloSB700.Visible = True
    Forms!frmFilePicker!btnSoloUnits.Visible = True
    Forms!frmFilePicker!btnSoloMATCAT.Visible = True
    Forms!frmFilePicker!btnSoloMVMT_TYPE.Visible = True
    Forms!frmFilePicker!btnSoloASIOE_CMI.Visible = True
    Forms!frmFilePicker!optPBICMATCAT.Visible = True
    Forms!frmFilePicker!btnHideButtons.Visible = True
    Forms!frmFilePicker!btnHideButtons.SetFocus
    Forms!frmFilePicker!btnShowButtons.Visible = False

End Sub

Private Sub btnHideButtons_Click()
'rehides single file update buttons
    Forms!frmFilePicker!btnShowButtons.Visible = True
    Forms!frmFilePicker!btnShowButtons.SetFocus
    Forms!frmFilePicker!btnSoloPBIC.Visible = False
    Forms!frmFilePicker!btnSoloPSDs.Visible = False
    Forms!frmFilePicker!btnSoloPBAuth.Visible = False
    Forms!frmFilePicker!btnSoloActReg.Visible = False
    Forms!frmFilePicker!btnSoloSB700.Visible = False
    Forms!frmFilePicker!btnSoloUnits.Visible = False
    Forms!frmFilePicker!btnSoloMATCAT.Visible = False
    Forms!frmFilePicker!btnSoloMVMT_TYPE.Visible = False
    Forms!frmFilePicker!btnSoloASIOE_CMI.Visible = False
    Forms!frmFilePicker!optPBICMATCAT.Visible = False
    Forms!frmFilePicker!btnHideButtons.Visible = False

End Sub

Private Sub btnSoloPSDs_Click()
'to call the SUBs to update only a single worksheet in LINKEDSHEETS
s = 5
    PSDsLinkedSheets s, cProg
End Sub

Private Sub btnSoloPBAuth_Click()
s = 5
    WVARNGLinkedSheets s, cProg
End Sub

Private Sub btnSoloActReg_Click()
s = 1
    ActRegLinkedSheets s
End Sub

Private Sub btnSoloSB700_Click()
s = 1
    SB700LinkedSheets s
End Sub

Private Sub btnSoloUnits_Click()
s = 1
    UnitsLinkedSheets s
End Sub

Private Sub btnSoloPBIC_Click()
s = 1
    PBICLinkedSheets s
End Sub

Private Sub btnSoloMATCAT_Click()
s = 1
    MATCATLinkedSheets s
End Sub

Private Sub btnSoloMVMT_TYPE_Click()
s = 1
    MVMT_TYPELinkedSheets s
End Sub

Private Sub btnSoloASIOE_CMI_Click()
s = 1
    ASIOE_CMILinkedSheets s
End Sub
