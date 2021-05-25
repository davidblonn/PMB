VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSelectUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private p_sDate As Date                 'start date
Private p_eDate As Date                 'end date
Private p_optUG As Integer              'option group value
Private p_rsCompare As DAO.Recordset    'compare against list box
Private p_searchFolder As String
Private p_sql As String
Private p_results As String
Private p_d As Boolean

Public curForm As Form

Private Sub btnLoadUnits_Click()

    Set curForm = Screen.ActiveForm

    UnitsInitialize curForm
    
    If Len(unitSelect.sql) < 1 Then
        Exit Sub
    End If
    
Dim resultText As String
    
    resultText = "Units Selected:" & _
        vbCrLf & vbCrLf & unitSelect.results & vbCrLf & vbCrLf
        
    If unitSelect.d = True Then
        resultText = resultText & "For Dates: " _
            & unitSelect.sDate & " to " & unitSelect.eDate
    End If
        
    If Len(unitSelect.searchFolder) > 3 Then
        resultText = resultText & vbCrLf & vbCrLf & _
                "Unit Folder to Save Activity Register:" & _
                vbCrLf & vbCrLf & unitSelect.searchFolder
    End If
    
    If CurrentProject.AllForms("frmActRegSearch").IsLoaded Then
        Forms!frmActRegSearch!lblResults = resultText
        Forms!frmActRegSearch!lblResults.Visible = True
        Forms!frmActRegSearch!btnPMBDocFiling.Visible = True
        Forms!frmActRegSearch!btnMonthlyReview.Visible = True
        Forms!frmActRegSearch!btnNewSelection.Visible = True
        Forms!frmActRegSearch!btnNewSelection.SetFocus
        Forms!frmActRegSearch!btnOpenUnits.Visible = False
    End If
    
    If CurrentProject.AllForms("frmOffReports").IsLoaded Then
        Forms!frmOffReports!lblResults = resultText
        Forms!frmOffReports!lblResults.Visible = True
        Forms!frmOffReports!btnNewSelection.Visible = True
        Forms!frmOffReports!btnNewSelection.SetFocus
        Forms!frmOffReports!btnOpenUnits.Visible = False
    End If
    
    DoCmd.DeleteObject acTable, "tempRsResults"

    btnCancel_Click
    Set curForm = Nothing
    
End Sub

Private Sub cmdFolder_Click()
'uses windows FILE DIALOG to select ROOT FOLDER for folder searches
'used in PMB DOCUMENT FILING
On Error GoTo Err_General
DoCmd.SetWarnings False

Dim diaFolder As FileDialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
        With diaFolder
            .AllowMultiSelect = False
            .Show
            If (.SelectedItems.Count = 0) Then
                MsgBox "Folder to Search MUST be Selected"
                Exit Sub
            Else
                Me.lblUnitFolder = .SelectedItems(1)
            End If
        End With
    Set diaFolder = Nothing
Exit_General:
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & " - " & Err.Description
    GoTo Exit_General
End Sub

Public Sub btnUnselect_Click()
'deselects all units in LISTBOX
Set curForm = Screen.ActiveForm
DeselectAll curForm
End Sub

Public Sub btnSelectAll_Click()
'selects all units in LISTBOX
Set curForm = Screen.ActiveForm
SelectAll curForm
End Sub

Private Sub btnCancel_Click()
DoCmd.Close acForm, "frmSelectUnits"
End Sub


Public Property Get rsCompare() As DAO.Recordset
    Set rsCompare = p_rsCompare
End Property
Public Property Set rsCompare(ByVal rs As DAO.Recordset)
    Set p_rsCompare = rs
End Property


Public Property Get sDate() As Date
    sDate = p_sDate
End Property
Public Property Let sDate(ByVal s As Date)
    p_sDate = s
End Property


Public Property Get eDate() As Date
    eDate = p_eDate
End Property
Public Property Let eDate(ByVal e As Date)
    p_eDate = e
End Property


Public Property Get optUG() As Integer
    optUG = p_optUG
End Property
Public Property Let optUG(ByVal og As Integer)
    p_optUG = og
End Property


Public Property Get searchFolder() As String
    searchFolder = p_searchFolder
End Property
Public Property Let searchFolder(ByVal str As String)
    p_searchFolder = str
End Property


Public Property Get sql() As String
    sql = p_sql
End Property
Public Property Let sql(ByVal str As String)
    p_sql = str
End Property


Public Property Get results() As String
    results = p_results
End Property
Public Property Let results(ByVal str As String)
    p_results = str
End Property


Public Property Get d() As Boolean
    d = p_d
End Property
Public Property Let d(ByVal a As Boolean)
    p_d = a
End Property
