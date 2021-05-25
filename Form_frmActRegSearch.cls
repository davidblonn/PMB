VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmActRegSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private p_fType As String                   'report type
Private p_sDate As Date                     'start date
Private p_eDate As Date                     'end date
Private p_sql As String                     'completed SQL statement
Private p_strIN As String                   'completed units list
Private p_optUG As Integer                  'option group value
Private p_saveFolder As String              'path to save file
Private p_myFolder As String                'folder to search/save
Private p_rsActReg As DAO.Recordset         'queried activity register records
Private p_rsFoundDOCs As DAO.Recordset      'table of found documents
Private p_xlBook As Excel.Workbook

Dim xlApp As Excel.Application
Dim a As Boolean

Private Sub btnPMBDocFiling_Click()

DoCmd.SetWarnings False

        If IsNull(unitSelect.sDate) Then
            MsgBox "Must Select START DATE to Run Report"
            GoTo Exit_General
        End If
        
    ActRegInitialize Screen.ActiveForm, "docFiling"
    
    GoTo Exit_General
    
Exit_General:
'close recordset
'clear objects
'delete temp table
    Set regReport = Nothing
    
    DoCmd.DeleteObject acTable, "tempFoundDOCs"
    DoCmd.SetWarnings True
    Exit Sub
    
Err_General:
    MsgBox Err.Number & " - " & Err.Description
    GoTo Exit_General

End Sub


'******************************* folder dialog to pick folder to search **********************************
Private Sub btnDocNumSearch_Click()
'Creates Excel spreadsheet of all found documents with
'hyperlinks to those documents
'uses START and END DATES, and selected ROOT FOLDER
'On Error GoTo Err_General
DoCmd.SetWarnings False

    a = True
'validate unit folder for searching and saving and dates for report
    FileFolderCheck Me.lblFolder, "folder", a
        If a = False Then
            MsgBox "Invalid Folder Selected"
            GoTo Exit_General
        End If
        
        If IsNull((unitSelect.sDate) Or _
                    (unitSelect.eDate)) Then
            MsgBox "Invalid Dates"
            GoTo Exit_General
        End If
        
        ActRegInitialize Screen.ActiveForm, "docSearch"

        GoTo Exit_General
    
Exit_General:
'close recordset
'clear objects
'delete temp table
    Set regReport = Nothing
    DoCmd.DeleteObject acTable, "tempFoundDOCs"
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & " - " & Err.Description
    GoTo Exit_General
End Sub
'***********************************************************************************************

Private Sub btnMonthlyReview_Click()
'Creates Access report of all activity register for chosen unit(s) and
'Excel spreadsheet of all missing documents
'uses START and END DATES, unit LISTBOX or OPTION GROUP
'On Error GoTo Err_General
DoCmd.SetWarnings False

    a = True
'validate unit folder for searching and saving and dates for report
    FileFolderCheck Me.lblSaveFolder, "folder", a
        If a = False Then
            MsgBox "Invalid Folder Selected"
            GoTo Exit_General
        End If
        
        If IsNull(unitSelect.sDate) Then
            MsgBox "Must Select START DATE to Run Report"
            GoTo Exit_General
        End If
        
        ActRegInitialize Screen.ActiveForm, "docMonthly"
    
    GoTo Exit_General

Exit_General:
'close recordset
'clear objects
'delete temp table
    Set regReport = Nothing
    DoCmd.DeleteObject acTable, "tblUnitActReg"
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & " - " & Err.Description
    GoTo Exit_General
End Sub

'*******************************************************************************
Private Sub btnDocSearch_Click()
'creates Excel spreadsheet of all activity found matching
'document number input
'uses DOC NUM SEARCH text label
'On Error GoTo Err_General
DoCmd.SetWarnings False

    Set xlApp = Excel.Application
    
    ActRegInitialize Screen.ActiveForm, "docNumSearch"

    xlApp.Visible = True

Exit_General:
'clear form text field
    Set regReport = Nothing
    Set xlApp = Nothing
    Me.lblNumSearch.Value = ""
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & " - " & Err.Description
    GoTo Exit_General
End Sub

Private Sub btnLINSearch_Click()
'creates Excel spreadsheet of all activity found matching
'LIN input
'uses LIN SEARCH text label
'On Error GoTo Err_General
DoCmd.SetWarnings False

    Set xlApp = Excel.Application
    
    ActRegInitialize Screen.ActiveForm, "linSearch"

    xlApp.Visible = True

Exit_General:
'clear form text field
    Set regReport = Nothing
    Set xlApp = Nothing
    Me.lblNumSearch.Value = ""
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & " - " & Err.Description
    GoTo Exit_General
End Sub

Private Sub cmdFolder_Click()
'uses windows FILE DIALOG to select ROOT FOLDER for folder searches
'used in FILE SEARCHES
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
                Me.lblFolder = .SelectedItems(1)
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

Private Sub btnOpenUnits_Click()
    DoCmd.openForm "frmSelectUnits"
    Forms!frmSelectUnits.Move Left:=7500, Top:=250

End Sub

Private Sub btnNewSelection_Click()
    Me.lblResults = ""
    Me.lblResults.Visible = False
    Forms!frmActRegSearch!btnPMBDocFiling.Visible = False
    Forms!frmActRegSearch!btnMonthlyReview.Visible = False
    btnOpenUnits_Click
End Sub

Private Sub Command10_Click()
'closes form
DoCmd.Close acForm, "frmActRegSearch"
End Sub


Public Property Get fType() As String
    fType = p_fType
End Property
Public Property Let fType(ByVal str As String)
    p_fType = str
End Property


Public Property Get sDate() As Date
    sDate = p_sDate
End Property
Public Property Let sDate(ByVal d As Date)
    p_sDate = d
End Property


Public Property Get eDate() As Date
    eDate = p_eDate
End Property
Public Property Let eDate(ByVal d As Date)
    p_eDate = d
End Property


Public Property Get sql() As String
    sql = p_sql
End Property
Public Property Let sql(ByVal s As String)
    p_sql = s
End Property


Public Property Get strIN() As String
    strIN = p_strIN
End Property
Public Property Let strIN(ByVal i As String)
    p_strIN = i
End Property


Public Property Get optUG() As Integer
    optUG = p_optUG
End Property
Public Property Let optUG(ByVal i As Integer)
    p_optUG = i
End Property


Public Property Get saveFolder() As String
    saveFolder = p_saveFolder
End Property
Public Property Let saveFolder(ByVal str As String)
    p_saveFolder = str
End Property


Public Property Get myFolder() As String
    myFolder = p_myFolder
End Property
Public Property Let myFolder(ByVal str As String)
    p_myFolder = str
End Property


Public Property Get rsActReg() As DAO.Recordset
    Set rsActReg = p_rsActReg
End Property
Public Property Set rsActReg(rs As DAO.Recordset)
    Set p_rsActReg = rs
End Property


Public Property Get rsFoundDOCs() As DAO.Recordset
    Set rsFoundDOCs = p_rsFoundDOCs
End Property
Public Property Set rsFoundDOCs(rs As DAO.Recordset)
    Set p_rsFoundDOCs = rs
End Property


Public Property Get xlBook() As Excel.Workbook
    Set xlBook = p_xlBook
End Property
Public Property Set xlBook(xl As Excel.Workbook)
    Set p_xlBook = xl
End Property
