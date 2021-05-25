VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOffReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private p_sheetName As String           'from combo box, report name
Private p_fileName As String            'full file path
Private p_filePath As String            'folder path
Private p_objType As String             'report type
Private p_optUG As Integer
Private p_strIN As String               'selected units
Private p_rsSQL As String               'concatenated sql
Private p_rsReport As DAO.Recordset     'sql results
Private p_xlBook As Excel.Workbook

Private Sub btnRunReport_Click()
'creates a PSD report based on the type of report chosen in COMBOBOX
'uses unit selection LISTBOX and OPTION GROUP
DoCmd.SetWarnings False
'On Error GoTo Err_General

Dim filePath As String
Dim b As Boolean            'validates file

    b = True
    filePath = Forms!frmOffReports!lblReportFolder
'validate that folder to save in is good
    FileFolderCheck filePath, "folder", b

    If b = False Then
        MsgBox "Folder Selected is no longer valid"
        GoTo Exit_General
    End If
    
    ReportsInitialize Screen.ActiveForm
    
    Me.lblResults = ""
    Me.lblResults.Visible = False

GoTo Exit_General

Exit_General:
    Set unitSelect = Nothing
    Set newReport = Nothing
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General

End Sub

Private Sub btnStaffReport_Click()
'uses previous STAFF REPORT to create temp table for comparison
'creates STAFF REPORT and saves it to chosen folder with new name
DoCmd.SetWarnings False
On Error GoTo Err_General

Dim b As Boolean                'file/folder validation

    b = True
    'check that all necessary files/folders have been chosen
    If IsNull(Me.lblStaffReport) And _
            IsNull(Me.lblSRCompare) And _
            IsNull(Me.lblReportFolder) Then
        MsgBox "Excel Staff Report Template and" & vbCrLf & _
            "Previous Report to Compare MUST be Selected"
        Exit Sub
    End If
    'validate each file/folder exists
    FileFolderCheck Me.lblReportFolder, "folder", b
        If b = False Then
            MsgBox "Folder Selected is not valid"
            GoTo Exit_General
        End If
    FileFolderCheck Me.lblStaffReport, "file", b
        If b = False Then
            MsgBox "Staff Report Template is not valid"
            GoTo Exit_General
        End If
    FileFolderCheck Me.lblSRCompare, "file", b
        If b = False Then
            MsgBox "Previous Staff Report file is not valid"
            GoTo Exit_General
        End If
    'fill temp table with previous report info to compare to
    SRCompare Me.lblSRCompare

    StaffInitialize Screen.ActiveForm

    GoTo Exit_General

Exit_General:
'close objects
'delete temp table
    Set unitSelect = Nothing
    Set newReport = Nothing
    DoCmd.DeleteObject acTable, "tblCompareStaffReport"
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General
End Sub

Private Sub btnReportFolder_Click()
'uses windows FILE DIALOG to choose folder to save reports into
'saves selection in table for default value on next form load
DoCmd.SetWarnings False
On Error GoTo Err_General

Dim diaFile As FileDialog
Dim filePath As String          'selected folder

    Set diaFile = Application.FileDialog(msoFileDialogFolderPicker)
        With diaFile
            .AllowMultiSelect = False
            .Show
            If (.SelectedItems.Count = 0) Then
                MsgBox "Folder to Save Reports MUST be Selected"
            Else
                filePath = .SelectedItems(1)
            End If
        End With
'run SQL to update table with new selected folder
    DoCmd.RunSQL "UPDATE tblFilePicker SET Folder = '" & filePath & "'"
    Me.Refresh

Exit_General:
    Set diaFile = Nothing
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General
End Sub

Private Sub btnSRCompare_Click()
'uses windows FILE DIALOG to choose file of older STAFF REPORT
'saves selection in table for default value on next form load
DoCmd.SetWarnings False
On Error GoTo Err_General

Dim diaFile As FileDialog

    Set diaFile = Application.FileDialog(msoFileDialogFilePicker)
        With diaFile
            .Filters.Add "Excel Files", "*.xlsm"
            .AllowMultiSelect = False
            .Show
            If (.SelectedItems.Count = 0) Then
                MsgBox "Previous Report to Compare MUST be Selected"
            Else
                Forms!frmOffReports!lblSRCompare = .SelectedItems(1)
            End If
        End With

Exit_General:
    Set diaFile = Nothing
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General
End Sub

Private Sub btnFileStaffReport_Click()
'uses windows FILE DIALOG to choose file of STAFF REPORT template
'saves selection in table for default value on next form load
DoCmd.SetWarnings False
On Error GoTo Err_General

Dim diaFile As FileDialog
Dim filePath As String              'staff report template

    Set diaFile = Application.FileDialog(msoFileDialogFilePicker)
        With diaFile
            .Filters.Add "Excel Files", "*.xlsm"
            .AllowMultiSelect = False
            .Show
            If (.SelectedItems.Count = 0) Then
                MsgBox "Excel Staff Report Template MUST be Selected"
            Else
                filePath = .SelectedItems(1)
            End If
    End With
'run SQL to update table with new selected folder
    DoCmd.RunSQL "UPDATE tblFilePicker SET StaffReport = '" & filePath & "'"
    Me.Refresh

Exit_General:
    Set diaFile = Nothing
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    GoTo Exit_General
DoCmd.SetWarnings True
End Sub

Private Sub btnOpenUnits_Click()
    DoCmd.openForm "frmSelectUnits"
    Forms!frmSelectUnits.Move Left:=7500, Top:=250
    Forms!frmSelectUnits!cmdFolder.Visible = False
    Forms!frmSelectUnits!Label2.Visible = False
    Forms!frmSelectUnits!lblUnitFolder.Visible = False
    Forms!frmSelectUnits!Label4.Visible = False
    Forms!frmSelectUnits!startDate.Visible = False
    Forms!frmSelectUnits!Label31.Visible = False
    Forms!frmSelectUnits!endDate.Visible = False
    Me.lblResults = ""
    Me.lblResults.Visible = False
    'UnitsInitialize
End Sub

Private Sub btnNewSelection_Click()
    Me.lblResults = ""
    Me.lblResults.Visible = False
    btnOpenUnits_Click
End Sub

Private Sub btnCloseForm_Click()
'closes form
DoCmd.Close acForm, "frmOffReports"
End Sub

Public Property Get sheetName() As String
    sheetName = p_sheetName
End Property
Public Property Let sheetName(ByVal str As String)
    p_sheetName = str
End Property


Public Property Get fileName() As String
    fileName = p_fileName
End Property
Public Property Let fileName(ByVal str As String)
    p_fileName = str
End Property


Public Property Get filePath() As String
    filePath = p_filePath
End Property
Public Property Let filePath(ByVal str As String)
    p_filePath = str
End Property


Public Property Get objType() As String
    objType = p_objType
End Property
Public Property Let objType(ByVal str As String)
    p_objType = str
End Property


Public Property Get optUG() As Integer
    optUG = p_optUG
End Property
Public Property Let optUG(ByVal i As Integer)
    p_optUG = i
End Property


Public Property Get strIN() As String
    strIN = p_strIN
End Property
Public Property Let strIN(ByVal str As String)
    p_strIN = str
End Property


Public Property Get rsSQL() As String
    rsSQL = p_rsSQL
End Property
Public Property Let rsSQL(ByVal str As String)
    p_rsSQL = str
End Property


Public Property Get rsReport() As DAO.Recordset
    Set rsReport = p_rsReport
End Property
Public Property Set rsReport(rs As DAO.Recordset)
    Set p_rsReport = rs
End Property


Public Property Get xlBook() As Excel.Workbook
    Set xlBook = p_xlBook
End Property
Public Property Set xlBook(xl As Excel.Workbook)
    Set p_xlBook = xl
End Property
