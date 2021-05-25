VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNavigate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCloseExcel_Click()
'closes all open Excel applications
CloseExcel
End Sub

Private Sub Command10_Click()
DoCmd.openForm "frmOffReports"
End Sub

Private Sub Command11_Click()
DoCmd.openForm "frmLINSearch"
End Sub

Private Sub Command8_Click()
DoCmd.openForm "frmFilePicker"
End Sub

Private Sub Command9_Click()
DoCmd.openForm "frmActRegSearch"
End Sub

Private Sub btnExportCode_Click()
'exports all project code to text file on desktop
AllCodeToDesktop
End Sub
