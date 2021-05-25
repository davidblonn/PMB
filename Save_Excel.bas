Attribute VB_Name = "Save_Excel"
Option Compare Database
Option Explicit

Public Sub SaveExcel(xlBook As Excel.Workbook, reportType As String, _
            Optional newReport As Form_frmOffReports, _
            Optional regReport As Form_frmActRegSearch, _
            Optional newSheet As newSheet)
'saves excel workbooks based on parameter type
'concatenates file path, report name, dates, and file types
Dim xlApp As Excel.Application
Dim today As Date                   'date report ran
Dim newFileName As String           'holds concatenation
Dim sFileSaveName As Variant        'choose folder prompt
Dim unitName As String              'passed to AddUnit

Set xlApp = CreateObject("Excel.Application")
'get current date
today = Date
'select by report type
    Select Case reportType
'Staff Report - frmOffReport - Staff Report
        Case "staffReport"
            Set xlBook = xlApp.Workbooks.Open(newReport.fileName)
            newFileName = newReport.filePath & "\" & newReport.sheetName & "_" & _
                            Format(today, "ddmmmyyyy") & ".xlsm"
            xlBook.SaveAs fileName:=newFileName, _
                ConflictResolution:=xlLocalSessionChanges
            
'Run Report - frmOffReport - Incoming, Outgoing, Turn-In, Open Vetting
        Case "psdReport"
            unitName = ""           'initiate variable
'if OPTION GROUP is chosen
Debug.Print newReport.optUG
            If (newReport.optUG > 0) Then
'get OPTION GROUP label name
                unitName = GetOptionGroupName(newReport.optUG)
            Else
'pass to function to fill with user input
                AddUnit unitName
            End If
            newFileName = newReport.filePath & "\" & unitName & "_" & _
                            newReport.sheetName & _
                            "_" & Format(today, "ddmmmyyyy") & ".xlsx"
            xlBook.SaveAs fileName:=newFileName, _
                ConflictResolution:=xlLocalSessionChanges
            
'PMB Document Filing - frmActRegSearch - unit activity register spreadsheet
        Case "docFiling"
            xlBook.SaveAs fileName:=regReport.saveFolder, _
                ConflictResolution:=xlLocalSessionChanges
            
'File Search - frmActRegSearch - folder search for all documents
        Case "docSearch"
'prompts user for folder to save file in
            sFileSaveName = xlApp.GetSaveAsFilename("Found_DOCs_" & _
                            Format(regReport.sDate, "ddmmmyy") & "-" & _
                            Format(regReport.eDate, "ddmmmyy") & ".xlsx", _
                            "Excel Files (*.xlsx), *.xlsx")
            If sFileSaveName <> False Then
                xlBook.SaveAs sFileSaveName
            End If
            
'Unit Monthly Review - frmActRegSearch - rptUnitActReg, spreadsheet of missing documents
        Case "docMonthly"
            xlBook.SaveAs fileName:=regReport.saveFolder & "_MissingDOCs", _
            ConflictResolution:=xlLocalSessionChanges
        
        Case "linkedSheet"
            xlBook.SaveAs fileName:=newSheet.sourceFile, _
                ConflictResolution:=xlLocalSessionChanges
                With newSheet
                    .xlBook.Close
                    .xlApp.quit
                End With
                Set newSheet = Nothing
                Set xlBook = Nothing
                Set xlApp = Nothing
            Exit Sub
        End Select

xlBook.Close
xlApp.quit

Set xlBook = Nothing
Set xlApp = Nothing

End Sub

Public Function AddUnit(unitName As String) As String
'adds user prompted text to file save path for custom file naming
Dim msg, default As String
    msg = "Add Unit Name to file"
    default = ""
    unitName = InputBox(msg, , default)
    AddUnit = unitName

End Function

Public Function GetFilenameFromPath(ByVal fileName As String) As String

If Right(fileName, 1) <> "\" And Len(fileName) > 0 Then
    GetFilenameFromPath = _
    GetFilenameFromPath(Left(fileName, Len(fileName) - 1)) + Right(fileName, 1)
End If
End Function

Public Function GetOptionGroupName(i As Integer) As String
'returns label name per OPTION GROUP value
    Select Case i
        Case 1
            GetOptionGroupName = "SFG"
        Case 2
            GetOptionGroupName = "ENG"
        Case 3
            GetOptionGroupName = "MPs"
        Case 4
            GetOptionGroupName = "AVN"
        Case 5
            GetOptionGroupName = "201st"
        Case 6
            GetOptionGroupName = "150th"
        Case 7
            GetOptionGroupName = "771st"
        Case 8
            GetOptionGroupName = "77thBDE"
        Case Else
            GetOptionGroupName = "TDA"
    End Select

End Function
