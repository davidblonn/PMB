Attribute VB_Name = "Activity_Register_Reports"
Option Compare Database
Option Explicit

Public regReport As Form_frmActRegSearch
Public curForm As Form
Private db As DAO.Database
Private FSO As New FileSystemObject

Dim pdfSave As String

Public Sub ActRegInitialize(openForm As Form, fType As String)

    Set db = CurrentDb
    Set curForm = openForm
    Set regReport = New Form_frmActRegSearch
    
    With regReport
        .fType = fType
        If CurrentProject.AllForms("frmSelectUnits").IsLoaded Then
            If Not IsNull(unitSelect.optUG) Then
                .optUG = unitSelect.optUG
            End If
        End If
    
        Select Case .fType
            
            Case "docFiling"
                .myFolder = unitSelect.searchFolder
                .saveFolder = BuildFileName
                CreateFoundDOCsTable
                RecFolderSearch (unitSelect.searchFolder)
                .sql = ActRegSQL
                Set .rsActReg = db.OpenRecordset(.sql, dbOpenDynaset)
                    If .rsActReg.recordCount = 0 Then
                        MsgBox "No Document Activity"
                        Exit Sub
                    End If
                Set .rsFoundDOCs = db.OpenRecordset("tempFoundDOCs", dbOpenDynaset)
                Set .xlBook = ExcelReport
                SaveExcel regReport.xlBook, regReport.fType, , regReport
                
            Case "docMonthly"
                .myFolder = curForm.lblSaveFolder
                .saveFolder = BuildFileName
                .sql = ActRegSQL
                DoCmd.RunSQL .sql
                Set .rsActReg = CurrentDb.OpenRecordset("SELECT * " & _
                    "FROM tblUnitActReg " & _
                    "WHERE (FirstOfReceived <> -1) " & _
                    "AND (FirstOfInfeasible <> -1)", dbOpenDynaset)
                    If .rsActReg.recordCount = 0 Then
                        MsgBox "No Missing Documents"
                    End If
                Forms!frmActRegSearch!startDate = unitSelect.sDate
                Forms!frmActRegSearch!endDate = unitSelect.eDate
                Set .xlBook = ExcelReport
                SaveExcel regReport.xlBook, regReport.fType, , regReport
                pdfSave = regReport.saveFolder & ".pdf"
                DoCmd.OutputTo acOutputReport, "rptUnitActReg", acFormatPDF, pdfSave
                
            Case "docSearch"
                .myFolder = curForm.lblFolder
                .saveFolder = BuildFileName
                CreateFoundDOCsTable
                RecFolderSearch (curForm.lblFolder)
                Set .rsFoundDOCs = db.OpenRecordset("tempFoundDOCs", dbOpenDynaset)
                Set .xlBook = ExcelReport
                SaveExcel regReport.xlBook, regReport.fType, , regReport
                
            Case "docNumSearch", "linSearch"
                .strIN = curForm.lblNumSearch
                .sql = ActRegSQL
                Set .rsFoundDOCs = db.OpenRecordset(.sql, dbOpenDynaset)
                        If .rsFoundDOCs.recordCount = 0 Then
                            MsgBox "No Document Number data found"
                            Exit Sub
                        End If
                Set .xlBook = ExcelReport
                GoTo Exit_General
        End Select

    End With
    GoTo Exit_General
        
Exit_General:
    Set FSO = Nothing
    Set curForm = Nothing
    Set db = Nothing
    Exit Sub

End Sub

Public Function ActRegSQL() As String
'assigns SELECTED UNITS and START and END DATES to SQL statement
Dim uics As String

    If (regReport.fType = "docNumSearch") Or _
        (regReport.fType = "linSearch") Then
        uics = regReport.strIN
    Else
        uics = unitSelect.sql
    End If
    
Select Case regReport.fType
    Case "docFiling"
'all activity register records for units within date range
'***Left(p_strIN, Len(p_strIN) - 1) & ")*** removes comma from end of list of units
    ActRegSQL = "SELECT * " & _
            "FROM tblActRegAppend " & _
            "WHERE tblActRegAppend.UIC IN (" & uics & ") " & _
            "AND DateValue('" & unitSelect.sDate & "') <= tblActRegAppend.LastUpdate " & _
            "AND DateValue('" & unitSelect.eDate & "') >= tblActRegAppend.LastUpdate " & _
            "ORDER BY tblActRegAppend.LastUpdate DESC"

    Case "docMonthly"
'CREATE TABLE of grouped document numbers for units within date range
'table is used to generate Access report
    ActRegSQL = "SELECT First(tblActRegAppend.UIC) AS FirstOfUIC, " & _
        "tblActRegAppend.DOC_NUM, " & _
        "First(tblActRegAppend.LIN) AS FirstOfLIN, " & _
        "First(tblActRegAppend.NOMEN) AS FirstOfNOMEN, " & _
        "First(tblActRegAppend.MVMT_TYPE) AS FirstOfMVMT_TYPE, " & _
        "First(tblActRegAppend.MVMT_Text) AS FirstOfMVMT_Text, " & _
        "First(tblActRegAppend.FORM_NUM) AS FirstOfFORM_NUM, " & _
        "First(tblActRegAppend.LastUpdate) AS FirstOfLastUpdate, " & _
        "First(tblActRegAppend.Received) AS FirstOfReceived, " & _
        "First(tblActRegAppend.Infeasible) AS FirstOfInfeasible, " & _
        "Units.Unit, " & _
        "Units.Sloc " & _
        "INTO tblUnitActReg " & _
        "FROM (tblActRegAppend " & _
        "INNER JOIN Units " & _
        "ON tblActRegAppend.UIC = Units.UIC) " & _
        "WHERE Units.UIC " & _
        "IN (" & uics & ") " & _
        "AND tblActRegAppend.LastUpdate >= DateValue('" & unitSelect.sDate & "') " & _
        "AND tblActRegAppend.LastUpdate <= DateValue('" & unitSelect.eDate & "') " & _
        "GROUP BY DOC_NUM, tblActRegAppend.UIC, LIN, NOMEN, " & _
        "MVMT_TYPE, FORM_NUM, LastUpdate, Received, " & _
        "Infeasible, Units.Unit, Units.Sloc;"
        
    Case "docNumSearch"
'all records from table matching document number input
    ActRegSQL = "SELECT * " & _
            "FROM tblActRegAppend " & _
            "WHERE (DOC_NUM = '" & uics & "') " & _
            "ORDER BY LastUpdate"
            
    Case "linSearch"
'all records from table matching LIN input
    ActRegSQL = "SELECT * " & _
            "FROM tblActRegAppend " & _
            "WHERE (LIN = '" & uics & "') " & _
            "ORDER BY LastUpdate"
            
    Debug.Print ActRegSQL
            
End Select

End Function

Public Sub RecFolderSearch(myFolder As String)
'recursive folder search starting with ROOT FOLDER
DoCmd.SetWarnings False

Dim topFolder As Folder                 'ROOT FOLDER
Dim subFolder As Folder                 'folders in ROOT FOLDER
Dim myFile As file                      'files in all folders

'set parameter to ROOT FOLDER
    Set topFolder = FSO.GetFolder(myFolder)
'loop through all folders in ROOT FOLDER
            For Each subFolder In topFolder.SubFolders
                On Error Resume Next
'loop through all files in currently selected folder
                    For Each myFile In subFolder.Files
                        On Error Resume Next
                            RegExFileSearch myFile  'matches file name to regular expression
                    Next
'assign argument to currently selected folder
                myFolder = subFolder.path
'recursive call to function
                RecFolderSearch myFolder
            Next
'loop through all files in original parameter folder
        For Each myFile In topFolder.Files
            On Error Resume Next
                RegExFileSearch myFile      'matches file name to regular expression
        Next
'UPDATES record of all found documents in temp table to tblActRegAppend
    db.Execute "UPDATE tblActRegAppend " & _
                "INNER JOIN tempFoundDOCs " & _
                "ON tblActRegAppend.DOC_NUM = tempFoundDOCs.FDOCNUM " & _
                "SET tblActRegAppend.Received = TRUE"
    
DoCmd.SetWarnings True

End Sub

Public Sub RegExFileSearch(myFile As file)
'used during recursive folder search to match file names to
'regular expressions that contain standard document number patterns and
'within START and END DATES range
'APPENDS matched file name records to temporary table
DoCmd.SetWarnings False

Dim sql As String               'INSERT into SQL for matching numbers
Dim f As String                 'file name
Dim d As Date                   'file date
Dim p As String                 'file path
Dim strPattern As String        'regular expression
Dim regEx As New RegExp
Dim fullDocNum As String        'concatenated document number

        f = FSO.GetFileName(myFile.name)
        d = FileDateTime(myFile)
        p = myFile.path
'validate file was created within date range
        If d < unitSelect.eDate And d > unitSelect.sDate Then
        
        With regEx
            .Global = True
            '.Multiline = True
            .IgnoreCase = True
        End With
'regular expression to find 2 document numbers in file name
            strPattern = ".*([W]{1}[a-zA-Z0-9]{5}).*([0-9]{4}).*([0-9]{4}).*" & _
                        "([W]{1}[a-zA-Z0-9]{5}).*([0-9]{4}).*([0-9]{4}).*"
'assign string to regular expression object
            regEx.pattern = strPattern
            
            If regEx.test(f) Then
'concatenate 14 char string and run APPEND SQL on 1st found document number
                fullDocNum = regEx.Replace(f, "$1" & "$2" & "$3")
                sql = "INSERT INTO tempFoundDOCs (FilePath, FDOCNUM) " & _
                        "VALUES ('" & p & "', '" & fullDocNum & "');"
                DoCmd.RunSQL sql
'concatenate 14 char string and run APPEND SQL on 2nd found document number
                fullDocNum = regEx.Replace(f, "$4" & "$5" & "$6")
                sql = "INSERT INTO tempFoundDOCs (FilePath, FDOCNUM) " & _
                        "VALUES ('" & p & "', '" & fullDocNum & "');"
                DoCmd.RunSQL sql

            ElseIf Not regEx.test(f) Then
'regular expression to find only 1 document number in file name
                strPattern = ".*([W]{1}[a-zA-Z0-9]{5}).*([0-9]{4}).*([0-9]{4}).*"
'assign string to regular expression object
                regEx.pattern = strPattern

                If regEx.test(f) Then
'concatenate 14 char string and run APPEND SQL on found document number
                    fullDocNum = regEx.Replace(f, "$1" & "$2" & "$3")
                    sql = "INSERT INTO tempFoundDOCs (FilePath, FDOCNUM) " & _
                            "VALUES ('" & p & "', '" & fullDocNum & "');"
                    DoCmd.RunSQL sql
                End If
            End If
        End If

End Sub

Public Sub CreateFoundDOCsTable()
'creates temporary table to APPEND document records to when searching through file names
'delete table if already exists
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempFoundDOCs'")) Then
        DoCmd.DeleteObject acTable, "tempFoundDOCs"
    End If

Dim tdf As DAO.TableDef             'table object
'assign table and field names
    Set tdf = CurrentDb.CreateTableDef("tempFoundDOCs")
    tdf.Fields.Append tdf.CreateField("FilePath", dbText)
    tdf.Fields.Append tdf.CreateField("FDOCNUM", dbText)
'create temp table
    CurrentDb.TableDefs.Append tdf

End Sub

Public Function BuildFileName() As String
'uses the unit folder name to create a new file name when saving reports
Select Case regReport.fType
'selection based on assigned report type
    Case "docFiling"                        'PMB Document Filing
        Dim unitFile As String
        Dim oneUIC As String
        Dim twoUIC As String
        Dim arYear As String
        
            oneUIC = Right(regReport.myFolder, 11)  'UIC and year from folder path
            twoUIC = Left(oneUIC, 6)        'UIC
            arYear = Right(oneUIC, 4)       'year
'concatenate for ACTIVITY REGISTER document name
            unitFile = regReport.myFolder & _
                "\Activity Register_" & arYear & "_" & twoUIC
        
            BuildFileName = unitFile
    Case "docMonthly"                       'Monthly Review Reports
'assigns OPTION GROUP name to file name
            Select Case regReport.optUG
                Case 1
                    unitFile = "SFG"
                Case 2
                    unitFile = "ENG"
                Case 3
                    unitFile = "MPs"
                Case 4
                    unitFile = "AVN"
                Case 5
                    unitFile = "201st"
                Case 6
                    unitFile = "150th"
                Case 7
                    unitFile = "771st"
                Case 8
                    unitFile = "77thBDE"
                Case 9
                    unitFile = "TDA"
                Case 0                  'prompt user to input file name
                    AddUnit (unitFile)
            End Select
'concatenate for report document name
        regReport.myFolder = regReport.myFolder & "\" & unitFile & "_" & _
                    Format(unitSelect.eDate, "mmmyy")
    
        BuildFileName = regReport.myFolder
    Case Else
        BuildFileName = regReport.myFolder
End Select

End Function

Public Function ExcelReport() As Excel.Workbook
'create Excel spredsheet based on report type and copy in recordset information
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim e As Long
Dim rangeStr As String

    Set xlApp = Excel.Application
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

        With xlSheet
            Select Case regReport.fType
                Case "docFiling"                'PMB Document Filing
            
                    rangeStr = "B1!700A Property Book/Hand Receipt! B2!Supporting Change Documents    " & _
                        unitSelect.sDate & " to " & unitSelect.eDate & "!" & _
                        "B3!PA:NA Destroy in CFA when 6 years old! B4!SUSPENSE! " & _
                        "B5!Created:   " & Date & "!" & _
                        "G2!The following documents have been pulled from GCSS-Army! " & _
                        "G3!Supporting Change Documentation must be filed digitally! " & _
                        "G4!within this same folder to support Property Book actions! " & _
                        "A7!UIC! !B7!Document Number! C7!NSN! D7!LIN! " & _
                        "E7!Nomenclature! F7!DOC_ID_CD! G7!UI! H7!UIC_GAIN! " & _
                        "I7!GAIN_LOSE! J7!MVMT_TYPE! K7!MVMT_Description! " & _
                        "L7!Date Close! M7!Trans Date! N7!Serial Number! O7!Last Update! " & _
                        "P7!Form Number! Q7!Received"
                    MultiRangeColumn xlSheet, rangeStr, "values"
                    
                    rangeStr = "O!mm/dd/yyyy! N!@! C!@"
                    MultiRangeColumn xlSheet, rangeStr, "numberFormat"
                    
                e = 8
                regReport.rsActReg.MoveFirst                    'recordset of all activity register documents
                
                Do While Not regReport.rsActReg.EOF
                    With regReport.rsActReg
                        rangeStr = "A" & e & "!" & Nz(!UIC, "") & "!" & _
                            "B" & e & "!" & Nz(!DOC_NUM, "") & "!" & _
                            "C" & e & "!" & Nz(!NSN, "") & "!" & _
                            "D" & e & "!" & Nz(!LIN, "") & "!" & _
                            "E" & e & "!" & Nz(!NOMEN, "") & "!" & _
                            "F" & e & "!" & Nz(!DOC_ID_CD, "") & "!" & _
                            "G" & e & "!" & Nz(!UI, "") & "!" & _
                            "H" & e & "!" & Nz(!UIC_GAIN, "") & "!" & _
                            "I" & e & "!" & Nz(!GAIN_LOSE_, "") & "!" & _
                            "J" & e & "!" & Nz(!MVMT_TYPE, "") & "!" & _
                            "L" & e & "!" & Nz(!MVMT_Text, "") & "!" & _
                            "N" & e & "!" & Nz(!SERIAL_NUM, "") & "!" & _
                            "O" & e & "!" & Nz(!LastUpdate, "") & "!" & _
                            "P" & e & "!" & Nz(!FORM_NUM, "") & "!" & _
                            "Q" & e & "!" & Nz(!Received, "")
                    End With
                    MultiRangeColumn xlSheet, rangeStr, "values"
        
                        If regReport.rsFoundDOCs.recordCount <> 0 Then 'recordset of found documents
                            regReport.rsFoundDOCs.MoveFirst
                            Do While Not regReport.rsFoundDOCs.EOF
'loops through recordset to find any matches in document numbers
                                If regReport.rsFoundDOCs!FDOCNUM = .Cells(e, 2) Then
'creates hyperlink in excel spreadsheet to document
                                    .Hyperlinks.Add Cells(e, 2), regReport.rsFoundDOCs!filePath
                                End If
                                regReport.rsFoundDOCs.MoveNext
                            Loop
                        End If
                        e = e + 1
                        regReport.rsActReg.MoveNext
                Loop
                
                    .range("A7:P7").AutoFilter
                    .range("A1:P1").Columns.AutoFit
                    .range("B1:B4").Font.Bold = True
                    .range("G2:G4").Font.Italic = True
                    .range("A1:G5").Locked = True

                    rangeStr = "A:A,D:D!8! B!19! O!15! E!12! C!10!"
                    MultiRangeColumn xlSheet, rangeStr, "width"
'make a copy of sheet 1
                Worksheets("Sheet1").Copy After:=Worksheets("Sheet1")
'filter second sheet to just unique document number rows
                            With xlBook.Worksheets("Sheet1 (2)")
                                .Select
                                .range("B:B").AdvancedFilter Action:=xlFilterInPlace, Unique:=True
                                .name = "Doc_Numbers"
                            End With
                
                xlSheet.name = "Complete_Register"
        
            Case "docSearch"                      'All Document Search
        
                    .range("B1").Value = "Document Number"
                    .range("D1").Value = "Path"
                    
                    regReport.rsFoundDOCs.MoveFirst
                    e = 2
'copy all recordset data in excel speadsheet
                        Do While Not regReport.rsFoundDOCs.EOF
                            .range("B" & e).Value = Nz(regReport.rsFoundDOCs!FDOCNUM, "")
                            .range("D" & e).Value = Nz(regReport.rsFoundDOCs!filePath, "")
'create hyperlink to document
                            .Hyperlinks.Add Cells(e, 2), regReport.rsFoundDOCs!filePath
                            e = e + 1
                            regReport.rsFoundDOCs.MoveNext
                        Loop
                    
                    .range("B1").AutoFilter
                    .range("A1:D1").Columns.AutoFit
                    .Columns("D").Cells.HorizontalAlignment = xlHAlignRight
                        
                xlSheet.name = "Found Documents"
              
            Case "docMonthly"                   'Monthly Review Reports
        
                rangeStr = "B1!Register of Documents NOT Received " & _
                    unitSelect.sDate & " to " & unitSelect.eDate & "!" & _
                    "B3!UNIT! C3!UIC! D3!DOC_NUM! E3!LIN! F3!NOMEN! G3!MVMT_TYPE! " & _
                    "H3!MVMT_Description! I3!LAST UPDATE! J3!FORM_NUM! K3!SLOC! L3!Remarks"
                MultiRangeColumn xlSheet, rangeStr, "values"
                
                e = 4
                Do While Not regReport.rsActReg.EOF            'recordset of all activity register documents
                    With regReport.rsActReg
                        rangeStr = "B" & e & "!" & Nz(!Unit, "") & "!" & _
                            "C" & e & "!" & Nz(!FirstOfUIC, "") & "!" & _
                            "D" & e & "!" & Nz(!DOC_NUM, "") & "!" & _
                            "E" & e & "!" & Nz(!FirstOfLIN, "") & "!" & _
                            "F" & e & "!" & Nz(!FirstOfNOMEN, "") & "!" & _
                            "G" & e & "!" & Nz(!FirstOfMVMT_TYPE, "") & "!" & _
                            "H" & e & "!" & Nz(!FirstOfMVMT_Text, "") & "!" & _
                            "I" & e & "!" & Nz(!FirstOfLastUpdate, "") & "!" & _
                            "J" & e & "!" & Nz(!FirstOfFORM_NUM, "") & "!" & _
                            "K" & e & "!" & Nz(!Sloc, "")
                    End With
                    MultiRangeColumn xlSheet, rangeStr, "values"
                    e = e + 1
                    regReport.rsActReg.MoveNext
                Loop
                
                .Columns("I").NumberFormat = "mmm-dd-yy"
                .range("B3:K3").AutoFilter
                .Cells.EntireColumn.AutoFit
                
            Case "docNumSearch", "linSearch"    'Document Number & LIN Search
        
                rangeStr = "A1!UIC! B1!DOC_NUM! C1!NSN! D1!LIN! E1!NOMEN! F1!DOC_ID_CD! " & _
                    "G1!UI! H1!UIC_GAIN! I1!GAIN_LOSE! J1!MVMT_TYPE! K1!MVMT_Description! " & _
                    "L1!DATE_CLOSE! M1!SERIAL_NUM! N1!LAST UPDATE! O1!FORM_NUM! " & _
                    "P1!RECEIVED! Q1!INFEASIBLE"
                MultiRangeColumn xlSheet, rangeStr, "values"
                
                e = 2
                
                rangeStr = "C:C,M:M!@"
                MultiRangeColumn xlSheet, rangeStr, "numberFormat"
                
                Do While Not regReport.rsFoundDOCs.EOF
                    With regReport.rsFoundDOCs
                        rangeStr = "A" & e & "!" & Nz(!UIC, "") & "!" & _
                            "B" & e & "!" & Nz(!DOC_NUM, "") & "!" & _
                            "C" & e & "!" & Nz(!NSN, "") & "!" & _
                            "D" & e & "!" & Nz(!LIN, "") & "!" & _
                            "E" & e & "!" & Nz(!NOMEN, "") & "!" & _
                            "F" & e & "!" & Nz(!DOC_ID_CD, "") & "!" & _
                            "G" & e & "!" & Nz(!UI, "") & "!" & _
                            "H" & e & "!" & Nz(!UIC_GAIN, "") & "!" & _
                            "I" & e & "!" & Nz(!GAIN_LOSE_, "") & "!" & _
                            "J" & e & "!" & Nz(!MVMT_TYPE, "") & "!" & _
                            "K" & e & "!" & Nz(!MVMT_Text, "") & "!" & _
                            "L" & e & "!" & Nz(!DATE_CLOSE, "") & "!" & _
                            "M" & e & "!" & Nz(!SERIAL_NUM, "") & "!" & _
                            "N" & e & "!" & Nz(!LastUpdate, "") & "!" & _
                            "O" & e & "!" & Nz(!FORM_NUM, "") & "!" & _
                            "P" & e & "!" & Nz(!Received, "") & "!" & _
                            "Q" & e & "!" & Nz(!Infeasible, "")
                    End With
                    MultiRangeColumn xlSheet, rangeStr, "values"
                    e = e + 1
                    regReport.rsFoundDOCs.MoveNext
                Loop
            .range("A1:Q1").AutoFilter
            .Cells.EntireColumn.AutoFit
            
        End Select
        
    End With
    
    Set ExcelReport = xlBook

    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
End Function
