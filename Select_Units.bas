Attribute VB_Name = "Select_Units"
Option Compare Database
Option Explicit

Public unitSelect As Form_frmSelectUnits
Private rsResult As DAO.Recordset
Private curForm As Form

Public Sub UnitsInitialize(openForm As Form)

    Set curForm = openForm

Dim e As Integer
Dim uics As String
Dim b As Boolean

    e = curForm.lstUnits.ItemsSelected.Count
    Set unitSelect = New Form_frmSelectUnits
    b = True

    If IsNull(curForm.optUnitGroups) And _
            (e = 0) And _
            IsNull(curForm.lblUnitFolder) Then
        MsgBox "No Units Selected"
        Exit Sub
    End If
    
    With unitSelect
        .d = False
        If Not IsNull(curForm.startDate) Then
            .sDate = curForm.startDate
            .d = True
        End If
        .eDate = curForm.endDate
        If Not IsNull(curForm.optUnitGroups) Then
            .optUG = curForm.optUnitGroups
        End If
        Set .rsCompare = RecordsetComparison
    tempRsResults
    Set rsResult = CurrentDb.OpenRecordset("tempRsResults", dbOpenDynaset)
        '.sql = ListBoxSelections
        If Not IsNull(curForm.lblUnitFolder) Then
            .searchFolder = curForm.lblUnitFolder
        End If
        
        If (.optUG > 0) And (e > 0) Then
            If MsgBox("Too many selections. Do you want to keep ListBox?", _
                vbYesNo + vbQuestion) = vbYes Then
                    .optUG = 0
                    .sql = ListBoxSelections(b)
            Else
                    .sql = ""
                    DeselectAll curForm
                    UnitGroups
                    .sql = ListBoxSelections
            End If
        ElseIf (.optUG > 0) And (e = 0) Then
                    UnitGroups
                    .sql = ListBoxSelections
        ElseIf (.optUG = 0) And (e > 0) Then
                    .sql = ListBoxSelections(b)
        End If
        
        If Len(.searchFolder) > 3 Then
            .optUG = 0
            .sql = ""
            UICGroups
            .sql = ListBoxSelections
        End If
    
            rsResult.MoveFirst
                Do Until rsResult.EOF
                    uics = uics & rsResult!UIC & _
                        " - " & rsResult!Unit & vbCrLf
                    rsResult.MoveNext
                Loop
        
            .results = uics
        
        'Debug.Print .sDate & " - " & .eDate & " - " & .optUG
        'Debug.Print .sql
        'Debug.Print "search folder - " & .searchFolder
    End With

    rsResult.Close

End Sub

Public Function RecordsetComparison() As DAO.Recordset

Dim sql As String
    sql = "SELECT UIC, Unit FROM Units ORDER BY Unit"
    Set RecordsetComparison = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

End Function

Public Function ListBoxSelections(Optional b As Boolean) As String

Dim uics As String
Dim i As Integer
    
    rsResult.MoveFirst

    For i = 0 To curForm.lstUnits.ListCount - 1
        If curForm.lstUnits.Selected(i) Then
            uics = uics & "'" & curForm.lstUnits.Column(0, i) & "',"
            rsResult.MoveNext
        Else
            If b = True Then
                rsResult.Delete
                rsResult.MoveNext
            End If
        End If
    Next i
    
    ListBoxSelections = Left(uics, Len(uics) - 1)
    Debug.Print ListBoxSelections

End Function

Public Sub tempRsResults()
DoCmd.SetWarnings False

If Not IsNull(DLookup("Name", "MSysObjects", "Name='tempRsResults'")) Then
    DoCmd.DeleteObject acTable, "tempRsResults"
End If

DoCmd.RunSQL "SELECT UIC, Unit INTO tempRsResults FROM Units ORDER BY Unit;"
curForm.Requery
DoCmd.SetWarnings True

End Sub

Public Sub UICGroups()
'makes LISTBOX selections based on unit FOLDER NAME

Dim sql As String
Dim i As Integer
Dim tableUIC As String
Dim folderUIC As String
Dim bnUIC As String

'same SQL data source as form LISTBOX
    'recordset clone of unit LISTBOX
    With unitSelect
        folderUIC = Right(.searchFolder, 11)                    'remove folder full path
        Debug.Print folderUIC
        bnUIC = Left(folderUIC, 4)                              'unit string to match
        Debug.Print bnUIC
        i = 0
    
        .rsCompare.MoveFirst
        rsResult.MoveFirst
            Do Until .rsCompare.EOF
                tableUIC = Left(.rsCompare!UIC, 4)              'recordset value to match
                Debug.Print tableUIC
                    If bnUIC = tableUIC Then                    'current recordset = unit
                        curForm.lstUnits.Selected(i) = True     'select unit
                    Else
                        rsResult.Delete
                    End If
                i = i + 1
                .rsCompare.MoveNext
                rsResult.MoveNext
            Loop
    End With
End Sub

Public Sub UnitGroups()
'makes LISTBOX selections based on OPTION GROUP selection

Dim i As Integer

'same SQL data source as form LISTBOX
   'recordset clone of unit LISTBOX
    i = 0
    
    With unitSelect
        .rsCompare.MoveFirst
        rsResult.MoveFirst
        
        Select Case .optUG
'based on OPTION GROUP selection
'loops through recordset of unit names that matches LISTBOX
'uses string based on OPTION GROUP selection to match to recordset unit names
            Case 1
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 4) = "19 S" Then
                            curForm.lstUnits.Selected(i) = True             'select if true
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
'duplicate procedures for all cases
            Case 2
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 3) = "ENG" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
            Case 3
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 2) = "MP" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
            Case 4
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 3) = "AVN" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
            Case 5
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 3) = "201" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
            Case 6
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 3) = "1-1" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
            Case 7
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 3) = "771" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
            Case 8
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 3) = "77T" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
            Case Else
'last case, for TDA selection, is all other units
                Do Until .rsCompare.EOF
                        If Left(.rsCompare!Unit, 3) <> "77T" And _
                            Left(.rsCompare!Unit, 3) <> "771" And _
                            Left(.rsCompare!Unit, 3) <> "1-1" And _
                            Left(.rsCompare!Unit, 3) <> "201" And _
                            Left(.rsCompare!Unit, 3) <> "AVN" And _
                            Left(.rsCompare!Unit, 2) <> "MP" And _
                            Left(.rsCompare!Unit, 3) <> "ENG" And _
                            Left(.rsCompare!Unit, 4) <> "19 S" Then
                            curForm.lstUnits.Selected(i) = True
                        Else
                            rsResult.Delete
                        End If
                    i = i + 1
                    .rsCompare.MoveNext
                    rsResult.MoveNext
                Loop
        End Select
    End With

End Sub

Public Sub DeselectAll(openForm As Form)

    Set curForm = openForm
    Dim s As Integer
    With curForm.lstUnits
        For s = 0 To .ListCount - 1
            .Selected(s) = False
        Next
    End With
    curForm.optUnitGroups = 0

End Sub


Public Sub SelectAll(openForm As Form)

    Set curForm = openForm
    Dim s As Integer
    With curForm.lstUnits
        For s = 0 To .ListCount - 1
            .Selected(s) = True
        Next
    End With
    curForm.optUnitGroups = 0

End Sub
