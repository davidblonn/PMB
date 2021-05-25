Attribute VB_Name = "Excel_Worksheet"
Option Compare Database
Option Explicit

Public Sub MultiRangeColumn(xlSheet As Worksheet, cellVal As String, mthd As String)

Dim c As String
Dim v As String
Dim searchChar As String
Dim pos As Long

    searchChar = "!"
    pos = 1
    
    Do While pos <> 0
'find first comma in string at end of address
        pos = InStr(1, cellVal, searchChar, pos)
'if pos = 0, no more ! were found
        If pos <> 0 Then
'get cell address and remove !
            c = Left(cellVal, pos - 1)
'if " " precedes cell address, remove it
            If Left(c, 1) = " " Then
                c = Right(c, Len(c) - 1)
            End If
'remove cell address and comma from string
            cellVal = Right(cellVal, Len(cellVal) - pos)
'find first comma in string at end of value
            pos = InStr(1, cellVal, searchChar, pos)
            
            If pos <> 0 Then
'get value for cell and remove !
                v = Left(cellVal, pos - 1)
            Else
'pos = 0 if no ! at end of last value in string
'doesnt remove ! from value
                v = cellVal
            End If
'if " " precedes value, remove it
                If Left(v, 1) = " " Then
                    v = Right(v, Len(v) - 1)
                End If
'remove the value from the string
                cellVal = Right(cellVal, Len(cellVal) - pos)
        End If
'select case to call appropriate method to run
        Select Case mthd
            Case "values"
                xlSheet.range(c).Value = Nz(v, "")
            Case "formulas"
                xlSheet.range(c).Formula = v
            Case "merge"
                xlSheet.range(c, v).Merge
            Case "numberFormat"
                If Len(c) = 1 Then
                    xlSheet.Columns(c).NumberFormat = v
                Else
                    xlSheet.range(c).NumberFormat = v
                End If
            Case "width"
                If Len(c) = 1 Then
                    xlSheet.Columns(c).ColumnWidth = v
                Else
                    xlSheet.range(c).ColumnWidth = v
                End If
        End Select
    Loop

End Sub

Public Function TestIfOpen(ByVal wbName As String) As Boolean
On Error GoTo Err_General
Dim xlApp As Object
    Set xlApp = GetObject(, "Excel.Application")
    
Dim xlBook As Variant

    For Each xlBook In xlApp.Workbooks

        If xlBook.name = wbName Then
            TestIfOpen = True
            'MsgBox "File is open"
            Exit Function
        Else
            TestIfOpen = False
        End If
    Next
    Set xlApp = Nothing
Err_General:
    'MsgBox "No Excel Instances Open"
    Exit Function
End Function

Public Sub CloseExcel()
On Error GoTo Err_General
Dim xlApp As Excel.Application
    Set xlApp = GetObject(, "Excel.Application")


Dim wb As Workbook
'loop through any workbooks currently open in excel
    For Each wb In xlApp.Workbooks
        MsgBox "Close - " & wb.name & "?"
'closes workbooks
        wb.Close savechanges:=False
    Next
'exit application
    xlApp.quit
Err_General:
    'MsgBox "No Excel Instances Open"
    Exit Sub
End Sub


