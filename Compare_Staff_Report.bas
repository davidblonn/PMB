Attribute VB_Name = "Compare_Staff_Report"
Option Compare Database
Option Explicit
'uses previous staff report to create a temp table
'creates recordset from temp table and pastes values into new staff report
'parameter is file path to old staff report
Public Sub SRCompare(fSRCName As String)

On Error GoTo Err_General
DoCmd.SetWarnings False

Dim rs As DAO.Recordset
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
Set xlBook = xlApp.Workbooks.Open(fSRCName)
Set xlSheet = xlBook.Worksheets("Staff_Report")

'delete temp table if it exists
If Not IsNull(DLookup("Name", "MSysObjects", "Name='tblCompareStaffReport'")) Then
    DoCmd.DeleteObject acTable, "tblCompareStaffReport"
End If

Dim tdf As DAO.TableDef

'create table name and field definitions
Set tdf = CurrentDb.CreateTableDef("tblCompareStaffReport")
tdf.Fields.Append tdf.CreateField("Date", dbDate)
tdf.Fields.Append tdf.CreateField("InStateTotal", dbLong)
tdf.Fields.Append tdf.CreateField("InState$", dbLong)
tdf.Fields.Append tdf.CreateField("OutgoingTotal", dbLong)
tdf.Fields.Append tdf.CreateField("Outgoings$", dbLong)
tdf.Fields.Append tdf.CreateField("TITotal", dbLong)
tdf.Fields.Append tdf.CreateField("TI$", dbLong)
tdf.Fields.Append tdf.CreateField("OpenVetTotal", dbLong)
tdf.Fields.Append tdf.CreateField("OpenVet$", dbLong)
tdf.Fields.Append tdf.CreateField("IncomingTotal", dbLong)
tdf.Fields.Append tdf.CreateField("Incoming$", dbLong)
tdf.Fields.Append tdf.CreateField("54Total", dbLong)
tdf.Fields.Append tdf.CreateField("54Vehicles", dbLong)
tdf.Fields.Append tdf.CreateField("54$", dbLong)

'create the temp table
CurrentDb.TableDefs.Append tdf
'create recordset
Set rs = CurrentDb.OpenRecordset("tblCompareStaffReport", dbOpenTable)
rs.AddNew
'fill recordset with values from old excel report
rs.Fields("Date") = xlSheet.range("L1").Value
rs.Fields("InStateTotal") = xlSheet.range("B1").Value
rs.Fields("InState$") = xlSheet.range("B5").Value
rs.Fields("OutgoingTotal") = xlSheet.range("B6").Value
rs.Fields("Outgoings$") = xlSheet.range("B10").Value
rs.Fields("TITotal") = xlSheet.range("B11").Value
rs.Fields("TI$") = xlSheet.range("B15").Value
rs.Fields("OpenVetTotal") = xlSheet.range("B16").Value
rs.Fields("OpenVet$") = xlSheet.range("B20").Value
rs.Fields("IncomingTotal") = xlSheet.range("B24").Value
rs.Fields("Incoming$") = xlSheet.range("B28").Value
rs.Fields("54Total") = xlSheet.range("B21").Value
rs.Fields("54Vehicles") = xlSheet.range("B23").Value
rs.Fields("54$") = xlSheet.range("B22").Value
'fill temp table with recordset data
rs.Update

GoTo Exit_General

Exit_General:
'close objects
'save old report
    Set tdf = Nothing
    rs.Close
    Set rs = Nothing
    xlBook.Save
    xlBook.Close True
    xlApp.quit
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlApp = Nothing
    DoCmd.SetWarnings True
    Exit Sub
Err_General:
    MsgBox Err.Number & Err.Description
    Resume Exit_General
End Sub
