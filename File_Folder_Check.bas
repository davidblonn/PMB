Attribute VB_Name = "File_Folder_Check"
Option Compare Database
Option Explicit

Public Function FileFolderCheck(fileFolder As String, fType As String, _
                                b As Boolean, Optional failName As String) As Boolean
'based on parameter, checks for file or folder existing in windows system
'returns false if does not exist
Dim FSO As FileSystemObject
Set FSO = CreateObject("Scripting.FileSystemObject")

If fType = "folder" Then
    If Not FSO.FolderExists(fileFolder) Then
        b = False
        MsgBox failName & "folder is not valid"
    End If
End If

If fType = "file" Then
    If Not FSO.FileExists(fileFolder) Then
        b = False
        MsgBox failName & "file is not valid"
    End If
End If

FileFolderCheck = b

Set FSO = Nothing

End Function

Public Sub AllCodeToDesktop()
'creates a text file and transfers all vba project objects as strings to it
   Dim fs As Object             'scripting file system object
   Dim f As Object              'text file
   Dim strMod As String         'code
   Dim mdl As Object            'iterator object
   Dim i As Integer             'code line count
   
   Set fs = CreateObject("Scripting.FileSystemObject")

        Set f = fs.CreateTextFile("C:\Users\david.a.blonn\Desktop" & "\" _
       & Replace(CurrentProject.name, ".", "") & ".txt")

'for each object in the project
        For Each mdl In VBE.ActiveVBProject.VBComponents
'uses a count of lines
            i = VBE.ActiveVBProject.VBComponents(mdl.name).CodeModule.CountOfLines
'put the code in a string
            If i > 0 Then
                strMod = VBE.ActiveVBProject.VBComponents(mdl.name).CodeModule.Lines(1, i)
            End If
'and then write it to a file, first marking the start with
'some equal signs and the component name.
            f.WriteLine String(15, "=") & vbCrLf & mdl.name _
                & vbCrLf & String(15, "=") & vbCrLf & strMod
        Next
        
   f.Close
   Set fs = Nothing
End Sub

