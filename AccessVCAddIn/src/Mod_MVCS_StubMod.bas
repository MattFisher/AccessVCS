Option Compare Database
Option Explicit

Public Sub Test_createStubDatabase()
createStubDatabase CurrentDb, GetFSO.GetParentFolderName(CurrentDb.Name) & "\" & GetFSO.GetBaseName(CurrentDb.Name) & "\"
End Sub

Public Sub createStubDatabase(wholeDB As DAO.Database, stubDBPath As String, Optional overWrite As Boolean = True)
' This will most likely only work with .mdbs anyway
On Error GoTo ErrProc

Dim stubDBFilename As String
Dim stubTempDBFilename As String
Dim stubDBPathAndFilename As String
Dim stubTempDBPathAndFilename As String

Dim acc As Access.Application
Dim i As Integer
Dim strName As String

stubDBFilename = GetFSO.GetBaseName(wholeDB.Name) & "_stub." & GetFSO.GetExtensionName(wholeDB.Name)
stubDBPathAndFilename = stubDBPath & "\" & stubDBFilename
stubTempDBFilename = GetFSO.GetBaseName(wholeDB.Name) & "_stubTemp." & GetFSO.GetExtensionName(wholeDB.Name)
stubTempDBPathAndFilename = stubDBPath & "\" & stubTempDBFilename

If FileExistsGFn(stubTempDBPathAndFilename) Then
    If overWrite Then
        Kill (stubTempDBPathAndFilename)
    Else
    'Do nothing - messagebox maybe
        Exit Sub
    End If
End If
GetFSO.CopyFile wholeDB.Name, stubTempDBPathAndFilename

'Now remove all database objects from the stub.

Set acc = New Access.Application
acc.OpenCurrentDatabase stubTempDBPathAndFilename
 
Debug.Print "Deleting Forms"
For i = acc.CurrentProject.AllForms.Count - 1 To 0 Step -1
    strName = acc.CurrentProject.AllForms(i).Name
    Debug.Print strName
    acc.DoCmd.DeleteObject acForm, strName
Next i
 
Debug.Print "Deleting Reports"
For i = acc.CurrentProject.AllReports.Count - 1 To 0 Step -1
    strName = acc.CurrentProject.AllReports(i).Name
    Debug.Print strName
    acc.DoCmd.DeleteObject acReport, strName
Next i
 
Debug.Print "Deleting Modules"
For i = acc.CurrentProject.AllModules.Count - 1 To 0 Step -1
    strName = acc.CurrentProject.AllModules(i).Name
    Debug.Print strName
    acc.DoCmd.DeleteObject acModule, strName
Next i

Debug.Print "Deleting Macros"
For i = acc.CurrentProject.AllMacros.Count - 1 To 0 Step -1
    strName = acc.CurrentProject.AllMacros(i).Name
    Debug.Print strName
    acc.DoCmd.DeleteObject acMacro, strName
Next i

Debug.Print "Deleting Queries"
For i = acc.CurrentData.AllQueries.Count - 1 To 0 Step -1
    strName = acc.CurrentData.AllQueries(i).Name
    Debug.Print strName
    acc.DoCmd.DeleteObject acQuery, strName
Next i

Debug.Print "Deleting Tables"
For i = acc.CurrentData.AllTables.Count - 1 To 0 Step -1
    strName = acc.CurrentData.AllTables(i).Name
    Debug.Print strName
    If Left(strName, 4) <> "MSys" Then
        acc.DoCmd.DeleteObject acTable, strName
    End If
Next i

acc.CloseCurrentDatabase

If FileExistsGFn(stubDBPathAndFilename) Then
    Kill (stubDBPathAndFilename)
End If
If (acc.CompactRepair(stubTempDBPathAndFilename, stubDBPathAndFilename)) Then
    GetFSO.DeleteFile stubTempDBPathAndFilename
End If
acc.Quit

Exit Sub
ErrProc:
DispErrMsgGSb Error$, "create a stub database"

''for each tabledef
'Dim objectType
'Dim objectTypes As New Collection
'objectTypes.Add "Forms"
''objectTypes.Add "Classes"
'objectTypes.Add "Reports"
'objectTypes.Add "Scripts"
'objectTypes.Add "Modules"
'objectTypes.Add "Tables"
'
'Dim c As Container
'Dim d As Document
'
'For Each objectType In objectTypes
'    Debug.Print objectType
'    Set c = stubDB.Containers(objectType)
'    For Each d In c.Documents
'        If (Left(d.Name, 1) <> "~") And (Left(d.Name, 4) <> "MSys") Then
'            Debug.Print "Deleting: " & d.Name
'            ' , d.Name
'        End If
'    Next d
'Next objectType
'
'Dim i As Integer
'For i = 0 To stubDB.QueryDefs.Count - 1
'    'Skip the embedded queries
'    If Left(stubDB.QueryDefs(i).Name, 1) <> "~" Then
'        Debug.Print "Exporting Query: " & stubDB.QueryDefs(i).Name
'    End If
'Next i
        
End Sub