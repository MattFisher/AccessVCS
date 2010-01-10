Option Compare Database
Option Explicit

Public Function ListTables() As Integer
'Currently (sometimes) takes 54 seconds! WHY?

'Tbl_ID, Tbl_Connect, Tbl_SourceTableName, Tbl_Selected, Tbl_FELinkTblName,
'Tbl_Attributes, Tbl_System, Tbl_Hidden, Tbl_AttachedTable, Tbl_AttachedODBC,
'Tbl_AttachSavePWD, Tbl_AttachExclusive

Dim db As Database
Dim td As TableDef
Dim d As Document
Dim c As Container
Dim TableList As DAO.Recordset
Dim i As Integer
Dim resetTableList As Boolean
resetTableList = True

Set db = Access.CurrentDb

'Use this option if the risk of old versions of TABLE_LIST outweighs
' the benefit of being able to customise export options in it.
'TODO: Change this to depend on a 'Version' Property for the table.
If Not TableExistsInDbGFn(TABLE_LIST_TABLENAME, db) Then
    ' Copy "__TABLE_LIST_TEMPLATE__" from codeDB to currentDB
    DoCmd.TransferDatabase acImport, "Microsoft Access", _
        CodeDb.Name, acTable, "__TABLE_LIST_TEMPLATE__", _
        TABLE_LIST_TABLENAME, True
End If

'If TableExistsInDbGFn(TABLE_LIST_TABLENAME, db) Then
'    'DoCmd.DeleteObject acTable, TABLE_LIST_TABLENAME
'End If
'' Copy "__TABLE_LIST_TEMPLATE__" from codeDB to currentDB
'DoCmd.TransferDatabase acImport, "Microsoft Access", _
'    CodeDb.Name, acTable, "__TABLE_LIST_TEMPLATE__", _
'    TABLE_LIST_TABLENAME, True


If resetTableList Then CurrentDb.Execute "DELETE * FROM " & TABLE_LIST_TABLENAME

Set TableList = CurrentDb.OpenRecordset(TABLE_LIST_TABLENAME, dbOpenTable)
    If Not TableList.EOF Then
        TableList.MoveFirst
    End If

    For Each td In db.TableDefs 'Tables
        Debug.Print "Name:        " & td.Name
        'Debug.Print "Connect:     " & td.Connect
        'Debug.Print "Attributes:  " & td.Attributes
        'Debug.Print "DateCreated: " & td.DateCreated
        'Debug.Print "LastUpdated: " & td.LastUpdated
        'Debug.Print "Source Tbl:  " & td.SourceTableName
        'dbSystemObject, dbHiddenObject, dbAttachedTable, dbAttachedODBC
        If (td.Attributes And dbSystemObject) Then Debug.Print "dbSystemObject"
        If (td.Attributes And dbHiddenObject) Then Debug.Print "dbHiddenObject"
        If (td.Attributes And dbAttachedTable) Then Debug.Print "dbAttachedTable"
        If (td.Attributes And dbAttachedODBC) Then Debug.Print "dbAttachedODBC"
    
        If (td.Name <> TABLE_LIST_TABLENAME) And _
           ((td.Attributes And dbSystemObject) = 0) And _
           ((td.Attributes And dbHiddenObject) = 0) And _
           (Left(td.Name, 4) <> "MSys") Then
            TableList.AddNew
                TableList("Tbl_Name") = td.Name
                TableList("Tbl_SourceTableName") = td.SourceTableName
                TableList("Tbl_Connect") = td.Connect
                TableList("Tbl_Attributes") = td.Attributes
                TableList("Tbl_System") = td.Attributes And dbSystemObject
                TableList("Tbl_Hidden") = td.Attributes And dbHiddenObject
                TableList("Tbl_AttachedTable") = td.Attributes And dbAttachedTable
                TableList("Tbl_AttachedODBC") = td.Attributes And dbAttachedODBC
                TableList("Tbl_AttachSavePWD") = td.Attributes And dbAttachSavePWD
                TableList("Tbl_AttachExclusive") = td.Attributes And dbAttachExclusive
                TableList("Tbl_ContainsBinary") = TableContainsOleFields(td.Name)
                
                If TableList("Tbl_System") Then
                    TableList("Tbl_DispType") = "System"
                ElseIf TableList("Tbl_AttachedTable") Then
                    TableList("Tbl_DispType") = "Linked"
                ElseIf TableList("Tbl_AttachedODBC") Then
                    TableList("Tbl_DispType") = "ODBC"
                Else
                    TableList("Tbl_DispType") = "Local"
                End If
                
                If TableList("Tbl_System") Or _
                   TableList("Tbl_AttachedTable") Or _
                   TableList("Tbl_AttachedODBC") Then
                    TableList("Tbl_ExportSchema") = False
                Else
                    TableList("Tbl_ExportSchema") = True
                End If
                TableList("Tbl_ExportData") = TableList("Tbl_ExportSchema")
            TableList.Update
            
            printHexTest (td.Attributes)
            i = i + 1
        End If
    Next td
TableList.Close

ListTables = i
End Function

''NOT CURRENTLY USED - Done in VBScript
'Private Sub RestoreTableList()
'Dim tableListFilename As String
'Dim db As DAO.Database
'importFolder = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\test\src\"
'tableListFilename = Dir(importFolder & TABLE_LIST_FILENAME)
'If tableListFilename <> "" Then
'    Set db = CodeDb
'    'db.Execute ("DELETE * FROM " & TABLE_LIST_TABLENAME)
'    On Error Resume Next
'    db.TableDefs.Delete TABLE_LIST_TABLENAME
'    Application.ImportXML importFolder & tableListFilename, acStructureAndData
'End If
'End Sub


''NOT CURRENTLY USED - Done in VBScript
'Public Sub RestoreTables()
'
''Tbl_ID, Tbl_Name, Tbl_Connect, Tbl_SourceTableName, Tbl_ExportSchema, Tbl_ExportData,
''Tbl_Attributes, Tbl_System, Tbl_Hidden, Tbl_AttachedTable, Tbl_AttachedODBC,
''Tbl_AttachSavePWD, Tbl_AttachExclusive, Tbl_ContainsBinary
'
'importFolder = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\test\src\"
'
'Dim db As Database
'Dim td As TableDef
'Dim TableList As DAO.Recordset
'Dim i As Integer
'Dim app As Access.Application
'
'Set db = Access.CurrentDb
'Set app = Access.Application
'
'Set TableList = db.OpenRecordset("SELECT * FROM " & TABLE_LIST_TABLENAME, dbOpenSnapshot)
'    If Not TableList.EOF Then
'        TableList.MoveFirst
'        While Not TableList.EOF
'            If TableList("Tbl_AttachedTable") Or _
'               TableList("Tbl_AttachedODBC") Then
'                'File(s) won't be there - create new TableDef
'                db.TableDefs.Delete (TableList("Tbl_Name"))
'                Set td = db.CreateTableDef(TableList("Tbl_Name"), _
'                                           0, _
'                                           TableList("Tbl_SourceTableName"), _
'                                           TableList("Tbl_Connect"))
'                db.TableDefs.Append td
'            ElseIf ((Not TableList("Tbl_System")) And _
'                    (Left(TableList("Tbl_Name"), 4) <> "MSys")) Then
'                'MSysAccessObjects 'Doesn't allow deletion
'                'MSysACEs 'System - doesn't allow deletion
'                'MSysObjects 'System - doesn't allow deletion
'                'MSysQueries 'System - doesn't allow deletion
'                'MSysRelationships 'System - doesn't allow deletion of table or insertion
'
'                'Need to check if the table exists before trying to delete it
'
'                db.TableDefs.Delete (TableList("Tbl_Name"))
'                If TableList("Tbl_ContainsBinary") Then
'                    'Import .xml file (combined schema and data)
'                    app.ImportXML importFolder & "Tbl_" & TableList("Tbl_Name") & ".xml", _
'                                  acStructureAndData
'                Else
'                    'Import .xsd file (schema) then .txt file (data)
'                    app.ImportXML importFolder & "Tbl_" & TableList("Tbl_Name") & ".xsd", _
'                                  acStructureOnly
'                    app.DoCmd.TransferText acImportDelim, , TableList("Tbl_Name"), _
'                                           importFolder & "Tbl_" & TableList("Tbl_Name") & ".txt", _
'                                           True
'                End If
'            Else
'                'System Table
'            End If
'            TableList.MoveNext
'        Wend
'    End If
'TableList.Close
'
'End Sub