Option Compare Database
Option Explicit

Dim C_TotNoOfTblsTxt As Variant
Dim C_TotThisDBProcTxt As Variant
Dim C_DBTblTxt As Variant
Dim SQLStr As String
Dim C_FBDBSelectionFra As Variant
Dim FBDSN As String
Dim RelinkingIsNecessary As Boolean
Dim C_OpDoingTxt As Variant
Dim C_DBNameTxt As Variant
Dim C_MDBorFB_Fra As Variant
Public MDB_BE_PathAndFilenameGbl As String
Public AppNameGbl As String
Public AppDirPathGbl As String
Public MDB_NWP_PathAndFilenameGbl As String
Public MDB_WP_PathAndFilenameGbl As String
Public AbortGbl As Boolean

Dim importFolder As String

Private Sub DelinkTbls(LoTbls As String)
On Error GoTo ErrProc
Dim RecSet As dao.Recordset
Dim FELinkTblName As String
Dim TblCnt As Long
TblCnt = 0
Set RecSet = CurrentDb.OpenRecordset(LoTbls, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveLast C_TotNoOfTblsTxt = RecSet.RecordCount
        RecSet.MoveFirst
        While Not RecSet.EOF
            TblCnt = TblCnt + 1
             C_TotThisDBProcTxt = TblCnt
             C_DBTblTxt = RecSet("BETblName")
            FELinkTblName = UCase(AssignVar2StrGFn(RecSet("FELinkTblName")))
            'On Error GoTo NoSuchTable
            CurrentDb.TableDefs.Delete FELinkTblName
        RecSet.MoveNext
        Wend
    End If
RecSet.Close
ExitProc:
    Exit Sub
ErrProc:
    Resume Next
End Sub

Private Sub Link2MDBTblsOld(TblLstingTblStr As String, DataBaseFilePath As String)
'ReLinkMDB LoTbls, ";DATABASE=" & DBFilePath, acTable, LoT, MT

'TblLstingTblStr is only given for debugging.
'This routine link a table from the Front End to the BackEnd.
'The table in the BackEnd is the usually given the same name in the FrontEnd but can be different if desired.
On Error GoTo ErrorTrap
Dim NameOfTblInBE  As String
Dim NameGiven2TblInFE As String
Dim tdf As dao.TableDef
'Display which Tbl is being delinked.
'NameOfTblInBE = RST("TblName")
'NameGiven2TblInFE = RST("TblName")
Set tdf = CurrentDb.CreateTableDef(NameGiven2TblInFE)
With tdf
    .Connect = DataBaseFilePath
    .SourceTableName = UCase(NameOfTblInBE)
    .Name = NameGiven2TblInFE
End With
CurrentDb.TableDefs.Append tdf
'DoCmd.TransferDatabase acLink, DataBaseType, DataBaseFilePath, acTable, NameOfTblInBE, NameGiven2TblInFE
ExitProc:
Exit Sub
ErrorTrap:

GoTo ExitProc
End Sub

Private Sub RecordMissingTable(TblName As String, CommentStr As String)
Dim MT As dao.Recordset
SQLStr = "SELECT * " & _
            "FROM zzMissingTbl " & _
            "WHERE MT_TblName='" & TblName & "'"
Set MT = CurrentDb.OpenRecordset(SQLStr, dbOpenDynaset)
    If MT.EOF Then
        MT.AddNew
            MT("MT_TblName") = TblName
            MT("MT_Op") = CommentStr
        MT.Update
    End If
MT.Close
End Sub

Private Sub Link2FBTbls(FBDSN As String, TblName As String, DatabaseType As String)
On Error GoTo ErrProc
'LstTblName is provided for use in Missing Tables Error Trap report.
'This routine link a table from the Front End to the BackEnd.
'The table in the BackEnd is the usually given the same name in the FrontEnd but can be different if desired.
'DoCmd.TransferDatabase acLink, "ODBC Database", _
'DataBaseFilePath, _
'acTable, BETblName, FETblName
Dim FBLot As dao.Recordset
SQLStr = "SELECT BETblName, FELinkTblName " & _
            "FROM " & TblName & " " & _
            "ORDER BY BETblName"
Set FBLot = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not FBLot.EOF Then
        FBLot.MoveFirst
        While Not FBLot.EOF
            DoCmd.TransferDatabase acLink, "ODBC Database", _
            FBDSN, acTable, UCase(FBLot("BETblName")), FBLot("FELinkTblName"), , True
        FBLot.MoveNext
        Wend
    End If
FBLot.Close
ExitProc:
    Exit Sub
ErrProc:
    Resume Next
    RecordMissingTable TblName, DatabaseType & " Missing Table - " & Error$
GoTo ExitProc
End Sub

Private Sub Link2MDBTbls(LoTbls As String, DBFilePath As String, DatabaseType As String)
'ListTblNameStr is the Table that has a list of all the tables which are to be delinked and then relinked.
On Error GoTo ErrProc
Dim LoT As dao.Recordset 'zzAppLoTMDBBETbl
'Fields in zzAppLoTMDBBETbl
'TblID, BETblName, Selected, FELinkTblName
Dim NoOfTblProc As Long
Dim NameOfTblInBE As String
Dim NameGiven2TblInFE As String
Dim TblDef As dao.TableDef
NoOfTblProc = 0
Set LoT = CurrentDb.OpenRecordset(LoTbls, dbOpenSnapshot)
     NoOfTblProc = 0 'Reset counter
     'For debugging we can elect, using zzLinkTablesForm!C_RelinkChk, not to reconnect the table
     'once we have completed the delink operation.
     'Doing this exposes all of those linked tables which are not mentioned in the particular table.
     'For example, if you are relinking the WP tables then once you have delinked,
     'there should be no WPTables evident.
     LoT.MoveLast C_TotNoOfTblsTxt = LoT.RecordCount
     LoT.MoveFirst
     While Not LoT.EOF
            'NoOfTblProc displays the number of tables processed within the present database being linked.
            NoOfTblProc = NoOfTblProc + 1:   C_TotThisDBProcTxt = NoOfTblProc
             C_DBTblTxt = LoT("BETblName")
            'Display which Tbl is being delinked.
            NameOfTblInBE = UCase(LoT("BETblName")) 'It is necessary for the table names to be in Uppercase.
            NameGiven2TblInFE = LoT("FELinkTblName")
            Set TblDef = CurrentDb.CreateTableDef(NameGiven2TblInFE)
            With TblDef
                .Connect = ";DATABASE=" & DBFilePath
                .SourceTableName = UCase(NameOfTblInBE)
                .Name = NameGiven2TblInFE
            End With
            CurrentDb.TableDefs.Append TblDef
          
          'ReLinkMDB LoTbls, ";DATABASE=" & DBFilePath, acTable, LoT, MT
     LoT.MoveNext
     Wend
LoT.Close
ExitProc:
    Exit Sub
ErrProc:
    RecordMissingTable AssignVar2StrGFn(LoT("BETblName"), "NoTableNameInRecord", True), DatabaseType & _
    " Missing Table - " & AssignVar2StrGFn(LoT("BETblName"), "NoTableNameInRecord", True)
    Resume Next
End Sub

Private Sub SetFBDSN()
Dim RecSet As dao.Recordset
SQLStr = "SELECT FB_DSN, FB_Location " & _
    "FROM zzFBDescTbl WHERE FB_TblID= " & AssignVar2LngGFn(C_FBDBSelectionFra)
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        FBDSN = AssignVar2StrGFn(RecSet("FB_DSN"), "DataSourceName Not In zzFBDescTbl", True)
        
    End If
RecSet.Close
End Sub

Private Function TablesMissingFn() As Boolean
Dim MT As dao.Recordset
SQLStr = "SELECT COUNT(MT_TblID) AS NoOfTables " & _
"FROM zzMissingTbl"
Set MT = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not MT.EOF Then
        If MT("NoOfTables") > 0 Then
            TablesMissingFn = True
        Else
            TablesMissingFn = False
        End If
    End If
MT.Close
End Function

Private Sub LinkTables()
ClearTheTableNamedGSb "zzMissingTbl"
'Get the path and file names of the two databases for the purposes of relinking.
'Note  we only come here once we have verified these files exist so there should not be an error.
'Determine the program's path and file na
'AppNameGbl
'AppDirPathGbl
'MDB_NWP_PathAndFilenameGbl
'MDB_WP_PathAndFilenameGbl
'MDB_BE_PathAndFilenameGbl
GetProgramFilesDetailsGSb RelinkingIsNecessary
'Delink all existing links to tables regardless of database being used.
 C_OpDoingTxt = "DELINKING"
 C_DBNameTxt = "WORKPAD"
DelinkTbls "zzAppLoTWPTbl"
 C_DBNameTxt = "MDB BACKEND"
DelinkTbls "zzAppLoTMDBBETbl"
 C_DBNameTxt = "FIREBIRD BACKEND"
DelinkTbls "zzAppLoTFBTbl"

'Link to tables in workpad regardless of backend database being used.
 C_OpDoingTxt = "LINKING"
 C_DBNameTxt = "WORKPAD"
Link2MDBTbls "zzAppLoTWPTbl", MDB_WP_PathAndFilenameGbl, "MDB-WorkPad"
'Depending on selection, link to Firebird or Link to MDB
Select Case C_MDBorFB_Fra
    Case 1 'MDB
         C_DBNameTxt = "MDB BACKEND"
        Link2MDBTbls "zzAppLoTMDBBETbl", MDB_BE_PathAndFilenameGbl, "MDB-BackEnd"
        'Temporarily link with FB also here.
         C_DBNameTxt = "FIREBIRD BACKEND"
        Link2FBTbls FBDSN, "zzAppLoTFBTbl", "FirebirdSQL-BackEnd"
    Case 2 'Firebird
         C_DBNameTxt = "FIREBIRD BACKEND"
        Link2FBTbls FBDSN, "zzAppLoTFBTbl", "FirebirdSQL-BackEnd"
End Select
If TablesMissingFn() Then
    'zzBvrObj.OpenForm "zzMissingTblsForm", , , , , acDialog
    MsgBox "Tables missing!"
Else
    MsgBox "All tables were linked successfully.", vbInformation, "Tables Linked Successfully"
End If
AbortGbl = False
DoCmd.Close
End Sub

Public Function ListTables() As Integer
'Currently (sometimes) takes 54 seconds! WHY?

'Tbl_ID, Tbl_Connect, Tbl_SourceTableName, Tbl_Selected, Tbl_FELinkTblName,
'Tbl_Attributes, Tbl_System, Tbl_Hidden, Tbl_AttachedTable, Tbl_AttachedODBC,
'Tbl_AttachSavePWD, Tbl_AttachExclusive

Dim db As Database
Dim td As TableDef
Dim d As Document
Dim c As Container
Dim TableList As dao.Recordset
Dim i As Integer
Dim resetTableList As Boolean
resetTableList = True

'MsgBox "In the ListTables Function, CurrentDB: " & Application.CurrentDb.Name
'If we want a reference to the Code (AddIn) Database at this point, how do we get it?
'Say if we wanted __TABLE_LIST__ to be a temporary table in the Addin.
'Would this lose preferences between runs?
'MsgBox "Access.CodeDb.Name: " & Access.CodeDb.Name

Set db = Access.CurrentDb
'If exported version of the Table List exists, import it.
If Not TableExistsInDbGFn(TABLE_LIST_TABLENAME, db) Then
    ' Copy "__TABLE_LIST_TEMPLATE__" from codeDB to currentDB
    DoCmd.TransferDatabase acImport, "Microsoft Access", _
        CodeDb.Name, acTable, "__TABLE_LIST_TEMPLATE__", _
        TABLE_LIST_TABLENAME, True
End If

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
           ((td.Attributes And dbHiddenObject) = 0) Then
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


Private Sub RestoreTableList()
Dim tableListFilename As String
Dim db As dao.Database
importFolder = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\test\src\"
tableListFilename = Dir(importFolder & TABLE_LIST_FILENAME)
If tableListFilename <> "" Then
    Set db = CodeDb
    'db.Execute ("DELETE * FROM " & TABLE_LIST_TABLENAME)
    On Error Resume Next
    db.TableDefs.Delete TABLE_LIST_TABLENAME
    Application.ImportXML importFolder & tableListFilename, acStructureAndData
End If
End Sub

Public Sub RestoreTables()

'Tbl_ID, Tbl_Name, Tbl_Connect, Tbl_SourceTableName, Tbl_ExportSchema, Tbl_ExportData,
'Tbl_Attributes, Tbl_System, Tbl_Hidden, Tbl_AttachedTable, Tbl_AttachedODBC,
'Tbl_AttachSavePWD, Tbl_AttachExclusive, Tbl_ContainsBinary

importFolder = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\test\src\"

Dim db As Database
Dim td As TableDef
Dim TableList As dao.Recordset
Dim i As Integer
Dim app As Access.Application

Set db = Access.CurrentDb
Set app = Access.Application

Set TableList = db.OpenRecordset("SELECT * FROM " & TABLE_LIST_TABLENAME, dbOpenSnapshot)
    If Not TableList.EOF Then
        TableList.MoveFirst
        While Not TableList.EOF
            If TableList("Tbl_AttachedTable") Or _
               TableList("Tbl_AttachedODBC") Then
                'File(s) won't be there - create new TableDef
                db.TableDefs.Delete (TableList("Tbl_Name"))
                Set td = db.CreateTableDef(TableList("Tbl_Name"), _
                                           0, _
                                           TableList("Tbl_SourceTableName"), _
                                           TableList("Tbl_Connect"))
                db.TableDefs.Append td
            ElseIf Not TableList("Tbl_System") Then
                'MSysAccessObjects 'Doesn't allow deletion
                'MSysACEs 'System - doesn't allow deletion
                'MSysObjects 'System - doesn't allow deletion
                'MSysQueries 'System - doesn't allow deletion
                'MSysRelationships 'System - doesn't allow deletion of table or insertion
                
                'Need to check if the table exists before trying to delete it
                
                db.TableDefs.Delete (TableList("Tbl_Name"))
                If TableList("Tbl_ContainsBinary") Then
                    'Import .xml file (combined schema and data)
                    app.ImportXML importFolder & "Tbl_" & TableList("Tbl_Name") & ".xml", _
                                  acStructureAndData
                Else
                    'Import .xsd file (schema) then .txt file (data)
                    app.ImportXML importFolder & "Tbl_" & TableList("Tbl_Name") & ".xsd", _
                                  acStructureOnly
                    app.DoCmd.TransferText acImportDelim, , TableList("Tbl_Name"), _
                                           importFolder & "Tbl_" & TableList("Tbl_Name") & ".txt", _
                                           True
                End If
            Else
                'System Table
            End If
            TableList.MoveNext
        Wend
    End If
TableList.Close

End Sub