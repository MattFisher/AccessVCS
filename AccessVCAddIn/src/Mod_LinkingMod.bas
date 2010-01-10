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
Dim RecSet As DAO.Recordset
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
Dim tdf As DAO.TableDef
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
Dim MT As DAO.Recordset
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
Dim FBLot As DAO.Recordset
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
Dim LoT As DAO.Recordset 'zzAppLoTMDBBETbl
'Fields in zzAppLoTMDBBETbl
'TblID, BETblName, Selected, FELinkTblName
Dim NoOfTblProc As Long
Dim NameOfTblInBE As String
Dim NameGiven2TblInFE As String
Dim TblDef As DAO.TableDef
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
Dim RecSet As DAO.Recordset
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
Dim MT As DAO.Recordset
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