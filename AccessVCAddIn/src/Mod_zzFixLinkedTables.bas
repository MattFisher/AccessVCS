Option Compare Database
Option Explicit

Public LoopCounterGbl As Long


Private Sub Delink(LstTblName As String, TblName As Variant, MT As DAO.Recordset)
'The MT recordset is for recording Missing Tables.
'LstTblName is provided for debugging purposes and for recording in the Missing Table Record is necessary.
On Error GoTo RecordErr 'On Error trigger gets reset every time you enter this routine.
'To try to reset it repetitively within a sub-routine produces unreliable results.
'Debug.Print TblName
If (Not IsNull(TblName)) And (Trim(TblName) <> "") Then
     CurrentDb.TableDefs.Delete TblName 'Deletes this listing from the API's TblDef's Tbl
Else
     TblName = "TblName'sBlank"
     GoTo RecordErr
End If
ExitProc:
     Exit Sub
RecordErr:
     MT.AddNew
          MT("MT_TblName") = TblName
          MT("MT_LstTblName") = LstTblName
          MT("MT_Op") = "Delinking - " & Error$
     MT.Update
     GoTo ExitProc
End Sub

Public Sub DelinkTbls(DataBaseFilePath As String, LstTblNameStr As String)
'ListTblNameStr is the Table that has a list of all the tables which are to be delinked and then relinked.
Dim LoT As DAO.Recordset
Dim MT As DAO.Recordset
Dim NoOfTblProc As Long
NoOfTblProc = 0
Set MT = CurrentDb.OpenRecordset("zzMissingTbl", dbOpenDynaset)
     'List of Tables in the FB Database
     Set LoT = CurrentDb.OpenRecordset( _
          "SELECT " & LstTblNameStr & ".* FROM " & LstTblNameStr & " " & _
          "ORDER BY " & LstTblNameStr & ".TblName;", dbOpenSnapshot)
          If Not LoT.EOF Then
               LoT.MoveFirst
               While Not LoT.EOF
                    Forms!zzLinkTablesForm!C_OperationBeingPerformedTxt = "DeLinking " & LstTblNameStr
                    'LoopCounterGbl displays the total number of Tbls processed.
                    LoopCounterGbl = LoopCounterGbl + 1
                    Forms!zzLinkTablesForm!C_TotNoProcTxt = LoopCounterGbl
                    'NoOfTblProc displays the number of tables processed within the present database being linked.
                    NoOfTblProc = NoOfTblProc + 1: Forms!zzLinkTablesForm!C_TotThisDBProcTxt = NoOfTblProc
                    Forms!zzLinkTablesForm!C_TableBeingProcessedTxt = LoT("TblName")
                    Forms!zzLinkTablesForm.Repaint
                    Delink LstTblNameStr, LoT("TblName"), MT 'Actual routine that does the delinking of Firebird
                    'database tables from the Access Front end.
               LoT.MoveNext
               Wend
          End If 'End if Delinking operation
     LoT.Close
MT.Close
End Sub

Private Sub ReLinkFB(FETblName As String, BETblName As String, LstTblName As String, _
DataBaseFilePath As String, acTable, MT As DAO.Recordset)
On Error GoTo ErrTrap
'LstTblName is provided for use in Missing Tables Error Trap report.
'This routine link a table from the Front End to the BackEnd.
'The table in the BackEnd is the usually given the same name in the FrontEnd but can be different if desired.
'DoCmd.TransferDatabase acLink, "ODBC Database", _
'DataBaseFilePath, _
'acTable, BETblName, FETblName
DoCmd.TransferDatabase acLink, "ODBC Database", _
DataBaseFilePath, acTable, UCase(BETblName), FETblName, , True
ExitProc:
Exit Sub
ErrTrap:
'Debug.Print Error$
MT.AddNew
    MT("MT_TblName") = BETblName
    MT("MT_LstTblName") = LstTblName
    MT("MT_Op") = "Relinking FB - " & Error$
MT.Update
GoTo ExitProc
End Sub

Public Sub ReLinkFBTbls(DataBaseFilePath As String, LstTblNameStr As String, Optional isODBC As Boolean = False)
'ListTblNameStr is the Table that has a list of all the tables which are to be delinked and then relinked.
Dim LoT As DAO.Recordset
Dim MT As DAO.Recordset
Dim NoOfTblProc As Long
Dim qdf As DAO.QueryDef
NoOfTblProc = 0
Set MT = CurrentDb.OpenRecordset("zzMissingTbl", dbOpenDynaset)
     Set LoT = CurrentDb.OpenRecordset( _
          "SELECT " & LstTblNameStr & ".* FROM " & LstTblNameStr & " " & _
          "ORDER BY " & LstTblNameStr & ".TblName;", dbOpenSnapshot)
          NoOfTblProc = 0 'Reset counter
          'For debugging we can elect, using zzLinkTablesForm!C_RelinkChk, not to reconnect the table
          'once we have completed the delink operation.
          'Doing this exposes all of those linked tables which are not mentioned in the particular table.
          'For example, if you are relinking the WP tables then once you have delinked,
          'there should be no WPTables evident.
          LoT.MoveFirst
          While Not LoT.EOF
               'LoopCounterGbl displays the total number of Tbls processed.
               LoopCounterGbl = LoopCounterGbl + 1: Forms!zzLinkTablesForm!C_TotNoProcTxt = LoopCounterGbl
               'NoOfTblProc displays the number of tables processed within the present database being linked.
               NoOfTblProc = NoOfTblProc + 1: Forms!zzLinkTablesForm!C_TotThisDBProcTxt = NoOfTblProc
               Forms!zzLinkTablesForm!C_OperationBeingPerformedTxt = "Linking " & LstTblNameStr
               Forms!zzLinkTablesForm!C_TableBeingProcessedTxt = LoT("TblName")
               Forms!zzLinkTablesForm.Repaint
               'RelinkFB FETblName, BETblName, LstTblName, DBFilePath, acTable, MissingTbl RecSet
               ReLinkFB LoT("TblName"), UCase(LoT("TblName")), LstTblNameStr, DataBaseFilePath, acTable, MT
          LoT.MoveNext
          Wend
          For Each qdf In CurrentDb.QueryDefs
               If UCase(Left(qdf.Name, 2)) = "BE" Then
                    qdf.Connect = DataBaseFilePath
               End If
          Next
      LoT.Close
MT.Close
End Sub
    
Private Sub ReLinkMDB(TblLstingTblStr As String, DataBaseFilePath As String, acTable, _
RST As DAO.Recordset, MT As DAO.Recordset)
'TblLstingTblStr is only given for debugging.
'This routine link a table from the Front End to the BackEnd.
'The table in the BackEnd is the usually given the same name in the FrontEnd but can be different if desired.
On Error GoTo ErrorTrap
Dim NameOfTblInBE  As String
Dim NameGiven2TblInFE As String
Dim tdf As DAO.TableDef
'Display which Tbl is being delinked.
NameOfTblInBE = RST("TblName")
NameGiven2TblInFE = RST("TblName")
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
    MT.AddNew
        MT("MT_TblName") = RST!TblName
        MT("MT_LstTblName") = TblLstingTblStr
        MT("MT_Op") = "RelinkingMDB - " & Error$
    MT.Update
GoTo ExitProc
End Sub

Public Sub ReLinkMDBTbls(DataBaseFilePath As String, ListTblNameStr As String, Optional isODBC As Boolean = False)
'ListTblNameStr is the Table that has a list of all the tables which are to be delinked and then relinked.
Dim LoT As DAO.Recordset
Dim MT As DAO.Recordset
Dim NoOfTblProc As Long
NoOfTblProc = 0
Set MT = CurrentDb.OpenRecordset("zzMissingTbl", dbOpenDynaset)
     Set LoT = CurrentDb.OpenRecordset( _
          "SELECT " & ListTblNameStr & ".* FROM " & ListTblNameStr & " " & _
          "ORDER BY " & ListTblNameStr & ".TblName;", dbOpenSnapshot)
          NoOfTblProc = 0 'Reset counter
          'For debugging we can elect, using zzLinkTablesForm!C_RelinkChk, not to reconnect the table
          'once we have completed the delink operation.
          'Doing this exposes all of those linked tables which are not mentioned in the particular table.
          'For example, if you are relinking the WP tables then once you have delinked,
          'there should be no WPTables evident.
          LoT.MoveFirst
          While Not LoT.EOF
               'LoopCounterGbl displays the total number of Tbls processed.
               LoopCounterGbl = LoopCounterGbl + 1: Forms!zzLinkTablesForm!C_TotNoProcTxt = LoopCounterGbl
               'NoOfTblProc displays the number of tables processed within the present database being linked.
               NoOfTblProc = NoOfTblProc + 1: Forms!zzLinkTablesForm!C_TotThisDBProcTxt = NoOfTblProc
               Forms!zzLinkTablesForm!C_OperationBeingPerformedTxt = "Linking " & ListTblNameStr
               Forms!zzLinkTablesForm!C_TableBeingProcessedTxt = LoT("TblName")
               Forms!zzLinkTablesForm.Repaint
               ReLinkMDB ListTblNameStr, ";DATABASE=" & DataBaseFilePath, acTable, LoT, MT
          LoT.MoveNext
          Wend
      LoT.Close
MT.Close
End Sub
'DoCmd.TransferDatabase acLink, "ODBC Database", _
    "ODBC;DSN=DataSource1;UID=User2;PWD=www;LANGUAGE=us_english;" _
    & "DATABASE=pubs", acTable, "Authors", "dboAuthors"