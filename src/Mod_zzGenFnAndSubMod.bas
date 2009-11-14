Option Compare Database
Option Explicit
Dim SQLStr As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const NoCurrCardIDGbl = -11111

Public Function MoveFileGSb(PathAndFileNameFm As String, PathAndFileNameTo As String)
Dim RetVal As Variant
Dim DOS_Cmd As String
Dim FNum As Variant
Dim Posn As Long
Dim AppPathAndFileName As String
Dim AppPath As String
Dim BatchFilePathAndFileName As String
AppPathAndFileName = Application.CurrentDb.Name
Posn = InStrRev(Application.CurrentDb.Name, "\")
AppPath = Left(AppPathAndFileName, Posn) 'Includes the "\"
BatchFilePathAndFileName = AppPath & "DOS.BAT"

'Now crate the DOS MOVE Command.
DOS_Cmd = "MOVE /Y """ & PathAndFileNameFm & """ """ & PathAndFileNameTo & """"

'Write the DOS Command to a file using standard old-fashioned File I/O
FNum = FreeFile() 'Assign a free file no. to FNum
Open BatchFilePathAndFileName For Output As FNum 'Let the .BAT file be in the same location as the Application
    Print #FNum, DOS_Cmd 'Write the contents of DOS_Cmd to the batch file.
Close #FNum

'Once written, call the batch file.  Next time it will be overwritten.
RetVal = Shell(BatchFilePathAndFileName, vbHide)

End Function

Public Function ExtractStrGFn(ByVal Str, ByRef StPosn As Long, DelimChar As String) As String
Dim EndPosn As Long
EndPosn = InStr(StPosn, Str, DelimChar)
ExtractStrGFn = Mid(Str, StPosn, EndPosn - StPosn)
StPosn = EndPosn + 1
End Function

Public Function FormatSecs2TmeStrGFn(ByVal TmeInSecs As Variant) As Variant
Dim N As Long
Dim Secs As Long
Dim SecsStr As String
Dim Mins As Long
Dim MinsStr As String
Dim Hrs As Long
Dim HrsStr As String
Dim Remainder As Long
If Not IsNumeric(TmeInSecs) Then
    FormatSecs2TmeStrGFn = Null
Else
    Remainder = TmeInSecs
    Hrs = Remainder \ 3600
    HrsStr = Format(Hrs, "  #,##0Hrs")
    Remainder = Remainder - (3600 * Hrs)
    Mins = (Remainder) \ 60
    Remainder = Remainder - (60 * Mins)
    Secs = Remainder
    
    MinsStr = Format(Mins, " 00Mins")
    SecsStr = Format(Secs, " 00Secs")
    FormatSecs2TmeStrGFn = HrsStr & MinsStr & SecsStr
End If
N = Len(AssignVar2StrGFn(FormatSecs2TmeStrGFn, "", True))
End Function

Private Sub TestConvertTme()
Const TimeLng = 31121
Dim Result As String
Result = FormatSecs2TmeStrGFn(TimeLng)
End Sub

Public Function DirPathExistsGFn(DirPathVar As Variant) As Boolean
On Error GoTo ErrProc
Dim DirPath As String
DirPath = AssignVar2StrGFn(DirPathVar, "No Director Found", True)
    If Dir(DirPath, vbDirectory) <> "" Then
        DirPathExistsGFn = True
    Else
        DirPathExistsGFn = False
    End If
ExitProc:
    Exit Function
ErrProc:
    DirPathExistsGFn = False
End Function

Public Function FileExistsGFn(FileAndPathNameVar As Variant) As Boolean
On Error GoTo ErrProc
Dim PathAndFileName As String
Dim FileName As String
PathAndFileName = AssignVar2StrGFn(FileAndPathNameVar, "No Director Found", True)
If Right(Trim(PathAndFileName), 1) = "\" Then GoTo ErrProc 'Have received a directory
    If Dir(PathAndFileName) <> "" Then
        FileExistsGFn = True
    Else
        FileExistsGFn = False
    End If
ExitProc:
    Exit Function
ErrProc:
    FileExistsGFn = False
End Function

Public Function DSN4TblNameGFn(TblName As String) As String
On Error GoTo ErrProc
DSN4TblNameGFn = CurrentDb.TableDefs(TblName).Connect 'This obtains the DSN attached to the table name passed to us.
ExitProc:
    Exit Function
ErrProc:
    DispErrMsgGSb Error$, "retrieve the DSN for " & TblName, True, "table misspelled, missing, or not an ODBC table"
    Resume ExitProc
End Function

'Public Sub LoadFormGSb(FormName As String, Optional FormView As AcFormView = acNormal, Optional Filter, _
'     Optional Where, Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
'     Optional WindowMode As AcWindowMode = acWindowNormal, Optional ByVal Args As String = "")
'On Error GoTo ErrProc
'DoCmd.OpenForm FormName, FormView, Filter, Where, DataMode, WindowMode, Args
'If WindowMode = acDialog Then
'     Do While (FormIsLoadedGFn(FormName))
'          If Forms(FormName).Visible = False Then Exit Do
'          DoCmd.OpenForm FormName, FormView, Filter, Where, DataMode, WindowMode, Args
'     Loop
'End If
'ExitProc:
'     Exit Sub
'ErrProc:
'     DispErrMsgGSb Error$, "load a form in Public Sub LoadFormGSb", True, _
'     "the form's name being inadvertently changed during source code maintenance"
'End Sub

Public Sub SetDBGeneratorGSb(TblName As String, Optional GenValLng As Long = 0)
Dim QryDef As DAO.QueryDef
SQLStr = "SET GENERATOR GEN_" & TblName & "_ID TO " & GenValLng & ";"
Set QryDef = CurrentDb.CreateQueryDef("")
    QryDef.Connect = DSN4TblNameGFn(TblName)
    QryDef.SQL = SQLStr
    QryDef.ReturnsRecords = False
    QryDef.Execute
QryDef.Close
End Sub

Public Sub ClearTheTableNamedGSb(tableName As String, Optional SetGenBln As Boolean = False)
On Error GoTo ErrProc
'Clean data out a named table
CurrentDb.Execute "DELETE FROM " & tableName
If SetGenBln Then SetDBGeneratorGSb tableName
ExitProc:
    Exit Sub
ErrProc:
    MsgBox "Sorry there has been an error deleting the table named " & tableName & "." & vbCrLf & _
    "Access reports the error as being " & Error$ & ".", vbExclamation, "Problem Deleting Table"
GoTo ExitProc
End Sub

Public Function GetFormIDGFn(formName As String) As Long
Dim RecSet As DAO.Recordset
SQLStr = "SELECT FL_TblID " & _
"FROM zzHlpFormsLstTbl " & _
"WHERE FL_Name ='" & formName & "'"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenDynaset)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        GetFormIDGFn = RecSet("FL_TblID")
    End If
RecSet.Close
End Function

Public Function DefaultValGFn(TblNameStr As String, IDFdNameStr As String, IDFdVal As Long, DefFdNameStr As String) As Boolean
Dim RecSet As DAO.Recordset
SQLStr = "SELECT " & DefFdNameStr & " FROM " & TblNameStr & " " & _
    "WHERE (" & IDFdNameStr & " = " & IDFdVal & ") " & _
    "AND (" & DefFdNameStr & " = -1)"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If (Not RecSet.EOF) And (Not RecSet.BOF) Then
        DefaultValGFn = True
    Else
        DefaultValGFn = False
    End If
RecSet.Close
End Function

'Private Function ConstructWhereStr(GpIDFdName As String, GpIDFdVal As Long, _
'ActiveFdName As String, SortOrderFdName As String) As String
'Dim WhereStr As String
''Creates an extra condition in the Where string of the SQL construction to include Families.
''In this way, GPID 1 can have sort order numbers running from 1 to infinity and GPID 2 can have sortorder numbers
''running from 1 to infinity etc.
'If (GpIDFdName <> "") And (ActiveFdName <> "") Then
'     WhereStr = " WHERE (" & ActiveFdName & "=-1) AND (" & GpIDFdName & "=" & GpIDFdVal & ") "
'ElseIf ActiveFdName <> "" Then
'     WhereStr = " WHERE (" & ActiveFdName & "=-1)"
'ElseIf GpIDFdName <> "" Then
'     WhereStr = " WHERE (" & GpIDFdName & "=" & GpIDFdVal & ") "
'End If
'If SortOrderFdName <> "" Then
'     If WhereStr <> "" Then
'          WhereStr = WhereStr & " AND (" & SortOrderFdName & "<>" & DelCardGbl & ")"
'     Else
'          WhereStr = " WHERE (" & SortOrderFdName & "<>" & DelCardGbl & ")"
'     End If
'End If
'ConstructWhereStr = WhereStr
'End Function

'Public Function GetNextSOGFn(TblName As String, SortOrderFdName As String, Optional ActiveFdName As String = "", _
'Optional GpIDFdName As String = "", Optional GpID As Long = NoCurrCardIDGbl, Optional FailInt As Long = 1) As Long
''This routine handles situations where there is possibly an "Active" Fd and a GpID Fd in the subject tbl.
''It finds the largest SortOrder Number and adds one to it within these of Active and GpID if they apply.
''It is expected that there will be another routine called RatSOGFn which will ensure that all SortOrderNumbers are
''sequential.  This latter function can be applied after Getting and Saving the next SO because then it will ensure that
''this next item will be the last in the list.
''WHERE THIS ROUTINE FAILS TO FIND A SORT ORDER, IT IMPLIES THAT THERE ARE NO ITEMS IN THE LIST AT ALL.
''THEREFORE THE FailInt IS SET TO 1; BEING THE FIRST ITEM IN THE LIST
'On Error GoTo ErrProc
'Dim NewSortOrder As Long
'Dim RecSet As DAO.Recordset
'SQLStr = "SELECT MAX(" & SortOrderFdName & ") AS MaxSO FROM " & TblName & _
'     ConstructWhereStr(GpIDFdName, GpID, ActiveFdName, SortOrderFdName)
''SaveDebugDataGSb SQLStr
'Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
'    If Not IsNull(RecSet("MaxSO")) Then
'        GetNextSOGFn = RecSet("MaxSO") + 1
'    Else
'        GetNextSOGFn = FailInt
'    End If
'RecSet.Close
'ExitProc:
'     Exit Function
'ErrProc:
'     DispErrMsgGSb Error$, "get the next sort order for the list", True, "passing incorrect arguments to this routine. SQLStr=" & SQLStr, _
'     "Cannot Get Next Sort Order Value"
'     GoTo ExitProc
'End Function

Public Function DetNextItemIDGFn(TblName As String, IDFdName As String, IDFdValLng As Long, _
Optional SortOrderFdName As String = "", _
Optional GpIDFdName As String = "", Optional GpIDValLng As Long = NoCurrCardIDGbl) As Long
'
Dim WhereClause As String
Dim OrderByClause As String
Dim RecSet As DAO.Recordset
If GpIDFdName <> "" Then
    WhereClause = "WHERE (" & GpIDFdName & "=" & GpIDValLng & ") "
End If
If SortOrderFdName <> "" Then
    OrderByClause = "ORDER BY " & SortOrderFdName
End If

SQLStr = "SELECT " & IDFdName & " " & _
"FROM " & TblName & " " & _
WhereClause & _
OrderByClause

Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.FindFirst (IDFdName & "=" & IDFdValLng)
        If Not RecSet.NoMatch Then
            If Not RecSet.EOF Then 'One ahead to point at
                RecSet.MoveNext
                DetNextItemIDGFn = RecSet(IDFdName)
            Else
                If Not RecSet.BOF Then
                    RecSet.MovePrevious
                    DetNextItemIDGFn = RecSet(IDFdName)
                Else 'BOF and EOF - signal no more records in list
                    DetNextItemIDGFn = 0
                End If
            End If
        Else 'Signal no records matching
            DetNextItemIDGFn = 0
        End If
    Else 'Signal no records matching
        DetNextItemIDGFn = 0
    End If
RecSet.Close
End Function

'Public Function GetIDGFn(TblName As String, RecSet2Add2 As DAO.Recordset) As Long
'On Error GoTo ErrProc
'Dim DSN4TblStr As String
'Dim RecSet As DAO.Recordset
'DSN4TblStr = CurrentDb.TableDefs(TblName).Connect 'This obtains the DSN attached to the table name passed to us.
'If UCase(Left(DSN4TblStr, 5)) = "ODBC;" Then
'    Set RecSet = ExecProcGFn("get an ID from " & TblName & ".", DSN4TblStr, "SP_GEN_" & TblName & "_ID", True)
'         If Not RecSet.EOF Then
'              RecSet.MoveFirst
'              GetIDGFn = RecSet("ID")
'         Else
'              MsgBox "Sorry. There has been a Database Failure. " & vbCrLf & _
'              "The function which is supposed to get a unique ID " & vbCrLf & _
'              "has been unable to retreive one from the " & TblName & " table.", vbExclamation, _
'              "Unable to Retrieve a Unique ID from " & TblName & " Table"
'         End If
'    RecSet.Close
'    RecSet2Add2(0) = GetIDGFn
'Else
'    GetIDGFn = RecSet2Add2(0)
'End If
'ExitProc:
'     Exit Function
'ErrProc:
'     DispErrMsgGSb Error$, "obtain a Table ID from " & TblName, True, "the table not being present or not being addressed by the correct name"
'End Function

Public Function NoDefaultFoundGFn(TblName As String, DefFdName As String, _
Optional ActiveFdName As String = "", _
Optional WhereStr4Gp As String = "") As Boolean
'Determines if there is a record with Default set within the Gp described by the WhereStr if Gps exist.
'If there is no GpFdName then the WhereGpStr is set to "".
'Expects the format of the WhereGpStr to be for example: "MatGp = 'Wood'" or "SerialGp = 1"
'In this manner we can handle strings or numbers.  We add to this " AND (" & WhereGpStr & ") "
Dim RecSet As DAO.Recordset
Dim WhereStr As String
Dim N
'Set up WhereGpStr if it is not set to ""
If WhereStr4Gp <> "" Then
    WhereStr = " AND (" & WhereStr4Gp & ") "
End If
If ActiveFdName <> "" Then
    WhereStr = WhereStr & " AND (" & ActiveFdName & "= -1) "
End If
'Construct Query String
SQLStr = "SELECT " & DefFdName & " " & _
     "FROM " & TblName & " " & _
     "WHERE (" & DefFdName & "=-1) " & WhereStr
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If (RecSet.BOF) And (RecSet.EOF) Then 'No active default record found
        NoDefaultFoundGFn = True
    Else
        NoDefaultFoundGFn = False
    End If
RecSet.Close

End Function

Private Sub SetNo1Rec2DefGSb(TblName As String, DefFdName As String, SortOrderFdName As String, _
                                Optional ActiveFdName = "", _
                                Optional WhereStr4Gp = "")
'This routine sets the active record with sort order of 1 to being the default.
Dim WhereStr As String
'Set up WhereGpStr if it is not set to ""
If WhereStr4Gp <> "" Then
    WhereStr = " AND (" & WhereStr4Gp & ") "
End If
If ActiveFdName <> "" Then
    WhereStr = WhereStr & " AND (" & ActiveFdName & "= -1) "
End If
SQLStr = "Update " & TblName & " Set " & DefFdName & "= -1 " & _
        "WHERE (" & SortOrderFdName & "=1) " & WhereStr
CurrentDb.Execute SQLStr
End Sub

Public Sub CheckDefaultOKGSb( _
    TblName As String, _
    DefFdName As String, _
    SortOrderFdName As String, _
    Optional ActiveFdName As String = "", _
    Optional WhereStr4Gp As String = "", _
    Optional CreateDefaultIfMissingBln As Boolean = False)
'Check if an active record has the default set.
'If not then set the default to be the record with sortorder of 1
If NoDefaultFoundGFn(TblName, DefFdName, ActiveFdName, WhereStr4Gp) And _
    CreateDefaultIfMissingBln Then
    'Updates table by setting the record with sortorder = 1 to be the default
    SetNo1Rec2DefGSb TblName, DefFdName, SortOrderFdName, ActiveFdName, WhereStr4Gp
End If
End Sub

Private Function GpConvFn(GpIDFdName As String, GpIDFdVal As Long) As String
'Creates an extra condition in the Where string of the SQL construction to include Types.
'In this way, GPID 1 can have sort order numbers running from 1 to infinity and GPID 2 can have sortorder numbers
'running from 1 to infinity etc.
If GpIDFdName <> "" Then
     GpConvFn = "AND (" & GpIDFdName & "=" & GpIDFdVal & ") "
Else
     GpConvFn = ""
End If
End Function

Private Function MakeBooleanSQLClause(BooleanFdName As String, BlnVal As Boolean, _
        Incl_OR As Boolean, GpClause As String) As String
'We encapsulate BooleanFdName in "[" and "]" in case it is some sort of reserved word
If BooleanFdName <> "" Then
    If Incl_OR Then 'Already has a WHERE in the front part of the WHERE Clause.
        MakeBooleanSQLClause = " OR ([" & BooleanFdName & "] = " & BlnVal & ") "
    Else 'Put a Where in front of it.
        MakeBooleanSQLClause = "WHERE ([" & BooleanFdName & "] = " & BlnVal & ") "
    End If
Else
    MakeBooleanSQLClause = "WHERE (1=1) "
End If

If GpClause <> "" Then
    MakeBooleanSQLClause = MakeBooleanSQLClause & GpClause
End If
End Function

Public Sub RatSortOrderSimpleGSb(ByVal TblName As String, ByVal IDFdName As String, ByVal IDFdVal As Long, _
    ByVal SortOrderFdName As String, ByVal SortOrderFdVal As Long, _
    Optional GpIDFdName As String = "", Optional GpIDFdVal As Long = NoCurrCardIDGbl)
'This sub is used when there is no Active Flag on a record and no default.


End Sub

'Public Sub RatSortOrderGSb(ByVal TblName As String, ByVal IDFdName As String, ByVal IDFdVal As Long, _
'    ByVal SortOrderFdName As String, ByVal SortOrderFdVal As Long, Optional ActiveFdName As String = "", _
'    Optional ByVal GpIDFdName As String = "", Optional GpIDFdVal As Long = NoCurrCardIDGbl, Optional NameFdName As String = "")
''This routine rationalises the Sort Order, ie, if there are anomalies in the sort order due to deletions, or additions
''or edits, it makes the sort order continuous starting from 1 through to the number of items in the list.
''In doing this, it both sorts numbers in the order mandated by the user and also "purifies" a corrupted table
''where sort orders are not continuous, ie, one number serially after the other.
'Dim RecSet As DAO.Recordset
'Dim SOCounter As Long
'Dim SeenIDFdValButNotInserted As Boolean
'Dim ActiveClause As String
'On Error GoTo ErrProc
''Make sure there are no null Sort Orders in the table (this should never happen)
''SQLStr = "UPDATE " & TblName & " " & _
''    "SET [" & SortOrderFdName & "] = " & DelCardGbl & " " & _
''    MakeBooleanSQLClause(ActiveFdName, False, True, (" AND [" & SortOrderFdName & "] Is Null"))
'If ActiveFdName <> "" Then
'    ActiveClause = " OR ([" & ActiveFdName & "] = 0)"
'Else
'    ActiveClause = ""
'End If
'SQLStr = "UPDATE " & TblName & " " & _
'    "SET [" & SortOrderFdName & "] = " & DelCardGbl & " " & _
'    "WHERE ([" & SortOrderFdName & "] Is Null)" & ActiveClause
'
'CurrentDb.Execute SQLStr
'
''Note: CLng([" & IDFdName & "] = " & IDFdVal & " in the SQL construction below produces a -1 when true
''and a 0 when false.  In this case if two records share the same sort order, the one that is the
''mandated CurrID will "win" in the sorting.
''NameFdName is used occasionally for debugging.
'If NameFdName <> "" Then NameFdName = ", " & NameFdName & " "
'SQLStr = "SELECT [" & SortOrderFdName & "], [" & IDFdName & "]  " & NameFdName & _
'    "FROM " & TblName & " " & _
'    MakeBooleanSQLClause(ActiveFdName, True, False, GpConvFn(GpIDFdName, GpIDFdVal)) & _
'    "ORDER BY " & SortOrderFdName & ", CLng([" & IDFdName & "] = " & IDFdVal & ")"
'
'SOCounter = 0
'SeenIDFdValButNotInserted = False
'Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenDynaset)
'    If Not RecSet.EOF Then
'        'Renumber all records, except for the record with the ID = to IDFdVal, consecutively.
'        'When we get to the record with the ID of interest we add 1 to the counter, thereby "skipping" this record.
'        RecSet.MoveFirst
'        While Not RecSet.EOF
'            If (RecSet(IDFdName) <> IDFdVal) Or (SOCounter >= SortOrderFdVal) Then
'                SOCounter = SOCounter + 1
'                If SeenIDFdValButNotInserted And (SOCounter = SortOrderFdVal) Then SOCounter = SOCounter + 1
'                RecSet.Edit
'                    RecSet(SortOrderFdName) = SOCounter
'                RecSet.Update
'            Else
'                SeenIDFdValButNotInserted = True
'            End If
'        RecSet.MoveNext
'        Wend
'    End If
'RecSet.Close
'If SeenIDFdValButNotInserted Then
'    If SortOrderFdVal > SOCounter Then SortOrderFdVal = SOCounter + 1
'    CurrentDb.Execute "UPDATE " & TblName & " " & _
'        "SET [" & SortOrderFdName & "] = " & SortOrderFdVal & " " & _
'        "WHERE [" & IDFdName & "] = " & IDFdVal
'End If
'ExitProc:
'     Exit Sub
'ErrProc:
'     MsgBox "Sorry. There has been a error rationalising " & SortOrderFdName & " in Table " & TblName & "." & vbCrLf & "Access reports the error as being:" & vbCrLf & vbCrLf & Error$
'End Sub

Public Function ValUniqueGFn(SearchSqlStr As String) As Boolean
'This function returns False if it finds that the value being input is the same as the value in another record.
'If there is a value provided for GpFdName and GpName it will allow identical names in various groups but not in the same group.
'Execute the query to see if the value exists
Dim RecSet As DAO.Recordset
Set RecSet = CurrentDb.OpenRecordset(SearchSqlStr, dbOpenSnapshot)
    If Not RecSet.EOF Then 'Found a name the same as the one we are trying to add or edit in a different record.
        ValUniqueGFn = False
    Else 'Didn't find value
        ValUniqueGFn = True
    End If
RecSet.Close
End Function

'Public Function ExecProcGFn(ErrMsgElaborationStr As String, DSNStr As String, ProcName As String, ReturnsRecords As Boolean, _
'     ParamArray Args()) As DAO.Recordset
'On Error GoTo ErrProc
'     Dim QryDef As DAO.QueryDef
'     Dim RecSet As DAO.Recordset
'     Dim N As Long
'     If ReturnsRecords Then
'          SQLStr = "SELECT * FROM " & ProcName
'     Else
'          SQLStr = "EXECUTE PROCEDURE " & ProcName
'     End If
'     If (UBound(Args) < LBound(Args)) Then
'          SQLStr = SQLStr & ";"
'     Else
'          SQLStr = SQLStr & "("
'          For N = LBound(Args) To UBound(Args)
'               Select Case VarType(Args(N))
'                    Case vbNull
'                         SQLStr = SQLStr & "NULL"
'                    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
'                         SQLStr = SQLStr & Str(Args(N))
'                    Case vbDate
'                         SQLStr = SQLStr & "'" & CStr(Args(N)) & "'"
'                    Case vbString
'                         SQLStr = SQLStr & "'" & ParseSQLParamGFn(Args(N)) & "'"
'               End Select
'               If N <> UBound(Args) Then
'                    SQLStr = SQLStr & ","
'               Else
'                    SQLStr = SQLStr & ");"
'               End If
'          Next N
'     End If
'     'If Right(SQLStr, 1) = "(" Then SQLStr = Mid(SQLStr, 1, Len(SQLStr) - 2)
'     'SaveDebugDataGSb SQLStr
'     'Debug.Print SQLStr
'     Set QryDef = CurrentDb.CreateQueryDef("")
'          QryDef.Connect = DSNStr
'          'Assign SQL string to QryDef
'          QryDef.SQL = SQLStr
'
'          If ReturnsRecords Then
'               Set ExecProcGFn = QryDef.OpenRecordset()
'          Else
'               QryDef.ReturnsRecords = False
'               QryDef.Execute
'               Set ExecProcGFn = Nothing
'          End If
'     QryDef.Close
'ExitProc:
'     Exit Function
'ErrProc:
'     DispErrMsgGSb Error$, "Execute a procedure using the SQL construction of '" & SQLStr & "'", _
'     True, "the back end procedure being missing or the argument lists no longer being appropriate. " & _
'     ErrMsgElaborationStr
'End Function

Public Function NoOfRecsInTblGFn(TblName As String, Optional WhereStr = "") As Long
On Error GoTo ErrProc
Dim RecSet As DAO.Recordset
SQLStr = "SELECT * FROM " & TblName & " " & WhereStr
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveLast
        NoOfRecsInTblGFn = RecSet.RecordCount
    Else
        NoOfRecsInTblGFn = 0
    End If
RecSet.Close
ExitProc:
    Exit Function
ErrProc:
    DispErrMsgGSb Error$, "determine the number of records in the table named " & TblName, True, _
    vbCrLf & _
    "1. the table not existing," & vbCrLf & _
    "2. the name of the table being spelt incorrectly, or" & vbCrLf & _
    "3. the correct value not being passed to the argument for this function"
End Function

Public Function CBln2YesNoGFn(IntIn As Variant, Optional FailStr As String = "No") As String
If Not IsNull(IntIn) Then
     If IntIn = -1 Then CBln2YesNoGFn = "Yes" Else CBln2YesNoGFn = "No"
Else
     CBln2YesNoGFn = FailStr
End If
End Function

Public Function CBln2TrueFalseGFn(IntIn As Variant, Optional FailVal As Boolean = False) As String
If Not IsNull(IntIn) Then
     If IntIn = -1 Then CBln2TrueFalseGFn = "True" Else CBln2TrueFalseGFn = "False"
Else
     CBln2TrueFalseGFn = "False"
End If
End Function