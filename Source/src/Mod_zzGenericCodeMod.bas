Option Compare Database
Option Explicit


Public Function FormatDirPathGFn(DirPath As Variant, Optional AllowZeroLengthDirPathBln As Boolean = False) As String
'Format a Dirpath to have a "\" at the end of it.  If the dirpath is null or zero length and this is not allowed then
'show an error and exit.
On Error GoTo ErrProc
DirPath = AssignVar2StrGFn(DirPath) 'Handles the Null value possibility
If DirPath <> "" Then
    If Right(Trim(DirPath), 1) <> "\" Then DirPath = Trim(DirPath) & "\" 'put a "\" on the end of the string if necessary
    FormatDirPathGFn = DirPath
ElseIf Not AllowZeroLengthDirPathBln Then 'Not allow it to be a Null or zero Length
    DispErrMsgGSb "", "format a directory path", "the directory path being of zero length or a null value"
Else 'Null or zero length
    FormatDirPathGFn = "\"
End If
ExitProc:
    Exit Function
ErrProc:
    DispErrMsgGSb Error$, "format the directory path given as " & DirPath, "the variable DirPath not being a string"
End Function

'
'Public Const NoCurrCardIDGbl = -11111
'
'Private SQLStr As String
'
'Public Function GetFormHlpTxtGFn(FormName As String, LocalProgMode As Long, Optional FormHlpTblID As Long = NoCurrCardIDGbl) As Variant
'On Error GoTo ErrProc
'Dim RecSet As DAO.Recordset
'SQLStr = "SELECT FH_TblID, FH_Help " & _
'        "FROM zzHlpFormsLstTbl " & _
'        "INNER JOIN zzHlp4FormsTxtTbl ON zzHlpFormsLstTbl.FL_TblID = zzHlp4FormsTxtTbl.FH_FormID " & _
'        "WHERE ((FL_Name='" & FormName & "') AND (FH_ProgModeID=" & LocalProgMode & "));"
'    Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
'        If Not RecSet.EOF Then
'            GetFormHlpTxtGFn = RecSet("FH_Help")
'            FormHlpTblID = RecSet("FH_TblID")
'        Else
'            GetFormHlpTxtGFn = ""
'            FormHlpTblID = NoCurrCardIDGbl
'        End If
'    RecSet.Close
'ExitProc:
'    Exit Function
'ErrProc:
'    'Don't stop.
'    GetFormHlpTxtGFn = ""
'    FormHlpTblID = NoCurrCardIDGbl
'End Function
'
'Public Function GetFormFdHlpTxtGFn(FdID As Long) As Variant
'Dim RecSet As DAO.Recordset
'SQLStr = "SELECT zzHlp4FdsTxtTbl.FFH_HELP " & _
'            "FROM zzHlp4FdsTxtTbl " & _
'            "WHERE (zzHlp4FdsTxtTbl.FFH_FDID=" & FdID & ")"
'Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
'    If Not RecSet.EOF Then
'        GetFormFdHlpTxtGFn = RecSet("FFH_Help")
'    Else
'        GetFormFdHlpTxtGFn = Null
'    End If
'RecSet.Close
'End Function
'
'Public Function GetFdHlpGFn(FormName As String, CtrlName As String) As Variant
'Dim RecSet As DAO.Recordset
'SQLStr = "SELECT zzHlpFdLstTbl.FFL_DefaultHlp " & _
'"FROM zzHlpFormsLstTbl " & _
'"INNER JOIN zzHlpFdLstTbl ON zzHlpFormsLstTbl.FL_TblID = zzHlpFdLstTbl.FFL_FormID " & _
'"WHERE ((zzHlpFdLstTbl.FFL_Name ='" & CtrlName & "') " & _
'"AND (zzHlpFormsLstTbl.FL_Name='" & FormName & "'));"
'
'Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
'    If Not RecSet.EOF Then
'        GetFdHlpGFn = AssignVar2StrGFn(RecSet("FFL_DefaultHlp"), "", True)
'    Else
'        GetFdHlpGFn = ""
'    End If
'RecSet.Close
''MsgBox "I am a stub.  Please implement me!"
'End Function
'
'Public Function GetFormParentGFn(Ctrl As Access.Control) As Object
'Dim obj As Object
'Set obj = Ctrl
'While (TypeOf obj Is Access.Control)
'    Set obj = Ctrl.Parent
'Wend
'ExitProc:
'    Set GetFormParentGFn = obj
'    Exit Function
'ErrProc:
'    MsgBox "Something isn't right"
'End Function
'
'Public Function FdOKGFn(obj As Object, _
'FrmFdName As String, FdDispName4Msg As String, _
'Optional TblName As String = "", Optional TblFdName As String = "", Optional TblFdTyp As String = "", _
'Optional TblIDFdName As String = "", Optional TblIDFdTyp As String = "", Optional TblIDFdVar As Variant = Null, _
'Optional TblGpFdName As String = "", Optional TblGpIDFdTyp As String, Optional TblGpIDFdVar As Variant = Null) As Boolean
''This routine checks:
''    1. There is a value in the field of the form.
''    2. The value is not blank spaces.
''If a TblName is specified then it checks:
''    1. The value is unique to that table or,
''    2. If a group is specified, the value is unique to that group within that table.
''The unique routine will check for both strings or for numbers depending on the field type specified.
'FdOKGFn = False
''Check that the value is not Null.
'If IsNull(obj(FrmFdName)) Then 'Hasn't put anything in the field. This is not allowed.
'    MsgBox "You must ensure that there is a value provided for the " & FdDispName4Msg & ".", vbInformation, _
'    "Form Not Yet Completely Filled Out"
'    obj(FrmFdName).SetFocus
'    Exit Function
'End If
''Given the value is not Null, now check that it is not a "" or a 0 depending on whether it is a string or a number.
'Select Case UCase(TblFdTyp)
'     Case "STR"
'          If Trim(obj(FrmFdName)) = "" Then 'Only has spaces.  This is not allowed.
'               MsgBox "At present there are only blank spaces in the " & FdDispName4Msg & "." & vbCrLf & _
'               "You must ensure that there is a proper value provided for this field.", vbInformation, _
'              "Form Not Yet Completely Filled Out"
'              obj(FrmFdName).SetFocus
'              Exit Function
'          End If
'     Case "NUM"
'          If obj(FrmFdName) = 0 Then 'Has a zero.  This is not allowed.
'               MsgBox "You presently have a zero value in the " & FdDispName4Msg & "." & vbCrLf & _
'               "You must ensure that there is a proper value provided for this field.", vbInformation, _
'              "Form Not Yet Completely Filled Out"
'              obj(FrmFdName).SetFocus
'              Exit Function
'          End If
'End Select
''It is not Null and not a zero length string nor a 0.
''If we have provided a table name, we check if the item is unique in that table.
''If we have provided a TblGpFdName then we check if the item is unique amongst that group as opposed to the entire table.
'Dim SearchSqlStr As String
'If TblName <> "" Then 'We want to check if the field is unique
'     SearchSqlStr = CreateSQLStr4CheckingUniqueness(TblName, _
'     TblFdName, TblFdTyp, obj(FrmFdName), _
'     TblIDFdName, TblIDFdTyp, TblIDFdVar, _
'     TblGpFdName, TblGpIDFdTyp, TblGpIDFdVar)
'     'SaveDebugDataGSb SearchSqlStr
'     If Not ValUniqueGFn(SearchSqlStr) Then
'          MsgBox "Sorry." & vbCrLf & _
'          "The " & FdDispName4Msg & " must be unique." & vbCrLf & _
'          "This applies to both active values and to inactive values.", vbExclamation, "Please chose a unique name."
'          obj(FrmFdName).SetFocus
'          Exit Function
'     End If
'End If
'FdOKGFn = True
'End Function
'
'Public Sub SetRowSrc4LstGSb(C_ItemLst As Access.ListBox, RowSrcStr As String, _
'Optional SetFocusBln As Boolean = True, Optional PointAt1stItem As Boolean = True, Optional LstID As Long = NoCurrCardIDGbl)
'On Error GoTo ErrProc
'C_ItemLst.RowSource = RowSrcStr
''SaveDebugDataGSb RowSrcStr
'C_ItemLst.Requery
''If there are items in the list then
'If (C_ItemLst.ListCount > 0) And PointAt1stItem Then
'    C_ItemLst = C_ItemLst.ItemData(0)
'    If LstID <> NoCurrCardIDGbl Then 'Programmer has indicated they want the function to return the ID value if it exists.
'        'Allow zero as a legitimate item in the list
'        LstID = AssignVar2LngGFn(C_ItemLst, , True)
'    End If
'ElseIf PointAt1stItem And (LstID <> NoCurrCardIDGbl) Then 'Given there are no items in the list
'    'and the caller is expecting a value to be returned in LstID, we send back NoCurrCardIDGbl to signal a "non-value"
'    LstID = NoCurrCardIDGbl
'End If
'
'If SetFocusBln Then C_ItemLst.SetFocus
'ExitProc:
'    Exit Sub
'ErrProc:
'    DispErrMsgGSb Error$, "set the rowsource of " & C_ItemLst.Name & " to " & vbCrLf & _
'        "'" & RowSrcStr & "', or point at the first option of that row source", True, _
'        "misnamed menu or bad syntax in rowsource"
''Resume Next
'End Sub
'
'Public Sub SetRowSrc4CboGSb(CboLst As Access.ComboBox, RowSrcStr As String, _
'Optional SetFocusBln As Boolean = True, Optional PointAt1stItem As Boolean = True, _
'Optional CboID As Long = NoCurrCardIDGbl)
'On Error GoTo ErrProc
'CboLst.RowSource = RowSrcStr
'CboLst.Requery
'If (CboLst.ListCount > 0) And PointAt1stItem Then
'    CboLst = CboLst.ItemData(0)
'    If CboID <> NoCurrCardIDGbl Then 'Programmer has indicated they want the function to return the ID value if it exists.
'        'Allow zero as a legitimate item in the list
'        CboID = AssignVar2LngGFn(CboLst, , True)
'    End If
'ElseIf PointAt1stItem And (CboID <> NoCurrCardIDGbl) Then 'Given there are no items in the Cbo
'    'and the programmer is asking for the ID value, we send back a NoCurrCardIDGbl to signal a "non-value".
'    CboID = NoCurrCardIDGbl
'End If
'
'If SetFocusBln Then CboLst.SetFocus
'ExitProc:
'    Exit Sub
'ErrProc:
'    DispErrMsgGSb Error$, "set the rowsource of " & CboLst.Name & " to " & vbCrLf & _
'        "'" & RowSrcStr & "', or point at the first option of that row source", True, _
'        "misnamed menu or bad syntax in rowsource"
'End Sub
'
'Public Sub DispNoOptMsgGSb(OptInt As Integer)
'MsgBox "Sorry." & vbCrLf & _
'"You have chosen '" & OptInt & "' as an option." & vbCrLf & _
'"This is not available", _
'vbInformation, "Option Not Available"
'End Sub
'
''If you Pass a zero length string to me as the Field to focus on, I won't focus on anything at all
'Public Sub UnLoadAllGSb(obj As Object, Optional Fd2FocOnStr As String = "C_FocFdFxt")
'On Error GoTo ProcErr
'If Fd2FocOnStr <> "" Then obj(Fd2FocOnStr).SetFocus
'Dim ctl As Control
'obj.RecordSource = ""
'For Each ctl In obj.Controls
'    Select Case ctl.Properties("ControlType")
'    Case acComboBox, acListBox
'        ctl.RowSource = ""
'    Case acSubform
'        ctl.Form.RecordSource = ""
'    Case Else
'        'do nothing
'    End Select
'    Set ctl = Nothing
'Next ctl
'ProcExit:
'    Exit Sub
'ProcErr:
'    MsgBox "Sorry.  There has been an error in the program whilst 'Unloading' the " & _
'    obj.Name & vbCrLf & _
'    "Access reports the error as being: " & Error$
'    GoTo ProcExit
'End Sub
'
Public Sub DispErrMsgGSb(ErrorMsg As Variant, WhilstTryingTo As Variant, _
Optional Bring2AttnOfSysAdmin As Boolean = True, _
Optional PossiblyCausedBy As String = "", _
Optional TitleStr As String = "Error Found Please Speak to your System Administrator")
'Note that no "." is required at the end of any of the arguments being passed to this routine.
'Punctuation is catered for by this routine.
Dim HintSentence As String
Dim NotifySysAdminSentence As String
Dim MSErrMsgStr As String
'Whilst Trying To phrase is mandatory.
If PossiblyCausedBy <> "" Then
     HintSentence = vbCrLf & "Errors of this sort are possibly caused by " & PossiblyCausedBy & "."
Else
     HintSentence = ""
End If
If Bring2AttnOfSysAdmin Then
     NotifySysAdminSentence = vbCrLf & "Please bring this incident to the attention of the System Administrator."
Else
     NotifySysAdminSentence = ""
End If

If ErrorMsg <> "" Then
     MSErrMsgStr = vbCrLf & "This error has been caught by the system and you may still use this program." & vbCrLf _
     & vbCrLf & "Microsoft Access reports the error as being: " & vbCrLf & ErrorMsg
Else
     MSErrMsgStr = ""
End If

MsgBox "Sorry. " & vbCrLf & "Some difficulty has been experienced trying to " & _
WhilstTryingTo & "." & _
HintSentence & _
NotifySysAdminSentence & _
MSErrMsgStr, _
vbExclamation, TitleStr
End Sub

''If you Pass a zero length string to me as the Field to focus on, I won't focus on anything at all
'Public Sub CheckFds4ContentGSb(obj As Object, Optional Fd2FocOnStr As String = "C_FocFdFxt")
'On Error GoTo ProcErr
'If Fd2FocOnStr <> "" Then obj(Fd2FocOnStr).SetFocus
'Dim ctl As Control
'Dim Str As String
'obj.RecordSource = ""
'For Each ctl In obj.Controls
'    Select Case ctl.Properties("ControlType")
'    Case acLabel
'          'Debug.Print ctl.Caption
'        If Right(ctl.Caption, 1) = "*" Then
'
'          MsgBox "Foundone"
'        End If
'    End Select
'Next ctl
'ProcExit:
'    Exit Sub
'ProcErr:
'    MsgBox "Sorry.  There has been an error in the program whilst 'Unloading' the " & obj.Name & vbCrLf & _
'    "Access reports the error as being: " & Error$
'    GoTo ProcExit
'End Sub
'
''IDFdInt is passed by reference because it is changed
''It is the ID that you wish to set.
''Obj refers to the Form that you are on.  Pass "Me" (without quotes) to this function,
''unless you know what you are doing
''AltCtl2SetFocusStr is what we set the focus to if there are no items in the list or combo
''This is usually set to C_FormOptionsMnu
''InitOrFillFdsFlg is used when you wish to Fill or Initialise the fields.
''This requires definition of public methods "FillFds" and "InitFds" with a string parameter
''which will be called with the control that we pointed at in the calling form.
'Public Function PointAtLstGFn(ByRef IDFdInt As Long, obj As Object, CtlNameStr As String, _
'     Optional AltCtl2SetFocusStr As String = "C_FormOptionsMnu", _
'     Optional InitOrFillFdsFlg As Boolean = True, _
'     Optional FocusOnLst As Boolean = True) As Boolean
'On Error GoTo ErrProc
'Dim Lst As Object
'Set Lst = obj(CtlNameStr)
'If Not IsNull(Lst.ItemData(0)) Then
'     If IDFdInt <> NoCurrCardIDGbl And IDFdInt <> 0 Then
'          Lst = IDFdInt
'          If IsNull(Lst.Column(0)) Then
'               IDFdInt = Lst.ItemData(0)
'               Lst = IDFdInt
'          End If
'     Else
'          IDFdInt = Lst.ItemData(0)
'          Lst = IDFdInt
'     End If
'     If (InitOrFillFdsFlg) Then obj.FillFds CtlNameStr
'     If FocusOnLst Then Lst.SetFocus
'     PointAtLstGFn = True
'Else
'     IDFdInt = NoCurrCardIDGbl
'     If AltCtl2SetFocusStr <> "" Then
'          obj(AltCtl2SetFocusStr).SetFocus
'     End If
'     If (InitOrFillFdsFlg) Then obj.InitFds CtlNameStr
'     PointAtLstGFn = False
'End If
'ExitProc:
'     Exit Function
'ErrProc:
'     DispErrMsgGSb Error$, "point at the " & CtlNameStr, True, _
'     "not having properly setup a Product or Program Parameter in 'The Joiner's Mate'", _
'     "A Value Does not Seem to be Correct - Possible Error or Omission"
'End Function
'
'Public Sub LockCtrlsGSb(LockedBln As Boolean, ParamArray Ctrls())
'Dim N As Long
'For N = LBound(Ctrls) To UBound(Ctrls)
'    Ctrls(N).Locked = LockedBln: Ctrls(N).Enabled = Not LockedBln
'Next N
'End Sub
'
'Public Sub LockFdsGSb(obj As Object, BlnVal As Boolean)
'Dim Ctrl As Access.Control
'For Each Ctrl In obj.Controls
'     If Left(Ctrl.Properties("Name"), Len(obj.DataFdsPrefix)) = obj.DataFdsPrefix Then
'          Ctrl.Properties("Locked") = BlnVal
'     End If
'Next
'End Sub
'