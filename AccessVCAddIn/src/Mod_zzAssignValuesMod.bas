Option Compare Database
Option Explicit

Public Const NoImageFilePathGbl = "NoImage.jpg"

Public Function AssignVar2PicPathGFn(PicPathVar As Variant) As String
If Not IsNull(PicPathVar) Then
     If Dir(PicPathVar) <> "" Then 'Picture actually exists
          AssignVar2PicPathGFn = PicPathVar
     Else
          If (Dir(NoImageFilePathGbl) <> "") Then
               AssignVar2PicPathGFn = NoImageFilePathGbl
          Else
               AssignVar2PicPathGFn = ""
          End If
     End If
Else
     If (Dir(NoImageFilePathGbl) <> "") Then
          AssignVar2PicPathGFn = NoImageFilePathGbl
     Else
          AssignVar2PicPathGFn = ""
     End If
End If
End Function

Public Function AssignVar2VarGFn(inVar As Variant) As Variant
'Removes zero length string condition.
Dim AssignNullFlg As Boolean
AssignNullFlg = False
Select Case varType(inVar)
Case vbString
     AssignNullFlg = ("" = Trim(inVar)) 'This is a condition being evaluated to a boolean
Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
     AssignNullFlg = (0 = inVar) 'This is a condition being evaluated to a boolean
End Select
If AssignNullFlg Then AssignVar2VarGFn = Null Else AssignVar2VarGFn = inVar
End Function

Public Function AssignVar2CurrGFn(CurrVar As Variant, Optional FailValCurr As Currency = 0#) As Currency
If Not IsNull(CurrVar) Then
     If IsNumeric(CurrVar) Then
          AssignVar2CurrGFn = CurrVar
     Else
          AssignVar2CurrGFn = FailValCurr
     End If
Else
     AssignVar2CurrGFn = FailValCurr
End If
End Function

Public Function AssignVar2SglGFn(NumVar As Variant, Optional FailValDbl As Single = 0#) As Single
If Not IsNull(NumVar) And IsNumeric(NumVar) Then
     AssignVar2SglGFn = NumVar
Else
     AssignVar2SglGFn = FailValDbl
End If
End Function

Public Function AssignVar2DblGFn(NumVar As Variant, Optional FailValDbl As Double = 0#) As Double
If Not IsNull(NumVar) And IsNumeric(NumVar) Then
     AssignVar2DblGFn = NumVar
Else
     AssignVar2DblGFn = FailValDbl
End If
End Function

Public Function AssignVar2LngGFn(LngInVar As Variant, _
Optional FailVar As Long = NoCurrCardIDGbl, _
Optional AllowZeroBln As Boolean = False) As Long
Dim LngIn As Long
If IsNull(LngInVar) Then 'We have a faulty value
     If IsNumeric(FailVar) Then 'We have specified a number in case of null
          AssignVar2LngGFn = FailVar
     Else
          AssignVar2LngGFn = NoCurrCardIDGbl
     End If
Else 'Not Null
     If IsNumeric(LngInVar) Then 'A number
        LngIn = LngInVar
        If AllowZeroBln Then 'Don't worry about checking if zero
            AssignVar2LngGFn = LngIn
        ElseIf LngIn = 0 Then 'Failure (we don't permit zeros)
            If IsNumeric(FailVar) Then
                 AssignVar2LngGFn = FailVar
            Else
                 AssignVar2LngGFn = NoCurrCardIDGbl
            End If
        Else 'Don't allow zeros and LngIn <> zero
            AssignVar2LngGFn = LngIn
        End If
    Else 'Not a number treat as a failure
        If IsNumeric(FailVar) Then 'Check if Failvar is valid number
             AssignVar2LngGFn = FailVar
        Else 'Set it to a default
             AssignVar2LngGFn = NoCurrCardIDGbl
        End If
    End If
End If
End Function

Public Function AssignStr2DBFdGFn(Str As Variant) As Variant
'Useful for inserting strings into the DB Tables where zero length strings are not tolerated.
'The string may be the contents of a txt fd on a form in which case it could be null in value.
'Hence the use of Variant.
'Takes a string and assigns it to the function.  If the string is null or zero length then the function returns Null.
'Note you should apply Trim in the argument.  In this manner, Trim is only performed once.
If Not IsNull(Str) Then 'Conditionals do not work on Nulls. Therefore this situation has to be treated separately.
    If Str <> "" Then
        AssignStr2DBFdGFn = Str
    Else
        AssignStr2DBFdGFn = Null
    End If
Else
     AssignStr2DBFdGFn = Null
End If
End Function

Public Function AssignVar2StrGFn(VarStr As Variant, _
    Optional FailStr As String = "", _
    Optional AllowSpacesBln As Boolean = False, _
    Optional AllowZeroLengthBln As Boolean = False, _
    Optional TreatAsStr As Boolean = True) As Variant
'Variant may be a number in which case we convert it to a string and then assign.
If Not IsNull(VarStr) Then
    'Check if Numeric
    If IsNumeric(VarStr) And Not TreatAsStr Then 'String a number and trim the result.
        AssignVar2StrGFn = Trim(Str(VarStr))
    Else
        If Not (AllowSpacesBln) Then
            AssignVar2StrGFn = Trim(VarStr)
        Else
            AssignVar2StrGFn = VarStr
        End If
        If Trim(VarStr) <> "" Then 'Not a zero length string
            'Most straight-forward situation.
            AssignVar2StrGFn = VarStr
        Else 'Zero Length String
            If Not AllowZeroLengthBln Then 'Treat as a failure
                If UCase(FailStr) = "NULL" Then 'Set the function to Null.
                    AssignVar2StrGFn = Null
                Else 'Set the function to the FailStr
                    AssignVar2StrGFn = FailStr
                End If
            Else 'Allow Zero Length String
                AssignVar2StrGFn = VarStr 'zero length
            End If
        End If
    End If
Else 'Null Value.  Treat according to "failure" rules.
    If UCase(FailStr) = "NULL" Then
        AssignVar2StrGFn = Null
    Else
        AssignVar2StrGFn = FailStr
    End If
End If
End Function

Public Function AssignVar2DateGFn(DateVar As Variant) As Date
On Error GoTo ErrProc
Dim DblVal As Double
'Because unusual dates like 156/10/09 can actually be classed by IsDate to be a legitimate date,
'we apply a "tolerance" test to a date to see if it will fit with the timeframe of
' 03/01/1900, a Double Precision value of 4, and 01/01/2100, with a Double Precision value of 73,051
'We will accept as legitimate values from 01/01/1900 to 01/01/2100.
'Depending on how the DateVar has failed we assign to it different values thereby providing an indication
'to the calling program as to why DateVar may have failed.
If Not IsNull(DateVar) Then 'Not Null so check it.
    If IsDate(DateVar) Then 'Seems to be a date
        DblVal = CDbl(CDate(DateVar))
        If (DblVal > 4) And (DblVal < 73051) Then
            AssignVar2DateGFn = DateVar
        Else
            AssignVar2DateGFn = CDate(1.00001) '31/12/1899 12:00:01 AM {for numerical values that are not dates, 156/10/09}
        End If
    Else
        AssignVar2DateGFn = CDate(2.00001) '1/01/1900 12:00:01 AM {for values that are not in a date format}
    End If
Else
    AssignVar2DateGFn = CDate(3.00001) '2/01/1900 12:00:01 AM {for Null Values}
End If
Exit Function
ErrProc:
    AssignVar2DateGFn = CDate(4.00001) '3/01/1900 12:00:01 AM {for undefined errors, ie, causes the compiler to exception}

End Function

Private Sub TestDate()
Dim TestDate As Date
Dim ResultDate As Date
Dim LitStr As String
Dim VarVal As String
'TestDate = Now()
LitStr = "tgif"
ResultDate = CDate(4.00001)
'      2 = 01/01/1900
' 73,051 = 01/01/2100
'AssignVar2DateGFn ("156/10/09")
End Sub

Public Function AssignVar2NumGFn(VarNum As Variant, Optional FailInt As Long = NoCurrCardIDGbl) As Variant
'We use variant because the number could be long, single, double, etc
If Not IsNull(VarNum) Then
'Check it is numeric, if so assign, if not fail it.
     If IsNumeric(VarNum) Then
          AssignVar2NumGFn = VarNum
     Else
          AssignVar2NumGFn = FailInt
     End If
Else
     AssignVar2NumGFn = FailInt 'Sets it to zero if it is a null.
End If
End Function

Public Function AssignNumVarGFn(NumVar As Variant) As Variant
If IsNull(NumVar) Then
     AssignNumVarGFn = 0
Else
     AssignNumVarGFn = NumVar
End If
End Function

Public Function AssignDBFd2NumFdGFn(DBFd As Variant, FailNumStr As String, TblNameStr As String, _
FdNameStr As String) As Variant
'If failnumstr is not numeric then it sets the field to null
If IsNull(DBFd) Then 'Then check what the failstr is
     If IsNumeric(FailNumStr) Then
          AssignDBFd2NumFdGFn = Val(FailNumStr)
     Else 'Not numeric, ie, "Null" - set the function to null.
          AssignDBFd2NumFdGFn = Null
     End If
Else
     If IsNumeric(DBFd) Then
          AssignDBFd2NumFdGFn = DBFd
     Else
          MsgBox "Error in Database. I have found a string in a field which should be numeric." & _
          "The Table in which this error has been found is " & TblNameStr & "." & vbCrLf & _
          "The Field Name is " & FdNameStr & ".", vbExclamation, "Error In Database"
          AssignDBFd2NumFdGFn = Null
     End If
End If
End Function

Public Function AssignCurrencyFd2DBFdGFn(NumVar As Variant, Optional FailInt As Currency = 0#) As Currency
If Not IsNull(NumVar) Then
     If IsNumeric(NumVar) Then
          AssignCurrencyFd2DBFdGFn = NumVar
     Else
          AssignCurrencyFd2DBFdGFn = FailInt
     End If
Else
     AssignCurrencyFd2DBFdGFn = FailInt
End If
End Function

Public Function AssignAvVar2CurrGFn(AvVar As Variant, Optional FailCurr As Currency = 0#) As Currency
If Not IsNull(AvVar) Then
    If IsNumeric(AvVar) Then
        If AvVar <> 0 Then
            AssignAvVar2CurrGFn = Format(AvVar, "#,###.####") 'Will format beyond 9,999.9999 as necessary
        Else
            AssignAvVar2CurrGFn = FailCurr
        End If
    Else
         AssignAvVar2CurrGFn = FailCurr
    End If
Else
     AssignAvVar2CurrGFn = FailCurr
End If
End Function

Public Function AssignFormFdStr2FormFd(StrVar As Variant) As Variant
'This takes a value from, for example, the column of a list and assigns it to a CboBox on same form.
If Not (IsNull(StrVar)) Then
     If (Trim(StrVar) = "") Then
          AssignFormFdStr2FormFd = Null
     Else
          AssignFormFdStr2FormFd = StrVar
     End If
Else
     AssignFormFdStr2FormFd = Null
End If
End Function

Public Function AssignStr2VarGFn(StrVar As Variant, Optional FailVar As String = "") As Variant
'Highly flexible string handling.
'Will treat null or "" as being a failure. If so it will assign FailStr to the variant.
'If FailStr is "NULL" (case insensitive) then it will set function to null value.
If Not (IsNull(StrVar)) Then
     If (Trim(StrVar) <> "") Then
          AssignStr2VarGFn = StrVar
     Else
          If (UCase(FailVar) <> "NULL") Then
               AssignStr2VarGFn = FailVar
          Else
               AssignStr2VarGFn = Null
          End If
     End If
Else
     If (UCase(FailVar) <> "NULL") Then
          AssignStr2VarGFn = FailVar
     Else
          AssignStr2VarGFn = Null
     End If
End If
End Function

Public Function AssignNum2VarGFn(NumVar As Variant, Optional FailVar As Variant = 0, _
Optional AllowZeroBln As Boolean = False) As Variant
'Highly flexible string handling.
'Will treat null or "" as being a failure. If so it will assign FailStr to the variant.
'If FailStr is "NULL" (case insensitive) then it will set function to null value.
If Not (IsNull(NumVar)) Then
     If IsNumeric(NumVar) Then
          If AllowZeroBln Then 'Don't worry about the value of the variant.
               AssignNum2VarGFn = NumVar
          Else 'Won't allow zeros so check if the value is zero
               If NumVar = 0 Then 'We don't allow zero. Set it to FailVar
                    If IsNumeric(FailVar) Then 'We set it to a numeric value
                         AssignNum2VarGFn = FailVar
                    Else 'Check if the FailVar is "NULL". If so set it to Null.
                         If UCase(FailVar) = "NULL" Then AssignNum2VarGFn = Null
                    End If
               Else 'Not zero so let it go.
                    AssignNum2VarGFn = NumVar
               End If
          End If
     Else 'It is not numeric. Set it to FailVar
          If IsNumeric(FailVar) Then 'Failvar is a number so assign it.
               AssignNum2VarGFn = FailVar
          Else 'FailVar is not a number. Check if it is "NULL". If so set subject to Null
               If UCase(FailVar) = "NULL" Then AssignNum2VarGFn = Null
          End If
     End If
Else 'NumVar is a null. Set it to the FailVar as required.
     If IsNumeric(FailVar) Then
          AssignNum2VarGFn = FailVar
     Else
          If UCase(FailVar) = "NULL" Then AssignNum2VarGFn = Null
     End If
End If
End Function

Public Function AssignVar2BlnGFn(VarBln As Variant, Optional FailBln As Boolean = False) As Boolean
If IsNull(VarBln) Then
     AssignVar2BlnGFn = FailBln
Else
     If IsNumeric(VarBln) Then
          If VarBln = -1 Then
               AssignVar2BlnGFn = True
          Else
               AssignVar2BlnGFn = False
          End If
     Else 'Not Numeric therefore False
          AssignVar2BlnGFn = FailBln
     End If
End If
End Function