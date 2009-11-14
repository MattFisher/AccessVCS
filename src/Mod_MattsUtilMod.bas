Option Compare Database
Option Explicit

Public Sub Test_Int2BinaryStr()
Dim i As Variant
Dim numArray() As Variant
numArray = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 15, 16, 32, 1024, 2048, -1, -2, -3, -4)
For Each i In numArray
    printHexTest (CLng(i))
Next i
End Sub

Public Sub printHexTest(number As Long)
Debug.Print number, Format(Replace(Format(Hex(number), "@@@@@@@@"), " ", "0"), "@@@@ @@@@"), Long2BinaryStr(number)
End Sub

Public Function Int2BinaryStr(aNumber As Long)
Dim bitNum As Long
Dim bitValue As Long
Dim binaryStr As String
binaryStr = "0"
'for each int in aNumber
'Debug.Print "Number: " & aNumber
If (aNumber < 0) Then binaryStr = "1"
For bitNum = 30 To 0 Step -1
    'Debug.Print "bitNum: " & bitNum & " bitValue: " & bitValue
    'Debug.Print "2^bitNum: " & 2 ^ bitNum
    'Debug.Print "2^bitNum Imp aNumber: " & (2 ^ bitNum Imp aNumber)
    'Debug.Print "aNumber or 2^bitNum : " & (aNumber Or 2 ^ bitNum)
    If (aNumber And 2 ^ bitNum) Then
        binaryStr = binaryStr & "1"
    Else
        binaryStr = binaryStr & "0"
    End If
    If (bitNum > 0) And (bitNum Mod 4 = 0) Then
        binaryStr = binaryStr & " "
    End If
Next bitNum
Int2BinaryStr = binaryStr
End Function


Public Function Long2BinaryStr(aNumber As Long)
Dim hexString As String
hexString = Replace(Format(Hex(aNumber), "@@@@@@@@"), " ", "0")
Dim i As Integer
Dim AStr As String
For i = 1 To 8
    Select Case UCase(Mid(hexString, i, 1))
        Case "0": AStr = AStr & "0000 "
        Case "1": AStr = AStr & "0001 "
        Case "2": AStr = AStr & "0010 "
        Case "3": AStr = AStr & "0011 "
        Case "4": AStr = AStr & "0100 "
        Case "5": AStr = AStr & "0101 "
        Case "6": AStr = AStr & "0110 "
        Case "7": AStr = AStr & "0111 "
        Case "8": AStr = AStr & "1000 "
        Case "9": AStr = AStr & "1001 "
        Case "A": AStr = AStr & "1010 "
        Case "B": AStr = AStr & "1011 "
        Case "C": AStr = AStr & "1100 "
        Case "D": AStr = AStr & "1101 "
        Case "E": AStr = AStr & "1110 "
        Case "F": AStr = AStr & "1111 "
    End Select
Next i
Long2BinaryStr = Mid(AStr, 1, Len(AStr) - 1) 'Get rid of the last space on the end
End Function


'Awwww these don't work in Access! Disappointing.
'
'Public Sub CheckClipboardGSb(msg As String)
'Dim clipData As New DataObject
'clipData.GetFromClipboard
'Debug.Print msg & " Clipboard: [" & clipData.GetText(1) & "]"
'End Sub
'
'Public Sub Save2ClipboardGSb(Number As Integer)
'Dim clipData As New DataObject
'clipData.SetText CStr(Number)
'clipData.PutInClipboard
''Debug.Print "Saved to Clipboard: [" & Number & "]"
'End Sub
'
'Public Function GetFromClipboardGFn() As Integer
'Dim clipData As New DataObject
'Dim TempStr As String
'clipData.GetFromClipboard
'TempStr = clipData.GetText(1)
'If IsNumeric(TempStr) Then GetFromClipboard = CInt(TempStr)
''Debug.Print "Got from Clipboard: [" & TempStr & "]"
'End Function

Public Sub PrintArrayGSb(myArray() As String)
Dim i As Integer
For i = 0 To UBound(myArray)
    Debug.Print i, "[" & myArray(i) & "]"
Next i
End Sub

Public Sub DeleteFolderIfThereGSb(aFolder As String)
'Be careful when using this - it doesn't care if there is anything in the folder,
'It just kills it.
If GetFSO.FolderExists(aFolder) Then
    GetFSO.DeleteFolder (aFolder)
End If
End Sub

Public Function CheckAndBuildFolderGFn(aFolder As String) As Boolean
On Error GoTo ErrProc
CheckAndBuildFolderGFn = False
If (aFolder = "") Then
    CheckAndBuildFolderGFn = False
    Exit Function
Else
    If (Dir(aFolder, vbDirectory) <> "") Then
        CheckAndBuildFolderGFn = True
        Exit Function
    Else
        If (CheckAndBuildFolderGFn(GetFSO.GetParentFolderName(aFolder))) Then
            MkDir aFolder
            CheckAndBuildFolderGFn = True
        End If
    End If
End If

ExitProc:
Exit Function

ErrProc:
CheckAndBuildFolderGFn = False
End Function

Public Function FileIsChangedAndNewerGFn(newFile As File, oldFile As File) As Boolean
On Error GoTo ErrProc

Dim oldFileName As String, newFileName As String
Dim oldFileChunk As String, newFileChunk As String
Dim oldFileNumber As Integer, newFileNumber As Integer
Dim filesAreDifferent As Boolean
filesAreDifferent = False
Dim newFileIsNewer As Boolean
newFileIsNewer = True
Dim StPos As Long
Dim chunkSize As Long
chunkSize = 1024 ' One kilobyte at a time
oldFileChunk = String(chunkSize, " ")
newFileChunk = String(chunkSize, " ")

If newFile.DateCreated > oldFile.DateCreated Then
    newFileIsNewer = True
    'This will most likely always be the case
    ' we just need to check if they're different
Else
    'Trying to overwrite a newer file - WTF?
    newFileIsNewer = False
End If

If newFile.Size <> oldFile.Size Then
    filesAreDifferent = True
    GoTo ExitProc
Else
    ' Same size, different timestamps. Bugger.
    ' Open both files for binary access.
    oldFileNumber = FreeFile
    Open oldFile.Path For Binary Access Read As #oldFileNumber
    newFileNumber = FreeFile
    Open newFile.Path For Binary Access Read As #newFileNumber
    StPos = 1
    ' Read both files a chunk at a time using the Get statement.
    While oldFileChunk = newFileChunk _
          And StPos <= newFile.Size
        Get #oldFileNumber, , oldFileChunk
        Get #newFileNumber, , newFileChunk
        If (oldFileChunk <> newFileChunk) Then
            filesAreDifferent = True
            GoTo ExitProc
        Else
            StPos = StPos + chunkSize
        End If
        'Debug.Print oldFileChunk
    Wend
    Close #oldFileNumber, #newFileNumber ' Close files.
End If

ExitProc:
    FileIsChangedAndNewerGFn = filesAreDifferent And newFileIsNewer
    Exit Function
ErrProc:

End Function

Private Sub Test_FileIsChangedAndNewerGFn()

'Setup
Dim Test_OldFileName As String, Test_NewFileName_SameButNewer As String
Test_OldFileName = "G:\repos\BAE\outlook\src\GlobVars.bas"
Test_NewFileName_SameButNewer = "G:\repos\BAE\outlook\src\Copy of GlobVars.bas"

'Tests
Debug.Print
Debug.Print "New file is changed and newer?: " & _
            FileIsChangedAndNewerGFn( _
                GetFSO.GetFile(Test_NewFileName_SameButNewer), _
                GetFSO.GetFile(Test_OldFileName))

'Teardown

End Sub

Public Function GetShell() As Object
  Static objShell As Object
  If objShell Is Nothing Then
     Set objShell = CreateObject("Wscript.Shell")
  End If
  Set GetShell = objShell
End Function

Public Function GetFSO() As Object
  Static objFSO As Object
  If objFSO Is Nothing Then
     Set objFSO = CreateObject("Scripting.FileSystemObject")
  End If
  Set GetFSO = objFSO
End Function