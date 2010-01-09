Option Explicit
'################  COMPUTER AND NET INFO ELEMENTS ######################
Public FbDbDSNGbl As String
Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
End Type

Private Declare Function RegOpenKeyEx _
 Lib "advapi32.dll" Alias "RegOpenKeyExA" _
 (ByVal hKey As Long, ByVal lpSubKey As String, _
 ByVal ulOptions As Long, ByVal samDesired As Long, _
 phkResult As Long) As Long
Private Declare Function RegCloseKey _
 Lib "advapi32.dll" _
 (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx _
 Lib "advapi32.dll" Alias "RegQueryValueExA" _
 (ByVal hKey As Long, ByVal lpValueName As String, _
 ByVal dwReserved As Long, lpType As Long, _
 lpData As Any, lpcbData As Long) As Long
' Registry constants
Private Const dhcSuccess = 0
Private Const dhcRegMaxDataSize = 2048
Private Const dhcRegNone = 0
Private Const dhcRegSz = 1
Private Const dhcRegExpandSz = 2
Private Const dhcRegBinary = 3
Private Const dhcRegDWord = 4
Private Const dhcRegDWordLittleEndian = 4
Private Const dhcRegDWordBigEndian = 5
Private Const dhcRegLink = 6
Private Const dhcRegMultiSz = 7
Private Const dhcRegResourceList = 8
Private Const dhcRegFullResourceDescriptor = 9
Private Const dhcRegResourceRequirementsList = 10
Private Const dhcRegOptionReserved = 0
Private Const dhcRegOptionNonVolatile = 0
Private Const dhcRegOptionVolatile = 1
Private Const dhcRegOptionCreateLink = 2
Private Const dhcRegOptionBackupRestore = 4
Private Const dhcReadControl = &H20000
Private Const dhcKeyQueryValue = &H1
Private Const dhcKeySetValue = &H2
Private Const dhcKeyCreateSubKey = &H4
Private Const dhcKeyEnumerateSubKeys = &H8
Private Const dhcKeyNotify = &H10
Private Const dhcKeyCreateLink = &H20
Private Const dhcKeyRead = dhcKeyQueryValue + dhcKeyEnumerateSubKeys + _
 dhcKeyNotify + dhcReadControl
Private Const dhcKeyWrite = dhcKeySetValue + dhcKeyCreateSubKey + dhcReadControl
Private Const dhcKeyExecute = dhcKeyRead
Private Const dhcKeyAllAccess = dhcKeyQueryValue + dhcKeySetValue + _
 dhcKeyCreateSubKey + dhcKeyEnumerateSubKeys + _
 dhcKeyNotify + dhcKeyCreateLink + dhcReadControl
Public Const dhcHKeyClassesRoot = &H80000000
Public Const dhcHKeyCurrentUser = &H80000001
Public Const dhcHKeyLocalMachine = &H80000002
Public Const dhcHKeyUsers = &H80000003
Public Const dhcHKeyPerformanceData = &H80000004

'####################### Screen Res Elements ##################################

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const ENUM_CURRENT_SETTINGS As Integer = -1

Public WidthTest As Long
Public HeightTest As Long

Private Declare Function EnumDisplaySettings Lib "User32" _
    Alias "EnumDisplaySettingsA" _
    (ByVal lpszDeviceName As Long, _
    ByVal iModeNum As Long, _
    ByRef lpDevMode As Any) As Boolean

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'################### GENERAL ELEMENTS ######################################
Dim SQLStr As String

Public Function GetValFmRegistry(RegHive As Long, KeyToRetrieve As String, ValueToRetrieve As String, strBuffer As String) As Boolean
Dim hKeyHandle As Long
Dim lngResult As Long
Dim cb As Long
strBuffer = Space(255)
If RegOpenKeyEx(RegHive, KeyToRetrieve, 0&, dhcKeyQueryValue, hKeyHandle) = dhcSuccess Then
    If RegQueryValueEx(hKeyHandle, ValueToRetrieve, 0&, dhcRegSz, ByVal strBuffer, cb) = dhcSuccess Then
        GetValFmRegistry = True
    Else
        GetValFmRegistry = (RegQueryValueEx(hKeyHandle, ValueToRetrieve, 0&, dhcRegSz, ByVal strBuffer, cb) = dhcSuccess)
    End If
    RegCloseKey hKeyHandle
Else
    GetValFmRegistry = False
End If
If GetValFmRegistry = True Then
    strBuffer = Left(strBuffer, cb - 1)
Else
    strBuffer = ""
End If
End Function

Public Function GetLogonNameGFn() As String
Dim lngResult As Long
Dim strBuffer As String
If GetValFmRegistry(dhcHKeyLocalMachine, "Network\Logon", "username", strBuffer) Then
    GetLogonNameGFn = strBuffer
ElseIf GetValFmRegistry(dhcHKeyCurrentUser, "Software\Microsoft\Windows\CurrentVersion\Explorer", "Logon User Name", strBuffer) Then
    GetLogonNameGFn = strBuffer
Else
    GetLogonNameGFn = "Unable To Determine"
End If
End Function

Public Function GetComputerNameGFn() As String
Dim lngResult As Long
Dim strBuffer As String
If GetValFmRegistry(dhcHKeyLocalMachine, "System\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName", strBuffer) Then
    GetComputerNameGFn = strBuffer
Else
    GetComputerNameGFn = "Unable To Determine"
End If
End Function

Public Function GetDesktopPathGFn(ByRef CurrDesktopPath As String) As Boolean
Dim lngResult As Long
Dim strBuffer As String
If GetValFmRegistry(dhcHKeyCurrentUser, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop", strBuffer) Then
    CurrDesktopPath = FormatDirPathGFn(strBuffer) 'Ensures there is a "\" at the end of the string
    GetDesktopPathGFn = True
Else
    CurrDesktopPath = "Unable To Determine desktop path"
    GetDesktopPathGFn = False
End If
End Function

Public Function GetDBDSNsGFn(ByRef DSNofMainDB As String, ByRef DSNofBkUpDB As String) As Boolean
On Error GoTo ErrProc
Dim RecSet As DAO.Recordset
GetDBDSNsGFn = False
SQLStr = "SELECT LTDBPath, LTDBPurpose " & _
     "FROM MDB_zzCSTbl ICST " & _
     "INNER JOIN zzLink2ClientsDBLstTbl L2CT ON ICST.CLIENTNAME4LINKING = L2CT.LTClientName "
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        RecSet.FindFirst "UCase(LTDBPurpose) = 'MAINDB'"
        If Not RecSet.NoMatch Then
            DSNofMainDB = RecSet("LTDBPath")
        Else
            DSNofMainDB = "DSNofMainDB not Specified"
            GetDBDSNsGFn = False
        End If
        RecSet.MoveFirst
        RecSet.FindFirst "UCase(LTDBPurpose) = 'BACKUPDB'"
        If Not RecSet.NoMatch Then
            DSNofBkUpDB = RecSet("LTDBPath")
        Else
            DSNofBkUpDB = "DSNofBkUpDB not Specified"
            GetDBDSNsGFn = False
        End If
    Else
        GetDBDSNsGFn = False
    End If
RecSet.Close
ExitProc:
    Exit Function
ErrProc:
    GetDBDSNsGFn = False
    DSNofMainDB = "DSNofMainDB not found because " & Error$
    DSNofBkUpDB = "DSNofBkUpDB not found because " & Error$
End Function

Public Function GetMainDBDSNGFn(ByRef MainDB_DSN As String) As Boolean
Dim RecSet As DAO.Recordset
SQLStr = "SELECT LTDBPath " & _
     "FROM MDB_zzCSTbl ICST " & _
     "INNER JOIN zzLink2ClientsDBLstTbl L2CT ON ICST.CLIENTNAME4LINKING = L2CT.LTClientName " & _
     "WHERE (UCase([LTDBPurpose]) = 'MAINDB');"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
     If Not RecSet.EOF Then
          RecSet.MoveFirst
          MainDB_DSN = RecSet("LTDBPath")
          GetMainDBDSNGFn = True
     Else
          MainDB_DSN = "MainDB_DSN Not Specified"
          GetMainDBDSNGFn = False
     End If
RecSet.Close
End Function

Public Function GetBkUpDBDSNGFn(ByRef BkupDB_DSN As String) As Boolean
Dim RecSet As DAO.Recordset
SQLStr = "SELECT LTDBPath " & _
     "FROM MDB_zzCSTbl ICST " & _
     "INNER JOIN zzLink2ClientsDBLstTbl L2CT ON ICST.CLIENTNAME4LINKING = L2CT.LTClientName " & _
     "WHERE (UCase([LTDBPurpose]) = 'BACKUPDB');"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
     If Not RecSet.EOF Then
          RecSet.MoveFirst
          BkupDB_DSN = RecSet("LTDBPath")
          GetBkUpDBDSNGFn = True
     Else
          BkupDB_DSN = "BkupDB_DSN Not Specified"
          GetBkUpDBDSNGFn = False
     End If
RecSet.Close
End Function

Public Function UserHasNoWkCensGFn(UserInits As String) As Boolean
On Error GoTo ErrProc
Dim SQLStr As String
Dim WkCenT As DAO.Recordset
SQLStr = "SELECT WkCen FROM WkCen1stTbl WHERE Mngr='" & UserInits & "'"
Set WkCenT = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
     If Not WkCenT.EOF Then
          UserHasNoWkCensGFn = False
     Else
          UserHasNoWkCensGFn = True
     End If
WkCenT.Close
ExitProc:
     Exit Function
ErrProc:
     DispErrMsgGSb Error$, "checking if user manages Work Centres", True, _
          "an anomaly in the database", _
          "Difficulty Encountered Checking if Application User Controls Work Centres."
End Function

Public Function GetValInTblGFn(IDFdValVar As Variant, IDFdNameStr As String, FdNameStr As String, TblNameStr As String, _
Optional FailNullStr As String = "NullVal", Optional FailNoRecStr As String = "NoRec") As String
'Finds a value held in table, TblNameStr, in a field, FdNameStr, for a record with an IDFd called IDFdNameStr
'where IDFdNameStr has a value of IDFdValVar
'Can handle IDFdValVar being null, a number or a string
'Sends back a string.
On Error GoTo ErrProc
     Dim SQLStr As String
     Dim RecSet As DAO.Recordset
     If Not IsNull(IDFdValVar) Then
          If IsNumeric(IDFdValVar) Then 'A number - construct SQLStr for a number in IDFdValVar
               SQLStr = "SELECT " & FdNameStr & " " & _
               "FROM " & TblNameStr & " " & _
               "WHERE " & IDFdNameStr & "=" & IDFdValVar & ";"
          Else 'A string so put IDFdValVar in quotes, ie, 'IDFdValVar'
               SQLStr = "SELECT " & FdNameStr & " " & _
               "FROM " & TblNameStr & " " & _
               "WHERE " & IDFdNameStr & "='" & IDFdValVar & "';"
          End If
               Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
                    If Not RecSet.EOF Then
                         RecSet.MoveFirst
                         If IsNumeric(RecSet(FdNameStr)) Then
                              GetValInTblGFn = Trim(Str(RecSet(FdNameStr)))
                         Else
                              GetValInTblGFn = RecSet(FdNameStr)
                         End If
                    Else
                         GetValInTblGFn = FailNoRecStr
                    End If
               RecSet.Close
     Else
          GetValInTblGFn = FailNullStr
     End If
ExitProc:
     Exit Function
ErrProc:
     DispErrMsgGSb Error$, "fetch the contractor reference that is used by your Accounting System", True, _
     "not having the table for the reference numbers properly set up", _
     "Problem Getting Accounting Systems Reference Number"
End Function

'################# SCREEN RES PROCEDURES ############################

Public Sub GetResGSb(ByRef Width As Long, ByRef Height As Long)
    Dim DevM As DEVMODE
    EnumDisplaySettings 0&, ENUM_CURRENT_SETTINGS, DevM
    Width = DevM.dmPelsWidth
    Height = DevM.dmPelsHeight
End Sub

Private Function FrontEndVersionNoFn() As String
Dim VRT As DAO.Recordset
Set VRT = CurrentDb.OpenRecordset("zzVersionTbl", dbOpenDynaset)
    If Not VRT.EOF Then
         VRT.MoveFirst
         FrontEndVersionNoFn = Trim(UCase(AssignVar2StrGFn(VRT("BackEndVersionNo"), "0000000000"))) 'Different to BE if not present
    Else
         FrontEndVersionNoFn = "0000000000"
    End If
VRT.Close
End Function

Private Function BacKEndVersionNoFn() As String
Dim RecSet As DAO.Recordset
Set RecSet = CurrentDb.OpenRecordset("SELECT DISTINCTROW BacKEndVersionNo FROM MDB_zzCSTbl", dbOpenSnapshot)
    If Not RecSet.EOF Then
         RecSet.MoveFirst
         BacKEndVersionNoFn = Trim(UCase(AssignVar2StrGFn(RecSet("BackEndVersionNo"), "9999999999"))) 'Different to FE if not present
    Else 'No record - set to a dummy value
         BacKEndVersionNoFn = "9999999999"
    End If
RecSet.Close
End Function

Public Function FEVersionNoMatchesBEVersionNoGFn() As Boolean
On Error GoTo ErrProc
If (FrontEndVersionNoFn = BacKEndVersionNoFn) Then
    FEVersionNoMatchesBEVersionNoGFn = True
Else
    FEVersionNoMatchesBEVersionNoGFn = False
End If
ExitProc:
    Exit Function
ErrProc:
    DispErrMsgGSb Error$, "determine if the version number held in the front end " & vbCrLf & _
    "matches the version number held in the back end", True, _
    "Program not being properly set up", "ERROR DETERMINING DATABASE VERSION NUMBERS"
End Function

Public Function GetFBDBDSNGFn(DBPurposeStr As String, DataSourceNameStr As String) As Boolean
'Return Data Source Name for the Firebird Database on the basis of the Purpose for which that database is used (given in DBPurposeStr).
'If the particular type of database is not selected for linking the function returns the string "NotSelectedForLinking".
'We have done this so as to have a generic approach.  Some datbases have a main db and a backup db for safety or journaling.
On Error GoTo ErrProc
Dim RecSet As DAO.Recordset
SQLStr = "SELECT ClientDetails.LTDBPath, ClientDetails.LTDBPurpose, ClientDetails.Selected4Linking " & _
            "FROM zzLink2ClientsDBLstTbl AS ClientDetails " & _
              "INNER JOIN MDB_zzCSTbl AS ClientDBConfig ON ClientDetails.LTClientName = ClientDBConfig.CSCLIENTNAME4LINKING " & _
            "WHERE (UCase(ClientDetails.LTDBPurpose)='" & DBPurposeStr & "');"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        If RecSet("Selected4Linking") Then
            DataSourceNameStr = RecSet("LTDBPath")
            GetFBDBDSNGFn = True
        Else
            DataSourceNameStr = "NotSelectedForLinking"
            GetFBDBDSNGFn = True
        End If
    Else
        GetFBDBDSNGFn = False
        DispErrMsgGSb "", "get the " & DBPurposeStr & " DSN", True, _
        "incorrect set up of program", "Unable to get DSN - Program will now halt"
    End If
RecSet.Close
ExitProc:
    Exit Function
ErrProc:
    DataSourceNameStr = ""
    GetFBDBDSNGFn = False
    DispErrMsgGSb Error$, "obtain program details specific to this program, namely " & DBPurposeStr, True, _
    "tables not being present or not being properly filled out"
End Function

Private Function GetBE_PathAndFileNameFn() As String
Dim RecSet As DAO.Recordset
SQLStr = "SELECT PSU_MDBBEPathAndFileName " & _
            "FROM  zzProgSetUpTbl"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        GetBE_PathAndFileNameFn = AssignVar2StrGFn(RecSet("PSU_MDBBEPathAndFileName"), "No Record Found", True)
    Else
        GetBE_PathAndFileNameFn = "No Record Found"
    End If
RecSet.Close
End Function

Private Function GetWP_PathAndFileNameFn() As String
Dim RecSet As DAO.Recordset
SQLStr = "SELECT PSU_MDBWPPathAndFileName " & _
            "FROM  zzProgSetUpTbl"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        GetWP_PathAndFileNameFn = AssignVar2StrGFn(RecSet("PSU_MDBWPPathAndFileName"), "No Record Found", True)
    Else
        GetWP_PathAndFileNameFn = "No Record Found"
    End If
RecSet.Close
End Function

Private Function GetNWP_PathAndFileNameFn() As String
Dim RecSet As DAO.Recordset
SQLStr = "SELECT PSU_MDBNWPPathAndFileName " & _
            "FROM  zzProgSetUpTbl"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        GetNWP_PathAndFileNameFn = AssignVar2StrGFn(RecSet("PSU_MDBNWPPathAndFileName"), "No Record Found", True)
    Else
        GetNWP_PathAndFileNameFn = "No Record Found"
    End If
RecSet.Close
End Function

Public Function GetProgSetUpValueFn(FdName As String) As String
On Error GoTo ErrProc
Dim RecSet As DAO.Recordset
SQLStr = "SELECT " & FdName & " " & _
            "FROM  zzProgSetUpTbl"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenSnapshot)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        GetProgSetUpValueFn = AssignVar2StrGFn(RecSet(FdName), "No Record Found", True)
    Else
        GetProgSetUpValueFn = "No Record Found"
    End If
RecSet.Close
Exit Function
ErrProc:
    GetProgSetUpValueFn = "No Record Found"
End Function

Private Sub StoreAppDirPath(FdVal As String)
On Error GoTo ErrProc
Dim RecSet As DAO.Recordset
SQLStr = "SELECT PSU_AppDirPath " & _
            "FROM  zzProgSetUpTbl"
Set RecSet = CurrentDb.OpenRecordset(SQLStr, dbOpenDynaset)
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        RecSet.Edit
            RecSet("PSU_AppDirPath") = FormatDirPathGFn(FdVal)
        RecSet.Update
    End If
RecSet.Close
Exit Sub
ErrProc:
    DispErrMsgGSb Error$, "save " & FdVal & " into the zzProgSetupTbl"
End Sub

Public Sub GetProgramFilesDetailsGSb(ByRef RelinkingRequired As Boolean)
Dim StoredApplicationPath As String
Dim ApplicationPathAndFileName As String
Dim Posn As Variant
'Determine the program's path and file name.
'AppNameGbl
'AppDirPathGbl
'MDB_NWP_PathAndFilenameGbl
'MDB_WP_PathAndFilenameGbl
'MDB_BE_PathAndFilenameGbl
'Determine position of last delimiter.
ApplicationPathAndFileName = UCase(Application.CurrentDb.Name)
Posn = InStrRev(ApplicationPathAndFileName, "\", -1)
AppNameGbl = Mid(ApplicationPathAndFileName, Posn + 1) 'Go to the end of the string.
AppDirPathGbl = UCase(Left(ApplicationPathAndFileName, Posn)) 'Include the "\" at the end of the string.
'Deterine name of New Work Pad (always the Application Root Name + "NWP" located in the application's directory.)
Posn = InStr(1, AppNameGbl, "FE") 'Determine point of FE in Application Name
If Posn <> 0 Then
    MDB_NWP_PathAndFilenameGbl = AppDirPathGbl & Left(AppNameGbl, Posn - 1) & "NWP" & Mid(AppNameGbl, Posn + 2)
    MDB_WP_PathAndFilenameGbl = AppDirPathGbl & Left(AppNameGbl, Posn - 1) & "WP" & Mid(AppNameGbl, Posn + 2)
Else
    MDB_NWP_PathAndFilenameGbl = "Application Name has no 'FE' - Not possible to determine New Work Pad's PathAndFileName"
    MDB_WP_PathAndFilenameGbl = "Application Name has no 'FE' - Not possible to determine Old Work Pad's PathAndFileName"
End If
StoredApplicationPath = UCase(FormatDirPathGFn(GetProgSetUpValueFn("PSU_AppDirPath")))

If StoredApplicationPath <> AppDirPathGbl Then
    'We will store the AppDirPathGbl later in SavePathAndFileNameDetails of zzReplaceWPFullForm.
    RelinkingRequired = True
Else
    RelinkingRequired = False
End If
MDB_BE_PathAndFilenameGbl = GetBE_PathAndFileNameFn
'MDB_NWP_PathAndFilenameGbl = GetWP_PathAndFileNameFn
'MDB_WP_PathAndFilenameGbl = GetNWP_PathAndFileNameFn
End Sub

Public Function GetFileDatesGFn(PathAndFileName As String, ByRef DateCreatedDte As Date, ByRef DateLastModifiedDte As Date) As Boolean
Dim oFSO As Object
Dim aFile As Object
If FileExistsGFn(PathAndFileName) Then
    GetFileDatesGFn = True
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set aFile = oFSO.GetFile(PathAndFileName)
    DateLastModifiedDte = aFile.DateLastModified
    DateCreatedDte = aFile.DateCreated
Else
    GetFileDatesGFn = False
    DateLastModifiedDte = 0
    DateCreatedDte = 0
End If
End Function