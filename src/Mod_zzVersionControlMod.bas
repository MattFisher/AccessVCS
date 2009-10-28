'' Matt's Version Control Module (Outlook 2003 Version)
'' Author: Matt Fisher <mrpfisher@gmail.com>
'' Created: 10 Oct 2009
'' Last Modified: 27 Oct 2009
''  _______________________________________________________________________
'' |                                                                       |
'' | This is a module to provide TortoiseSVN and TortoiseGit integration   |
'' | in the Outlook 2003 VBIDE.                                            |
'' |  Yes, it's a kludge, but what a kludge!                               |
'' |_______________________________________________________________________|
''
'' TODO: Figure out the best way to handle not removing the VC Code, but still updating it.
'
'Option Explicit
'
''Modify the following constants to point to your project directories
'Const SvnProjectFolder = "C:\Documents and Settings\Matt\My Documents\Projects\BAE\outlook"
'Const GitProjectFolder = "G:\repos\MattsVCS\MattsVCS-Outlook"
'Const SourceFolder = "\src"
'Const RemoteGitRepoUrl2FetchFrom = ""
'Const TempFolder = "\OutlookVCTemp"
'
''Modify the following constant to point to the directory where 'Application Data' sits
'Const AppDataBaseFolder = "C:\Documents and Settings\Matt\"
''These should not need to be changed.
'Const VbaProjectRelativePathAndFilename = "Application Data\Microsoft\Outlook\VbaProject.OTM"
'Const VbaProjectPathAndFilename = AppDataBaseFolder & VbaProjectRelativePathAndFilename
'
''Uncomment the appropriate line to control which system files are imported from
'Const importAfterRemove = "git"
''Const importAfterRemove = "svn"
''Const importAfterRemove = ""
'
'Dim lastItem As String
'Dim newItemStr As String
'Dim finished As Boolean
'Dim MyData As New DataObject
'Dim modCount, formCount, classCount As Integer
'Dim fileList() As String
'Dim fileNum As Integer
'Dim itemNum As Integer
'
'Private Sub SaveVbaProject2Git()
'CopyVbaProject2Folder GitProjectFolder
'Commit2Git GitProjectFolder & "\VbaProject.OTM"
'End Sub
'
'Private Sub SaveVbaProject2Svn()
'CopyVbaProject2Folder SvnProjectFolder
'Commit2Subversion SvnProjectFolder & "\VbaProject.OTM"
'End Sub
'
'Private Sub Test_CopyVbaProject()
''FileCopy VbaProjectPathAndFilename, VbaProjectPathAndFilename & ".bak"
'GetFSO.CopyFile VbaProjectPathAndFilename, VbaProjectPathAndFilename & ".bak"
'End Sub
'
'Private Sub CopyVbaProject2Folder(toFolder As String)
'SendKeys "^s", True 'Save project
'If Dir(VbaProjectPathAndFilename) <> "" Then
'    GetFSO.CopyFile VbaProjectPathAndFilename, toFolder & "\VbaProject.OTM", True
'    'CopyFile doesn't work - is there any way to copy VbaProject.OTM while in the Outlook VBIDE?
'    'FileSystemObject.CopyFile seems to work fine.
'Else
'    'How can you be running Outlook VBIDE without a VbaProject.OTM file?
'    MsgBox VbaProjectPathAndFilename & vbCrLf & "was not found." & vbCrLf & _
'           "Please check that your AppDataBaseFolder points to the directory" & vbCrLf & _
'           "your ""Application Data"" folder is in.", , _
'           "VbaProject.OTM Not Found"
'End If
'End Sub
'
'Public Sub ExportAndCommitAll2Svn()
'ExportItems SvnProjectFolder & SourceFolder
'CopyVbaProject2Folder SvnProjectFolder
'Commit2Subversion SvnProjectFolder & SourceFolder & "*" & SvnProjectFolder & "\VbaProject.OTM"
'End Sub
'
'Public Sub ExportAndCommitAll2Git()
'ExportChangedItems GitProjectFolder & SourceFolder
'CopyVbaProject2Folder GitProjectFolder
'Commit2Git GitProjectFolder & SourceFolder & "*" & GitProjectFolder & "\VbaProject.OTM"
'End Sub
'
'Private Sub ExportChangedItems(srcFolder As String)
'On Error GoTo ErrProc
'Dim tempPath As String, tempFilename As Variant
'Dim tempFile As File, oldFile As File
'Dim fso As Object
'Set fso = GetFSO
'tempPath = srcFolder & TempFolder
'If Not CheckAndBuildFolderGFn(tempPath) Then
'    'Error - couldn't create a temp directory!
'    Exit Sub
'End If
'ExportItems tempPath
'Debug.Print "Started File Comparisons at " & Now()
'CreateFileList tempPath, fileList
''Check which files in TempFolder have changed wrt those in SourceFolder
''Copy changed files to srcFolder, overwriting old versions.
'For Each tempFilename In fileList
'    If tempFilename <> "" Then
'        Set tempFile = fso.GetFile(tempPath & "\" & tempFilename)
'        If (Dir(srcFolder & "\" & tempFilename) <> "") Then
'            Set oldFile = fso.GetFile(srcFolder & "\" & tempFilename)
'            If FileIsChangedAndNewerGFn(tempFile, oldFile) Then
'                'Overwrite old with new
'                'Kill oldFile.Path
'                Debug.Print " Changed!: [" & tempFilename & "]"
'                fso.CopyFile tempFile.Path, oldFile.Path, True
'            Else
'                Debug.Print "Unchanged: [" & tempFilename & "]"
'            End If
'        Else
'            'If the 'old' file doesn't exist, the new one should be added.
'            Debug.Print "     New!: [" & tempFilename & "]"
'            fso.CopyFile tempFile.Path, srcFolder & "\" & tempFile.Name, True
'        End If
'    End If
'
'Next tempFilename
'
'fso.DeleteFolder srcFolder & TempFolder
'
'Debug.Print "Comparisons complete at " & Now() & "!"
'
'ExitProc:
'Exit Sub
'ErrProc:
'    'Error!
'    MsgBox Error$
'    'Resume Next
'End Sub
'
'Public Sub UpdateFromSvnAndImportAll()
'UpdateFromSubversion SvnProjectFolder & SourceFolder
'If (MsgBox("Click OK when Subversion has finished updating, or click cancel not to update the project." & vbCrLf & _
'           "This will remove all Forms, Modules and Classes from the current project and then re-import them from the given directory.", _
'           vbOKCancel, "Ready to refresh project") = vbOK) Then
'    RemoveItems 'Note that this will remove all items then
'                're-import them if importAfterRemove is true
'Else
'    'Do something else?
'End If
'End Sub
'
''Not Implemented Yet
''Public Sub FetchFromGitAndImportAll()
''FetchFromGit GitProjectFolder & SourceFolder, RemoteGitRepoUrl2FetchFrom
''If (MsgBox("Click OK when Git has finished updating, or click cancel not to update the project." & vbCrLf & _
''           "This will remove all Forms, Modules and Classes from the current project and then re-import them from the given directory.", _
''           vbOKCancel, "Ready to refresh project") = vbOK) Then
''    RemoveItems 'Note that this will remove all items then
''                're-import them if importAfterRemove is true
''Else
''    'Do something else?
''End If
''End Sub
'
'Private Sub RemoveItems()
'MyData.Clear
'modCount = 0
'formCount = 0
'classCount = 0
'
'Debug.Print ""
'Debug.Print "--------------------------------------------------------"
'Debug.Print "Removing all Modules, Classes and Forms"
'Debug.Print "Time is " & Now()
'
'finished = False
'
'itemNum = 1
''This little conceit is to get around the fact that VB resets all
''variables (static or otherwise) between runs. Hence, you can't store
''variables and have them persist.
''The clipboard is not ideal, but it may work.
'Save2Clipboard itemNum
'RemoveItem
'
'End Sub
'
'Public Sub RemoveItem()
'
'itemNum = GetFromClipboard
'Move2Item itemNum
'newItemStr = SKGetItemName()
'If newItemStr <> "ThisOutlookSession.cls" Then
'    If Right(newItemStr, 4) = ".cls" Then classCount = classCount + 1
'    If Right(newItemStr, 4) = ".bas" Then modCount = modCount + 1
'    If Right(newItemStr, 4) = ".frm" Then formCount = formCount + 1
'    Move2Item itemNum
'    Debug.Print "Removing : [" & newItemStr & "]"
'
'    Save2Clipboard itemNum
'    SendKeys "%frn{F5}ThisOutlookSession.removeItem{ENTER}" ' File -> Remove Item -> No(don't export first)
'    'Note - 'Remove Item' is disabled for ThisOutlookSession!
'    'Using 'true' to wait for processing to complete doesn't work, because
'    'Modules can't be removed while code is running.
'
'    'This is a genius hack, if I do say so myself.
'    'By running this command just before the end of the subroutine,
'    'The VBIDE returns control to the user before re-starting this subroutine.
'    'This allows the last few SendKeys commands to be processed outside
'    'of the code loop - as if there were no subroutine running.
'    'This is important because the VBIDE does not allow items to be removed
'    'from the project while code is running, or in break mode.
'    'Hence, this routine allows items to be removed programatically.
'
'Else
'    If itemNum = 1 Then
'        Debug.Print " Skipping: [ThisOutlookSession.cls]"
'        itemNum = 2
'        Save2Clipboard itemNum
'        SendKeys "{F5}ThisOutlookSession.RemoveItem{ENTER}"
'    Else
'        'Finished!
'        Debug.Print "Removal Complete at " & Now() & "!"
'        Debug.Print classCount, " classes"
'        Debug.Print modCount, " modules"
'        Debug.Print formCount, " forms"
'        Debug.Print "--------------------------------------------------------"
'        Debug.Print ""
'
'        If importAfterRemove = "git" Then
'            ImportItems GitProjectFolder & SourceFolder
'        ElseIf importAfterRemove = "svn" Then
'            ImportItems SvnProjectFolder & SourceFolder
'        Else
'            'Do nothing
'        End If
'
'    End If
'End If
'End Sub
'
'Private Sub ClearImmediateWindow()
'SendKeys "^g^a{DEL}"
'End Sub
'
'Private Sub ImportItems(fromFolder As String)
'Dim currentFile As String
'fileNum = 0
'
'Debug.Print ""
'Debug.Print "--------------------------------------------------------"
'Debug.Print "Started Import at " & Now()
'Debug.Print "Importing files from " & fromFolder
'
'CreateFileList fromFolder, fileList
'currentFile = fileList(fileNum)
'While currentFile <> ""
'    Debug.Print "Importing: " & currentFile
'    importFile (fromFolder & "\" & currentFile)
'    fileNum = fileNum + 1
'    currentFile = fileList(fileNum)
'Wend
'
'Debug.Print "Import complete at " & Now() & "!"
'Debug.Print classCount, " classes"
'Debug.Print modCount, " modules"
'Debug.Print formCount, " forms"
'Debug.Print "--------------------------------------------------------"
'Debug.Print ""
'
'End Sub
'
'Private Sub importFile(pathAndFilename As String)
'SendKeys "^m" & pathAndFilename & "{ENTER}", True
'End Sub
'
'Private Sub Test_CheckAndBuildFolderGFn()
'
'Dim Test_ExistentDir As String
'Dim Test_MalformedDir As String
'Dim Test_NonExistentDir As String
'Dim Test_NonExistentParentDir As String
'Dim Test_NonExistentDirInParentDir As String
'Test_ExistentDir = "C:\VCTest"
'Test_MalformedDir = "blahBlah"
'Test_NonExistentDir = "C:\notThere"
'Test_NonExistentParentDir = "C:\alsoNotThere"
'Test_NonExistentDirInParentDir = "C:\alsoNotThere\againAlsoNotThere"
'
''Setup
'If Not GetFSO.FolderExists(Test_ExistentDir) Then
'    MkDir (Test_ExistentDir)
'End If
'DeleteFolderIfThere Test_NonExistentDir
'DeleteFolderIfThere Test_NonExistentParentDir
'
''Tests
'Debug.Print "Malformed dir result: " & CheckAndBuildFolderGFn(Test_MalformedDir)
'Debug.Print "Existent dir result: " & CheckAndBuildFolderGFn(Test_ExistentDir)
'Debug.Print "NonExistent dir result: " & CheckAndBuildFolderGFn(Test_NonExistentDir)
'Debug.Print "NonExistentParent dir result: " & CheckAndBuildFolderGFn(Test_NonExistentDirInParentDir)
'MsgBox "Check if the folders are there!"
'
''Teardown
'DeleteFolderIfThere Test_ExistentDir
'DeleteFolderIfThere Test_NonExistentDir
'DeleteFolderIfThere Test_NonExistentParentDir
'
'End Sub
'
'Private Sub ExportItems(toFolder As String)
''Dim itemNum As Integer
'Dim lastItemStr, newItemStr As String
'MyData.Clear
'modCount = 0
'formCount = 0
'classCount = 0
'
'Debug.Print ""
'Debug.Print "--------------------------------------------------------"
'Debug.Print "Started Export at " & Now()
'Debug.Print "Exporting files to " & toFolder
'Debug.Print "Deleting files currently in folder..."
'
'CheckAndBuildFolderGFn toFolder
'If (Dir(toFolder & "\*") <> "") Then Kill toFolder & "\*"
'
'finished = False
'lastItemStr = ""
''SKEnsureProjectIsExpanded
'
'itemNum = 0
'While Not finished
'    itemNum = itemNum + 1
'    Move2Item itemNum
'    newItemStr = SKGetItemName()
'    If (newItemStr <> lastItemStr) Then
'        If newItemStr <> "ThisOutlookSession.cls" Then
'            If Right(newItemStr, 4) = ".cls" Then classCount = classCount + 1
'            If Right(newItemStr, 4) = ".bas" Then modCount = modCount + 1
'            If Right(newItemStr, 4) = ".frm" Then formCount = formCount + 1
'            Move2Item itemNum
'            Debug.Print "Exporting: [" & newItemStr & "]"
'            SendKeys "^e{HOME}" & toFolder & "\{ENTER}", True
'        Else
'            Debug.Print " Skipping: [ThisOutlookSession.cls]"
'        End If
'    Else
'        finished = True
'    End If
'    lastItemStr = newItemStr
'Wend
'
'Debug.Print "Export Complete at " & Now() & "!"
'Debug.Print classCount, " classes"
'Debug.Print modCount, " modules"
'Debug.Print formCount, " forms"
'Debug.Print "--------------------------------------------------------"
'Debug.Print ""
'
'End Sub
'
'
'Private Sub Test_CreateFileList()
'Dim myFileList() As String
'CreateFileList GitProjectFolder & SourceFolder, myFileList
'PrintArray myFileList
'End Sub
'
'Private Sub CreateFileList(folder As String, ByRef fileList() As String)
'Dim currentFile As String
'ReDim fileList(10)
'Dim fileNum As Integer
'fileNum = 0
'modCount = 0
'formCount = 0
'classCount = 0
'
'currentFile = Dir(folder & "\*")
'Do While currentFile <> ""
'    'Debug.Print currentFile
'    If Right(currentFile, 4) = ".cls" Or _
'       Right(currentFile, 4) = ".bas" Or _
'       Right(currentFile, 4) = ".frm" Then
'        If Right(currentFile, 4) = ".cls" Then classCount = classCount + 1
'        If Right(currentFile, 4) = ".bas" Then modCount = modCount + 1
'        If Right(currentFile, 4) = ".frm" Then formCount = formCount + 1
'        'Debug.Print "Adding at " & fileNum
'        fileList(fileNum) = currentFile
'        'Expand the array if needed
'        If (fileNum = UBound(fileList)) Then
'            'Debug.Print "Resizing to " & (UBound(fileList) * 2)
'            ReDim Preserve fileList(UBound(fileList) * 2)
'        End If
'        fileNum = fileNum + 1
'    End If
'    currentFile = Dir
'Loop
'
'End Sub
'
'Private Sub Test_Import()
'importFile (SvnProjectFolder & SourceFolder & "\AttachmentMod.bas")
'End Sub
'
'
'