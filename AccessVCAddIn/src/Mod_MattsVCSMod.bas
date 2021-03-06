'Author:    Matt Fisher
'Created:   13 May 2008

'TODO:
'   Make tables from external database files distinguishable once exported
'   Recreate external database files on import
'   Put linked tables in external files and relink to front db
'   Investigate additional DB info - startup, icon, properties, etc.

'   Figure out WTF is going on WRT no DB being open when we come into Export the first time
'   Figure out how to overwrite objects if they already exist
'   Stop database program running when opened

Option Compare Database
Option Explicit

Const FILE_EXT_MODULE As String = ".bas"
Const FILE_EXT_FORM As String = ".frm"
Const FILE_EXT_CLASS As String = ".cls"
Const FILE_EXT_REPORT As String = ".rpt"
Const FILE_EXT_MACRO As String = ".mac"
Const FILE_EXT_PAGE As String = ".adp"
Const FILE_EXT_QUERY As String = ".qry"
Const FILE_EXT_TABLE_SCHEMA As String = ".xsd"
Const FILE_EXT_TABLE_DATA As String = ".txt"
Const FILE_EXT_TABLE_COMBINED As String = ".xml"

Const PREFIX_MODULE As String = "Mod_"
Const PREFIX_FORM As String = "Frm_"
Const PREFIX_CLASS As String = "Cls_"
Const PREFIX_REPORT As String = "Rpt_"
Const PREFIX_MACRO As String = "Mac_"
Const PREFIX_PAGE As String = "Pge_"
Const PREFIX_QUERY As String = "Qry_"
Const PREFIX_TABLE As String = "Tbl_"

'Public Const TABLE_LIST_TABLENAME As String = "__TABLE_LIST__"
Public Const TABLE_LIST_TABLENAME As String = "__TABLES__BE"
Public Const TABLE_LIST_FILENAME As String = "__TABLES__.xml"
Public Const DEFAULT_EXPORT_LOC As String = _
    "C:\Documents and Settings\Matt\My Documents\Projects\test"
Public exportLocGbl As String


'These are constants for tableDefs that are confusing, if not wrong.

'Public Const dbHiddenObject As Long = &H2
'This is defined as &H1 within Access

'Public Const dbSystemObject As Long = &H80000000
'This is defined as &H80000002 within Access
'This makes some kind of sense, in that every system object
'is hidden, but it means
' (tableDef.Attributes And dbSystemObject) = true
'when tableDef.Attributes = dbHiddenObject.
            
'PictureTypes
Const Embedded = 0
Const Linked = 1

Dim app As Access.Application
Const SADebug As Boolean = True

Dim TableCount As Integer, queryCount As Integer, formCount As Integer, moduleCount As Integer
Dim macroCount As Integer, reportCount As Integer, pageCount As Integer, classCount As Integer
Dim processTables As Boolean, processQueries As Boolean, processForms As Boolean, processModules As Boolean
Dim processMacros As Boolean, processReports As Boolean, processPages As Boolean, processClasses As Boolean
Dim processProperties As Boolean, processOptions As Boolean, processRelationships As Boolean
Dim processReferences As Boolean

Private Sub ExportForm(formName As String)
Set app = Access.Application
app.SaveAsText acForm, formName, _
               exportLocGbl & formName & ".frm"
End Sub

Private Sub Test_ExportForm1()
'ExportForm "MattsVCSFrm"
ExportForm "TableSubFrm"
End Sub

Private Sub ImportForm(formName As String)
Set app = Access.Application
app.LoadFromText acForm, formName, _
                 exportLocGbl & formName & ".frm"
End Sub

Private Sub Test_ImportForm()
'ImportForm "MattsVCSFrm"
ImportForm "TableSubFrm"
End Sub

Private Sub ExportTableSchemaAsXsd(tableNameStr As String, exportLoc As String)
Set app = Access.Application
app.ExportXML objectType:=acExportTable, _
              DataSource:=tableNameStr, _
              SchemaTarget:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_SCHEMA
End Sub

Private Sub ExportTableSchemaAndDataAsXml(tableNameStr As String, exportLoc As String)
Set app = Access.Application
app.ExportXML objectType:=acExportTable, _
              DataSource:=tableNameStr, _
              DataTarget:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_COMBINED, _
              OtherFlags:=acEmbedSchema
End Sub

Private Sub ExportTableDataOnlyAsXml(tableNameStr As String, exportLoc As String)
Set app = Access.Application
app.ExportXML objectType:=acExportTable, _
              DataSource:=tableNameStr, _
              DataTarget:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_COMBINED
End Sub

Private Sub ExportTableDataAsTxt(tableNameStr As String, exportLoc As String)
Set app = Access.Application
app.DoCmd.TransferText TransferType:=acExportDelim, _
                       TableName:=tableNameStr, _
                       FileName:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_DATA, _
                       HasFieldNames:=True
End Sub

Private Sub ExportSpecialTable(tableNameStr As String, exportLoc As String)
' Copy special table
' (TABLE_LIST_TABLENAME, PROPERTY_LIST_TABLENAME, REFERENCES_TABLENAME or RELATIONS_TABLENAME)
' from codeDB to currentDB, then save it as an XML file and delete the table again.
Dim localTableNameStr As String
localTableNameStr = Left(tableNameStr, Len(tableNameStr) - 2)
DoCmd.TransferDatabase acImport, "Microsoft Access", _
    CodeDb.Name, acTable, tableNameStr, localTableNameStr, False
' Export the table to an XML File
app.ExportXML objectType:=acExportTable, _
              DataSource:=localTableNameStr, _
              DataTarget:=exportLoc & localTableNameStr & ".xml", _
              OtherFlags:=acEmbedSchema
' Delete the table again
Dim attemptCount As Integer
On Error Resume Next
DoCmd.DeleteObject acTable, localTableNameStr
While TableExistsInDbGFn(localTableNameStr)
    attemptCount = attemptCount + 1
    If (attemptCount >= 100) Then
        MsgBox "Can't deleted the table " & localTableNameStr & ", and I've tried 100 times!"
        Exit Sub
    End If
    Sleep (100)
    DoCmd.DeleteObject acTable, localTableNameStr
Wend
End Sub

Private Function ExportListedTables(exportLoc As String) As Integer
Dim TableCount As Integer
TableCount = 0
Set app = Access.Application
'exportLoc = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\test\src\"
'exportLoc = GetDBFolderNameGFn(CurrentDb) & "\" & GetFSO.GetBaseName(CurrentDb.Name) & "\src\"
If (Right(exportLoc, 1) <> "\") Then exportLoc = exportLoc & "\"
CheckAndBuildFolderGFn (exportLoc)

Log "***** Tables *****"
Debug.Print TABLE_LIST_TABLENAME
ExportSpecialTable TABLE_LIST_TABLENAME, exportLoc

Dim TableList As DAO.Recordset
Set TableList = CodeDb.OpenRecordset("SELECT * FROM " & TABLE_LIST_TABLENAME, dbOpenSnapshot)
    If Not TableList.EOF Then
        TableList.MoveFirst
        While Not TableList.EOF
            If TableList("Tbl_ContainsBinary") Then
                If TableList("Tbl_ExportSchema") Then
                    If TableList("Tbl_ExportData") Then
                        Log "Exporting " & TableList("Tbl_Name") & " (Combined XML)"
                        ExportTableSchemaAndDataAsXml TableList("Tbl_Name"), exportLoc
                        TableCount = TableCount + 1
                    Else
                        Log "Exporting " & TableList("Tbl_Name") & " (XSD Schema)"
                        ExportTableSchemaAsXsd TableList("Tbl_Name"), exportLoc
                        TableCount = TableCount + 1
                    End If
                Else
                    If TableList("Tbl_ExportData") Then
                        Log "Exporting " & TableList("Tbl_Name") & " (XML Data Only)"
                        ExportTableDataOnlyAsXml TableList("Tbl_Name"), exportLoc
                        TableCount = TableCount + 1
                    End If
                End If
            Else
                If TableList("Tbl_ExportSchema") Then
                    Log "Exporting " & TableList("Tbl_Name") & " (XSD Schema)"
                    ExportTableSchemaAsXsd TableList("Tbl_Name"), exportLoc
                    TableCount = TableCount + 1
                    If TableList("Tbl_ExportData") Then
                        Log "Exporting " & TableList("Tbl_Name") & " (TXT Data)"
                        ExportTableDataAsTxt TableList("Tbl_Name"), exportLoc
                    End If
                Else
                    If TableList("Tbl_ExportData") Then
                        Log "Exporting " & TableList("Tbl_Name") & " (TXT Data)"
                        ExportTableDataAsTxt TableList("Tbl_Name"), exportLoc
                        TableCount = TableCount + 1
                    End If
                End If
            End If
            TableList.MoveNext
        Wend
    End If
TableList.Close

Log TableCount & " Tables Exported"

Dim exportIniFile As String
exportIniFile = Dir(exportLoc & "export.ini")
If exportIniFile <> "" Then Kill exportLoc & exportIniFile

ExportListedTables = TableCount
End Function

Public Sub Test_PrintContainerNames()
Dim c As Container
For Each c In CurrentDb.Containers
    Debug.Print c.Name
Next c
End Sub

'Containers
'"DataAccessPages"
'"Databases"
'"Forms"
'"Modules"
'"Relationships"
'"Reports"
'"Scripts"
'"SysRel"
'"Tables"

'acObjectTypes
'Constant Value
'acDefault -1
'acTable 0
'acQuery 1
'acForm 2
'acReport 3
'acMacro 4
'acModule 5
'acDataAccessPage 6
'acServerView 7
'acDiagram 8
'acStoredProcedure 9
'acFunction 10

'Private Sub processDocs(docType As Integer)
'If processDocType(docType) Then
'    Set c = db.Containers(getDocTypeName(docType))
'    For Each d In c.Documents
'        If performExport Then
'            app.SaveAsText docType, d.Name, _
'                exportLocation & getDocTypePrefix(docType) & _
'                d.Name & EXPORT_FILE_EXT
'        End If
'        'TODO: implement dictionaries for these?
'        docTypeCount(docType) = docTypeCount(docType) + 1
'        CheckTimer
'    Next d
'End If
'End Sub
'

Public Sub Test_GetTableDefs()
Dim td As TableDef
Dim db As DAO.Database
Set db = CurrentDb
Set td = db.TableDefs("FB-NEW_TABLE")
End Sub

Public Sub Test_FormsHaveEmbeddedImages()
MsgBox FormsHaveEmbeddedImages()
End Sub

Public Function FormsHaveEmbeddedImages() As Boolean
FormsHaveEmbeddedImages = False
Dim c As Container
Dim d As Document
Dim db As Database
Set db = Application.CurrentDb
Dim myForm As Form
Dim currControl As Control
Set c = db.Containers("Forms")
For Each d In c.Documents
    If (d.Name <> "MattsVCSFrm") And (d.Name <> "TableSubFrm") Then
        DoCmd.OpenForm d.Name, acDesign, , , acFormPropertySettings, acHidden
        Set myForm = Forms(d.Name)
        If (myForm.PictureType = Embedded) And _
           (myForm.Picture <> "(none)") Then 'This will probably just be "(bitmap)"
            FormsHaveEmbeddedImages = True
        Else
            For Each currControl In myForm.Controls
                If (currControl.ControlType = acImage) Then
                    If ((currControl.PictureType = Embedded) And _
                       (currControl.Picture <> "")) Then
                        FormsHaveEmbeddedImages = True
                        DoCmd.Close acForm, d.Name, acSaveNo
                        GoTo ExitProc
                    End If
                End If
            Next currControl
        End If
        DoCmd.Close acForm, d.Name, acSaveNo
    End If
Next d

ExitProc:
Exit Function

ErrProc:
MsgBox Err
End Function

Public Sub test_project()
Dim db As Database
Dim myForm As AccessObject
Dim myProject
Set db = Application.CurrentDb
Set myProject = Application.CurrentProject

MsgBox db.Containers("Forms").Documents.Count & " forms found in forms Container."
MsgBox myProject.AllForms.Count & " forms found in project AllForms."
MsgBox Forms.Count & " forms found in Forms."

Dim myCollection
'Set myCollection = myProject.AllForms
'For Each myForm In myCollection
'    DoCmd.OpenForm myForm.Name, acNormal, , , acFormPropertySettings, acHidden
'    MsgBox myForm.Name & " picture type " & Forms(myForm.Name).PictureType
'Next myForm

Dim myDoc As Document
Set myCollection = db.Containers("Forms").Documents
For Each myDoc In myCollection
    DoCmd.OpenForm myDoc.Name, acDesign, , , acFormPropertySettings, acHidden
    MsgBox myDoc.Name & " picture type " & Forms(myDoc.Name).PictureType
Next myDoc

End Sub

'Exports all objects to files in exportLocGbl
Public Function ExportDatabaseObjects(exportLoc As String) As String
On Error GoTo Err_ExportDatabaseObjects

Dim db As Database
Dim td As TableDef
Dim d As Document
Dim c As Container
Dim i As Integer
Dim tableDataInXML As Boolean
tableDataInXML = False
Dim performExport As Boolean
performExport = True

TableCount = 0
queryCount = 0
formCount = 0
moduleCount = 0
macroCount = 0
reportCount = 0
pageCount = 0
classCount = 0

processTables = True
processQueries = True
processForms = True
processModules = True
processMacros = True
processReports = True
processPages = True
processProperties = True
processOptions = False
processRelationships = True
processReferences = True

Set app = Access.Application
Set db = app.CurrentDb

StartTimer

If Not (db Is Nothing) Then
    If processProperties Then
        ListAllProperties
        If processOptions Then ListAllOptions
        ExportSpecialTable PROPERTY_LIST_TABLENAME, exportLoc
    End If
    
    If processRelationships Then
        ListRelationships
        ExportSpecialTable RELATIONS_TABLENAME, exportLoc
    End If
    
    If processReferences Then
        ListReferences
        ExportSpecialTable REFERENCES_TABLENAME, exportLoc
    End If
    
    If processTables Then
        'tableCount = ListTables
        TableCount = ExportListedTables(exportLoc)
    End If
    ExportDatabaseObjects = TableCount & " tables" & vbCrLf

    If processForms Then
        Log "***** Forms *****"
        Set c = db.Containers("Forms")
        For Each d In c.Documents
            If performExport Then
                Log "Exporting Form: " & d.Name
                app.SaveAsText acForm, d.Name, exportLoc & PREFIX_FORM & d.Name & FILE_EXT_FORM
            End If
            formCount = formCount + 1
            CheckTimer
        Next d
    End If
    Log formCount & " Forms Exported"
    ExportDatabaseObjects = ExportDatabaseObjects & formCount & " forms" & vbCrLf
    
    'If processPages ...
    'pageCount & " data access pages"
    
    If processReports Then
        Log "***** Reports *****"
        Set c = db.Containers("Reports")
        For Each d In c.Documents
            If performExport Then
                Log "Exporting Report: " & d.Name
                app.SaveAsText acReport, d.Name, exportLoc & PREFIX_REPORT & d.Name & FILE_EXT_REPORT
            End If
            reportCount = reportCount + 1
            CheckTimer
        Next d
    End If
    Log reportCount & " Reports Exported"
    ExportDatabaseObjects = ExportDatabaseObjects & reportCount & " reports" & vbCrLf
    
    If processMacros Then
        Log "***** Macros *****"
        Set c = db.Containers("Scripts")
        For Each d In c.Documents
            If performExport Then
                Log "Exporting Macro: " & d.Name
                app.SaveAsText acMacro, d.Name, _
                                     exportLoc & PREFIX_MACRO & d.Name & FILE_EXT_MACRO
            End If
            macroCount = macroCount + 1
            CheckTimer
        Next d
    End If
    Log macroCount & " Macros Exported"
    ExportDatabaseObjects = ExportDatabaseObjects & macroCount & " macros" & vbCrLf
    
    If processModules Then
        Log "***** Modules *****"
        Set c = db.Containers("Modules")
        For Each d In c.Documents
            If performExport Then
                Log "Exporting Module: " & d.Name
                app.SaveAsText acModule, d.Name, _
                                     exportLoc & PREFIX_MODULE & d.Name & FILE_EXT_MODULE
            End If
            moduleCount = moduleCount + 1
            CheckTimer
        Next d
    End If
    Log moduleCount & " Modules Exported"
    ExportDatabaseObjects = ExportDatabaseObjects & moduleCount & " modules" & vbCrLf
    
    If processQueries Then
        Log "***** Queries *****"
        For i = 0 To db.QueryDefs.Count - 1
            'Skip the embedded queries
            If Left(db.QueryDefs(i).Name, 1) <> "~" Then
                If performExport Then
                    Log "Exporting Query: " & db.QueryDefs(i).Name
                    app.SaveAsText acQuery, db.QueryDefs(i).Name, _
                                         exportLoc & PREFIX_QUERY & db.QueryDefs(i).Name & FILE_EXT_QUERY
                End If
                queryCount = queryCount + 1
                CheckTimer
            End If
        Next i
    End If
    Log queryCount & " Queries Exported"
    ExportDatabaseObjects = ExportDatabaseObjects & queryCount & " queries" & vbCrLf
    
End If

Set db = Nothing
Set c = Nothing

If False Then MsgBox "All database objects have been exported as text and XML files to " & exportLoc & vbCrLf & _
       "Total time taken: " & GetTimeString(CheckTimer), _
       vbInformation

Exit_ExportDatabaseObjects:
    Exit Function
    
Err_ExportDatabaseObjects:
    MsgBox Err.number & " - " & Err.Description & vbCrLf & _
    Error$
    Resume Next

End Function

Private Function ExportChangedItems(srcFolder As String) As String
'Exports only the changed items to the given directory (so as not to confuse Git/Svn)
On Error GoTo ErrProc
Dim tempPath As String, tempFilename As Variant, oldFilename As Variant
Dim tempFile As File, oldFile As File
Dim resultStr As String
Dim TempFolder As String
TempFolder = "__TEMP__"
Dim FSO As Object
Set FSO = GetFSO
tempPath = srcFolder & TempFolder
If Not CheckAndBuildFolderGFn(tempPath) Then
    'Error - couldn't create a temp directory!
    Exit Function
End If
resultStr = ExportDatabaseObjects(tempPath & "\")
Log "Started File Comparisons"
Dim newFileList() As String
Dim oldFileList() As String
CreateFileList tempPath, newFileList
CreateFileList srcFolder, oldFileList

'Delete any 'old' versions that don't have 'new' versions.
For Each oldFilename In oldFileList
    If oldFilename <> "" Then
        If (Dir(tempPath & "\" & oldFilename) = "") Then
            Set oldFile = FSO.GetFile(srcFolder & "\" & oldFilename)
            'If the 'new' file doesn't exist, the old one should be deleted.
            Log "File Removed!:  [" & oldFilename & "]"
            Kill oldFile.Path
        End If
    End If
Next oldFilename

'Check which files in TempFolder have changed wrt those in SourceFolder
'Copy changed files to srcFolder, overwriting old versions.
For Each tempFilename In newFileList
    If tempFilename <> "" Then
        Set tempFile = FSO.GetFile(tempPath & "\" & tempFilename)
        If (Dir(srcFolder & "\" & tempFilename) <> "") Then
            Set oldFile = FSO.GetFile(srcFolder & "\" & tempFilename)
            If FileIsChangedAndNewerGFn(tempFile, oldFile) Then
                'Overwrite old with new
                'Kill oldFile.Path
                Log "File Changed!:  [" & tempFilename & "]"
                FSO.CopyFile tempFile.Path, oldFile.Path, True
            Else
                Log "File Unchanged: [" & tempFilename & "]"
            End If
        Else
            'If the 'old' file doesn't exist, the new one should be added.
            Log "File Added!:    [" & tempFilename & "]"
            FSO.CopyFile tempFile.Path, srcFolder & "\" & tempFile.Name, True
        End If
    End If

Next tempFilename

Close 'In case the copying leaves any files open
DeleteFolderIfThereGSb srcFolder & TempFolder

Debug.Print "Comparisons complete!"
ExportChangedItems = resultStr

ExitProc:
Exit Function
ErrProc:
    'Error!
    MsgBox Error$
    'Resume Next
End Function

Private Sub CreateFileList(folder As String, ByRef fileList() As String)
Dim currentFile As String
ReDim fileList(10)
Dim fileNum As Integer
Dim filePrefix As String
Dim Count_Mod As Integer
Dim Count_Frm As Integer
Dim Count_Rpt As Integer
Dim Count_Mac As Integer
Dim Count_Pge As Integer
Dim Count_Qry As Integer
Dim Count_Tbl As Integer
fileNum = 0


currentFile = Dir(folder & "\*")
Do While currentFile <> ""
    'Debug.Print currentFile
    filePrefix = Left(currentFile, 4)
    If filePrefix = PREFIX_MODULE Or _
       filePrefix = PREFIX_FORM Or _
       filePrefix = PREFIX_REPORT Or _
       filePrefix = PREFIX_MACRO Or _
       filePrefix = PREFIX_PAGE Or _
       filePrefix = PREFIX_QUERY Or _
       filePrefix = PREFIX_TABLE Or _
       currentFile = TABLE_LIST_FILENAME Or _
       currentFile = PROPERTY_LIST_FILENAME Or _
       currentFile = RELATIONS_FILENAME Or _
       currentFile = REFERENCES_FILENAME Then
        If filePrefix = PREFIX_MODULE Then Count_Mod = Count_Mod + 1
        If filePrefix = PREFIX_FORM Then Count_Frm = Count_Frm + 1
        If filePrefix = PREFIX_REPORT Then Count_Rpt = Count_Rpt + 1
        If filePrefix = PREFIX_MACRO Then Count_Mac = Count_Mac + 1
        If filePrefix = PREFIX_PAGE Then Count_Pge = Count_Pge + 1
        If filePrefix = PREFIX_QUERY Then Count_Qry = Count_Qry + 1
        If filePrefix = PREFIX_TABLE Then Count_Tbl = Count_Tbl + 1
        'Debug.Print "Adding at " & fileNum
        fileList(fileNum) = currentFile
        'Expand the array if needed
        If (fileNum = UBound(fileList)) Then
            'Debug.Print "Resizing to " & (UBound(fileList) * 2)
            ReDim Preserve fileList(UBound(fileList) * 2)
        End If
        fileNum = fileNum + 1
    End If
    currentFile = Dir
Loop

End Sub


'NOT USED - Done in VBScript
'Imports all valid text files in the importFolder to the currentDB of app.
Public Function ImportDatabaseObjects(importFolder As String, _
                                      Optional importObjects As Boolean = True) _
                                      As String
On Error GoTo Err_ImportDatabaseObjects

'Do not forget the closing back slash! ie: C:\Temp\
If (Right(importFolder, 1) <> "\") Then
    importFolder = importFolder & "\"
End If

Dim origFileName As String
Dim ucFileName As String
Dim objectType As String
Dim objectName As String
Dim dataFileName As String

TableCount = 0
queryCount = 0
formCount = 0
moduleCount = 0
macroCount = 0
reportCount = 0
pageCount = 0

processTables = True
processQueries = True
processForms = True
processModules = True
processMacros = True
processReports = True
processPages = True

origFileName = Dir(importFolder, vbNormal)
ucFileName = UCase(origFileName)

While ucFileName <> ""

    If Right(ucFileName, Len(FILE_EXT_TABLE_COMBINED)) = FILE_EXT_TABLE_COMBINED Then
        '.xml file
        If SADebug Then Debug.Print origFileName & " is a " & FILE_EXT_TABLE_COMBINED & " file"
        If ((Left(ucFileName, Len(PREFIX_TABLE)) = PREFIX_TABLE) And processTables) Then
            '"Tbl_[].xml" file
            If SADebug Then Debug.Print origFileName & " is a combined table file"
            TableCount = TableCount + 1

            If importObjects Then
                app.ImportXML importFolder & origFileName, acStructureAndData
            End If
        End If

    ElseIf Right(ucFileName, Len(FILE_EXT_TABLE_SCHEMA)) = FILE_EXT_TABLE_SCHEMA Then
        '.xsd file
        If SADebug Then Debug.Print origFileName & " is a " & FILE_EXT_TABLE_SCHEMA & " file"
        If ((Left(ucFileName, Len(PREFIX_TABLE)) = PREFIX_TABLE) And processTables) Then
            '"Tbl_[].xsd" file
            If SADebug Then Debug.Print origFileName & " is a table schema file"
            TableCount = TableCount + 1

            If importObjects Then
                app.ImportXML importFolder & origFileName, acStructureOnly
                'Also import the corresponding data file
                objectName = Mid(origFileName, Len(PREFIX_TABLE) + 1, _
                             Len(origFileName) - Len(PREFIX_TABLE) - Len(FILE_EXT_TABLE_SCHEMA))
                dataFileName = importFolder & PREFIX_TABLE & objectName & FILE_EXT_TABLE_DATA
                If GetFSO.FileExists(dataFileName) Then
                    app.DoCmd.TransferText acImportDelim, , objectName, dataFileName, True
                End If
            End If
        End If

    Else
        If (Right(ucFileName, Len(FILE_EXT_QUERY)) = FILE_EXT_QUERY) And _
               (Left(ucFileName, Len(PREFIX_QUERY)) = PREFIX_QUERY) And _
               processQueries Then
            '"Qry_[].qry" file
            If SADebug Then Debug.Print origFileName & " is a query file"
            objectName = Mid(origFileName, Len(PREFIX_QUERY) + 1, _
                         Len(origFileName) - Len(PREFIX_QUERY) - Len(FILE_EXT_QUERY))
            'Skip embedded queries
            If Left(objectName, 1) <> "~" Then
                objectType = acQuery
                queryCount = queryCount + 1
            End If

        ElseIf (Right(ucFileName, Len(FILE_EXT_MODULE)) = FILE_EXT_MODULE) And _
               (Left(ucFileName, Len(PREFIX_MODULE)) = PREFIX_MODULE) And _
               processModules Then
            '"Mod_[].bas" file
            If SADebug Then Debug.Print origFileName & " is a module file"
            objectName = Mid(origFileName, Len(PREFIX_MODULE) + 1, _
                         Len(origFileName) - Len(PREFIX_MODULE) - Len(FILE_EXT_MODULE))
            ' Don't overwrite yourself
            If (objectName <> "StandalonePorterMod") Then
                objectType = acModule
                moduleCount = moduleCount + 1
            End If

        ElseIf (Right(ucFileName, Len(FILE_EXT_FORM)) = FILE_EXT_FORM) And _
               (Left(ucFileName, Len(PREFIX_FORM)) = PREFIX_FORM) And _
               processForms Then
            '"Frm_[].frm" file
            If SADebug Then Debug.Print origFileName & " is a form file"
            objectType = acForm
            formCount = formCount + 1
            objectName = Mid(origFileName, Len(PREFIX_FORM) + 1, _
                         Len(origFileName) - Len(PREFIX_FORM) - Len(FILE_EXT_FORM))

        ElseIf (Right(ucFileName, Len(FILE_EXT_MACRO)) = FILE_EXT_MACRO) And _
               (Left(ucFileName, Len(PREFIX_MACRO)) = PREFIX_MACRO) And _
               processMacros Then
            '"Mcr_[].bas" file
            If SADebug Then Debug.Print origFileName & " is a macro file"
            objectType = acMacro
            macroCount = macroCount + 1
            objectName = Mid(origFileName, Len(PREFIX_MACRO) + 1, _
                         Len(origFileName) - Len(PREFIX_MACRO) - Len(FILE_EXT_MACRO))

        ElseIf (Right(ucFileName, Len(FILE_EXT_REPORT)) = FILE_EXT_REPORT) And _
               (Left(ucFileName, Len(PREFIX_REPORT)) = PREFIX_REPORT) And _
               processReports Then
            '"Rpt_[].rpt" file
            If SADebug Then Debug.Print origFileName & " is a report file"
            objectType = acReport
            reportCount = reportCount + 1
            objectName = Mid(origFileName, Len(PREFIX_REPORT) + 1, _
                         Len(origFileName) - Len(PREFIX_REPORT) - Len(FILE_EXT_REPORT))

        ElseIf (Right(ucFileName, Len(FILE_EXT_PAGE)) = FILE_EXT_PAGE) And _
               (Left(ucFileName, Len(PREFIX_PAGE)) = PREFIX_PAGE) And _
               processPages Then
            '"Pge_[].adp" file
            If SADebug Then Debug.Print origFileName & " is a data access page file"
            objectType = acPage
            pageCount = pageCount + 1
            objectName = Mid(origFileName, Len(PREFIX_PAGE) + 1, _
                         Len(origFileName) - Len(PREFIX_PAGE) - Len(FILE_EXT_PAGE))

        Else
            'Unknown file type.  Ignore it.
        End If

        If importObjects And (objectType <> "") Then
            app.LoadFromText objectType, objectName, importFolder & origFileName
        End If

    End If

    objectType = ""
    objectName = ""
    origFileName = Dir
    ucFileName = UCase(origFileName)

Wend

'"Statistics for " & importFolder & ":" & vbCrLf & vbCrLf &
ImportDatabaseObjects = _
       TableCount & " tables" & vbCrLf & _
       queryCount & " queries" & vbCrLf & _
       formCount & " forms" & vbCrLf & _
       moduleCount & " modules" & vbCrLf & _
       macroCount & " macros" & vbCrLf & _
       reportCount & " reports" & vbCrLf & _
       pageCount & " data access pages"

Exit_ImportDatabaseObjects:
    Exit Function

Err_ImportDatabaseObjects:
    MsgBox Err.number & " - " & Err.Description
    Resume Exit_ImportDatabaseObjects

End Function

Public Function TableContainsOleFields(TableName As String) As Boolean
On Error GoTo ErrProc
Dim db As DAO.Database
Dim td As TableDef
Dim f As Field
TableContainsOleFields = False
Set db = CurrentDb
Set td = db.TableDefs(TableName)
For Each f In td.Fields
    If (f.Type = dbLongBinary) Or (f.Type = dbVarBinary) Then
        TableContainsOleFields = True
        Exit For
    End If
Next f
Exit Function
ErrProc:
DispErrMsgGSb Error$, "check if a table contains OLE fields"
Resume Next
End Function

Public Function ExportThisDataBase() As Boolean
On Error GoTo ErrProc

If Not IsNull(Forms("MattsVCSFrm")) Then
    exportLocGbl = Form_MattsVCSFrm.C_SourceDirNxt
Else
    exportLocGbl = DEFAULT_EXPORT_LOC
End If
If (Right(exportLocGbl, 1) <> "\") Then
    exportLocGbl = exportLocGbl & "\"
End If
Log "Exporting to " & exportLocGbl
CheckAndBuildFolderGFn (exportLocGbl)

ExportThisDataBase = False

'Set app = Application
'createStubDatabase CurrentDb, GetFSO.GetParentFolderName(exportLocGbl), False

MsgBox "EXPORTED:" & vbCrLf & ExportChangedItems(exportLocGbl)
CreateShortcut CreateBuildScript(GetFSO.GetParentFolderName(exportLocGbl)), _
                                 GetFSO.GetFileName(CurrentDb.Name)
ExportThisDataBase = True

Exit Function
ErrProc:
DispErrMsgGSb Error$, "export the database"
End Function

Public Function ImportThisDataBase() As Boolean

Dim importLoc As String
importLoc = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\src"
ImportThisDataBase = False

Set app = Application
MsgBox "IMPORTED:" & vbCrLf & ImportDatabaseObjects(importLoc, True)
ImportThisDataBase = True

End Function


Public Function testStuff()
Debug.Print "dbAttachExclusive: "
printHexTest dbAttachExclusive
Debug.Print "dbAttachSavePWD: "
printHexTest dbAttachSavePWD
Debug.Print "dbSystemObject: "
printHexTest dbSystemObject
Debug.Print "dbHiddenObject: "
printHexTest dbHiddenObject
Debug.Print "dbAttachedTable: "
printHexTest dbAttachedTable
Debug.Print "dbAttachedODBC: "
printHexTest dbAttachedODBC
End Function

Public Function OpenMattsVCSForm()
DoCmd.OpenForm "MattsVCSFrm"
'MsgBox "Code Project Name: " & Application.CodeProject.Name & vbCrLf &
'       "Code Project FullName: " & Application.CodeProject.FullName & vbCrLf & _
'       "Count of Modules: " & Application.CodeProject.AllModules.Count

End Function

Private Function CreateBuildScript(pathname As String) As String
On Error GoTo ErrProc
Dim scriptFilename As String
Dim scriptPathAndFilename As String
Dim scriptStr As String

If (Right(pathname, 1) <> "\") Then pathname = pathname & "\"
scriptFilename = "Build_Script.vbs"
scriptPathAndFilename = pathname & scriptFilename

Log "Creating Build Script: " & scriptPathAndFilename

Dim ScriptTblRS As DAO.Recordset
Set ScriptTblRS = CodeDb.OpenRecordset("SELECT BuildScript FROM __ScriptTbl__", dbOpenSnapshot)
    If Not ScriptTblRS.EOF Then
        ScriptTblRS.MoveFirst
        scriptStr = ScriptTblRS("BuildScript")
    End If
ScriptTblRS.Close

If Not IsNull(scriptStr) Then
    Dim FNum As Integer
    'Write out file
    FNum = FreeFile() 'Assign a free file no. to FNum
    Open scriptPathAndFilename For Output As FNum
    Print #FNum, scriptStr 'Write the contents of scriptStr to the script file.
    Close #FNum
End If

CreateBuildScript = scriptPathAndFilename
Exit Function

ErrProc:
CreateBuildScript = ""
End Function

Public Function Test_CreateBuildScript()
CreateBuildScript "C:\test"
End Function

Private Function CreateShortcut(scriptPathAndFilename As String, dbName As String) As String
Dim Link As Object
'Dim DesktopPath As String
Dim scriptPath As String
Dim scriptFilename As String
scriptPath = GetFSO.GetParentFolderName(scriptPathAndFilename)
scriptFilename = GetFSO.GetFileName(scriptPathAndFilename)

Log "Creating shortcut to build script " & scriptPath & "\MAKE.lnk"

'DesktopPath = GetShell.SpecialFolders("Desktop")
Set Link = GetShell.CreateShortcut(scriptPath & "\MAKE.lnk")
Link.Arguments = " """ & dbName & """"
Link.Description = "Shortcut to build " & dbName
Link.Hotkey = ""
'Link.IconLocation = ""
Link.TargetPath = """" & scriptPathAndFilename & """"
'"C:\Documents and Settings\Matt\My Documents\Projects\MattsVCS-Access\MattsVCS-Access-Addin\AccessVCAddIn\Build_Script.vbs" AccessVCAddin.mdb
Link.WindowStyle = 1 'Normal Window
Link.WorkingDirectory = scriptPath
Link.Save

CreateShortcut = Link.FullName

End Function

Public Function Test_CreateShortcut()
Debug.Print "Shortcut filename: " & CreateShortcut(CreateBuildScript("C:\test"), "Hello")
End Function

Public Function BuildDatabase()
On Error GoTo ErrProc
Dim ShellStr As String
If Not IsNull(Forms("MattsVCSFrm")) Then
    ShellStr = GetFSO.GetParentFolderName(Form_MattsVCSFrm.C_SourceDirNxt) & "\MAKE.lnk"
    If (Dir(ShellStr) <> "") Then
        GetShell.Run """" & ShellStr & """", vbHidden, False 'Don't wait for completion
        'Application.Quit acQuitSaveNone
    Else
        MsgBox "There is no MAKE link in the appropriate directory." & vbCrLf & _
               "Ensure you have exported this database.", , "No MAKE file!"
    End If
End If
ExitProc:
Exit Function
ErrProc:
DispErrMsgGSb Error$, "build a new copy of the database"
End Function