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

Public Const TABLE_LIST_TABLENAME As String = "__TABLE_LIST__"
Public Const TABLE_LIST_FILENAME As String = TABLE_LIST_TABLENAME & ".xml"
Public exportLoc As String

'These are constants for tableDefs that are wrong.
'This is incorrectly defined as &H1 within Access
'Public Const dbHiddenObject As Long = &H2
'This is incorrectly defined as &H80000002 within Access
'Public Const dbSystemObject As Long = &H80000000
'This makes some kind of sense, in that every system object
'is hidden, but it means
' (tableDef.Attributes And dbSystemObject) = true
'when tableDef.Attributes = dbHiddenObject.
            
'PictureTypes
Const Embedded = 0
Const Linked = 1

Dim app As Access.Application
Const SADebug As Boolean = True

Dim tableCount, queryCount, formCount, moduleCount As Integer
Dim macroCount, reportCount, pageCount, classCount As Integer
Dim processTables, processQueries, processForms, processModules As Boolean
Dim processMacros, processReports, processPages, processClasses As Boolean

Private Sub ExportForm(formName As String)
Set app = Access.Application
app.SaveAsText acForm, formName, _
               exportLoc & formName & ".frm"
End Sub

Private Sub Test_ExportForm1()
'ExportForm "MattsVCSFrm"
ExportForm "TableSubFrm"
End Sub

Private Sub ImportForm(formName As String)
Set app = Access.Application
app.LoadFromText acForm, formName, _
                 exportLoc & formName & ".frm"
End Sub

Private Sub Test_ImportForm()
'ImportForm "MattsVCSFrm"
ImportForm "TableSubFrm"
End Sub

Private Sub ExportTableSchemaAsXsd(tableNameStr As String)
Set app = Access.Application
app.ExportXML objectType:=acExportTable, _
              DataSource:=tableNameStr, _
              SchemaTarget:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_SCHEMA
End Sub

Private Sub ExportTableSchemaAndDataAsXml(tableNameStr As String)
Set app = Access.Application
app.ExportXML objectType:=acExportTable, _
              DataSource:=tableNameStr, _
              DataTarget:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_COMBINED, _
              OtherFlags:=acEmbedSchema
End Sub

Private Sub ExportTableDataOnlyAsXml(tableNameStr As String)
Set app = Access.Application
app.ExportXML objectType:=acExportTable, _
              DataSource:=tableNameStr, _
              DataTarget:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_COMBINED
End Sub

Private Sub ExportTableDataAsTxt(tableNameStr As String)
Set app = Access.Application
app.DoCmd.TransferText TransferType:=acExportDelim, _
                       TableName:=tableNameStr, _
                       FileName:=exportLoc & PREFIX_TABLE & tableNameStr & FILE_EXT_TABLE_DATA, _
                       HasFieldNames:=True
End Sub

Private Sub ExportListedTables()
'Currently takes 16 seconds
Set app = Access.Application
Debug.Print "***** Tables *****"
'exportLoc = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\test\src\"
exportLoc = GetDBFolderNameGFn(CurrentDb) & "\" & GetFSO.GetBaseName(CurrentDb.Name) & "\src\"
CheckAndBuildFolderGFn (exportLoc)
Debug.Print TABLE_LIST_TABLENAME
app.ExportXML objectType:=acExportTable, _
              DataSource:=TABLE_LIST_TABLENAME, _
              DataTarget:=exportLoc & TABLE_LIST_FILENAME, _
              OtherFlags:=acEmbedSchema
Dim TableList As dao.Recordset
Set TableList = CurrentDb.OpenRecordset("SELECT * FROM " & TABLE_LIST_TABLENAME, dbOpenSnapshot)
    If Not TableList.EOF Then
        TableList.MoveFirst
        While Not TableList.EOF
            If TableList("Tbl_ContainsBinary") Then
                If TableList("Tbl_ExportSchema") Then
                    If TableList("Tbl_ExportData") Then
                        Debug.Print TableList("Tbl_Name") & " (Combined XML)"
                        ExportTableSchemaAndDataAsXml TableList("Tbl_Name")
                    Else
                        Debug.Print TableList("Tbl_Name") & " (XSD Schema)"
                        ExportTableSchemaAsXsd TableList("Tbl_Name")
                    End If
                Else
                    If TableList("Tbl_ExportData") Then
                        Debug.Print TableList("Tbl_Name") & " (XML Data Only)"
                        ExportTableDataOnlyAsXml TableList("Tbl_Name")
                    End If
                End If
            Else
                If TableList("Tbl_ExportSchema") Then
                    Debug.Print TableList("Tbl_Name") & " (XSD Schema)"
                    ExportTableSchemaAsXsd TableList("Tbl_Name")
                    If TableList("Tbl_ExportData") Then
                        Debug.Print TableList("Tbl_Name") & " (TXT Data)"
                        ExportTableDataAsTxt TableList("Tbl_Name")
                    End If
                Else
                    If TableList("Tbl_ExportData") Then
                        Debug.Print TableList("Tbl_Name") & " (TXT Data)"
                        ExportTableDataAsTxt TableList("Tbl_Name")
                    End If
                End If
            End If
            TableList.MoveNext
        Wend
    End If
TableList.Close

Dim exportIniFile As String
exportIniFile = Dir(exportLoc & "export.ini")
If exportIniFile <> "" Then Kill exportLoc & exportIniFile

End Sub

Public Sub test()
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

Public Sub test_TableDefs()
Dim td As TableDef
Dim db As dao.Database
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

Public Sub Test_ExportForm()
Application.SaveAsText acForm, "Copy of MattsVCSFrm", _
        "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\src\Frm_MattsVCSFrm.frm"
End Sub

'Takes the exportFilename and exports all objects to files in the given exportLocation
Public Function ExportDatabaseObjects(exportLocation As String, _
                                Optional performExport As Boolean = True) As String
On Error GoTo Err_ExportDatabaseObjects

Dim db As Database
Dim td As TableDef
Dim d As Document
Dim c As Container
Dim i As Integer
'Dim exportLocation As String
Dim tableDataInXML As Boolean
tableDataInXML = False

tableCount = 0
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

Set db = app.CurrentDb

StartTimer

'exportLocation = "C:\OPT\TestDBs\Export\" 'Do not forget the closing back slash! ie: C:\Temp\
If Right(exportLocation, 1) <> "\" Then
    exportLocation = exportLocation & "\"
End If

If Not (db Is Nothing) Then
    If processTables Then
        'tableCount = ListTables
        ExportListedTables
    End If
    ExportDatabaseObjects = tableCount & " tables" & vbCrLf

    If processForms Then
        Debug.Print "***** Forms *****"
        Set c = db.Containers("Forms")
        For Each d In c.Documents
            If performExport Then
                Debug.Print "Exporting Form: " & d.Name
                app.SaveAsText acForm, d.Name, exportLocation & PREFIX_FORM & d.Name & FILE_EXT_FORM
            End If
            formCount = formCount + 1
            CheckTimer
        Next d
    End If
    ExportDatabaseObjects = ExportDatabaseObjects & formCount & " forms" & vbCrLf
    
    If processClasses Then
        Debug.Print "***** Classes *****"
        Set c = db.Containers("Classes")
        For Each d In c.Documents
            If performExport Then
                Debug.Print "Exporting Class: " & d.Name
                app.SaveAsText acReport, d.Name, exportLocation & PREFIX_CLASS & d.Name & FILE_EXT_CLASS
            End If
            classCount = classCount + 1
            CheckTimer
        Next d
    End If
    ExportDatabaseObjects = ExportDatabaseObjects & classCount & " classes" & vbCrLf
    
    'If processPages ...
    'pageCount & " data access pages"
    
    If processReports Then
        Debug.Print "***** Reports *****"
        Set c = db.Containers("Reports")
        For Each d In c.Documents
            If performExport Then
                Debug.Print "Exporting Report: " & d.Name
                app.SaveAsText acReport, d.Name, exportLocation & PREFIX_REPORT & d.Name & FILE_EXT_REPORT
            End If
            reportCount = reportCount + 1
            CheckTimer
        Next d
    End If
    ExportDatabaseObjects = ExportDatabaseObjects & reportCount & " reports" & vbCrLf
    
    If processMacros Then
        Debug.Print "***** Macros *****"
        Set c = db.Containers("Scripts")
        For Each d In c.Documents
            If performExport Then
                Debug.Print "Exporting Macro: " & d.Name
                app.SaveAsText acMacro, d.Name, _
                                     exportLocation & PREFIX_MACRO & d.Name & FILE_EXT_MACRO
            End If
            macroCount = macroCount + 1
            CheckTimer
        Next d
    End If
    ExportDatabaseObjects = ExportDatabaseObjects & macroCount & " macros" & vbCrLf
    
    If processModules Then
        Debug.Print "***** Modules *****"
        Set c = db.Containers("Modules")
        For Each d In c.Documents
            If performExport Then
                Debug.Print "Exporting Module: " & d.Name
                app.SaveAsText acModule, d.Name, _
                                     exportLocation & PREFIX_MODULE & d.Name & FILE_EXT_MODULE
            End If
            moduleCount = moduleCount + 1
            CheckTimer
        Next d
    End If
    ExportDatabaseObjects = ExportDatabaseObjects & moduleCount & " modules" & vbCrLf
    
    If processQueries Then
        Debug.Print "***** Queries *****"
        For i = 0 To db.QueryDefs.Count - 1
            'Skip the embedded queries
            If Left(db.QueryDefs(i).Name, 1) <> "~" Then
                If performExport Then
                    Debug.Print "Exporting Query: " & db.QueryDefs(i).Name
                    app.SaveAsText acQuery, db.QueryDefs(i).Name, _
                                         exportLocation & PREFIX_QUERY & db.QueryDefs(i).Name & FILE_EXT_QUERY
                End If
                queryCount = queryCount + 1
                CheckTimer
            End If
        Next i
    End If
    ExportDatabaseObjects = ExportDatabaseObjects & queryCount & " queries" & vbCrLf
    
End If

Set db = Nothing
Set c = Nothing

If False Then MsgBox "All database objects have been exported as text and XML files to " & exportLocation & vbCrLf & _
       "Total time taken: " & GetTimeString(CheckTimer), _
       vbInformation

Exit_ExportDatabaseObjects:
    Exit Function
    
Err_ExportDatabaseObjects:
    MsgBox Err.number & " - " & Err.Description & vbCrLf & _
    Error$
    Resume Next

End Function

'Imports all valid text files in the importFolder to the currentDB of app.
Public Function ImportDatabaseObjects(importFolder As String, _
                                      Optional importObjects As Boolean = True) _
                                      As String
On Error GoTo Err_ImportDatabaseObjects

'importFolder = "C:\OPT\TestDBs\Export\" 'Do not forget the closing back slash! ie: C:\Temp\

If (Right(importFolder, 1) <> "\") Then
    importFolder = importFolder & "\"
End If

Dim origFileName As String
Dim ucFileName As String
Dim objectType As String
Dim objectName As String
Dim dataFileName As String

tableCount = 0
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
            tableCount = tableCount + 1
            
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
            tableCount = tableCount + 1
            
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
       tableCount & " tables" & vbCrLf & _
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
Dim db As dao.Database
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

Dim exportLoc As String
If Not IsNull(Forms("MattsVCSFrm")) Then
    exportLoc = Form_MattsVCSFrm.C_SourceDirNxt
Else
    exportLoc = "C:\Documents and Settings\Matt\My Documents\Projects\test"
End If
ExportThisDataBase = False

Set app = Application
MsgBox "EXPORTED:" & vbCrLf & ExportDatabaseObjects(exportLoc, True)
ExportThisDataBase = True

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

Public Function Test_OpenForm()
DoCmd.OpenForm "MattsVCSFrm"
MsgBox "Code Project Name: " & Application.CodeProject.Name & vbCrLf & _
       "Code Project FullName: " & Application.CodeProject.FullName & vbCrLf & _
       "Count of Modules: " & Application.CodeProject.AllModules.Count

End Function