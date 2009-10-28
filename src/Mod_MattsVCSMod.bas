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

'PictureTypes
Const Embedded = 0
Const Linked = 1

Dim appAccess As Access.Application
Const SADebug As Boolean = True

Dim tableCount, queryCount, formCount, moduleCount As Integer
Dim macroCount, reportCount, pageCount, classCount As Integer
Dim processTables, processQueries, processForms, processModules As Boolean
Dim processMacros, processReports, processPages, processClasses As Boolean

Public Sub test()
Dim c As Container
For Each c In CurrentDb.Containers
    Debug.Print c.Name
Next c
End Sub

Public Sub test2()
Dim d As Document
Dim c As Container
Dim db As Database
Debug.Print "MODULES"
Set db = Application.CurrentDb
Set c = db.Containers("Modules")
For Each d In c.Documents
    Debug.Print d.Name
Next d
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
'            appAccess.SaveAsText docType, d.Name, _
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
    If d.Name <> "MattsVCSFrm" Then
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

Public Sub Test_exportForm()
Application.SaveAsText acForm, "MattsVCSFrm", _
        "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\src\Frm_MattsVCSFrm.frm"
End Sub

'Takes the exportFilename and exports all objects to files in the given exportLocation
Public Function SAExportDatabaseObjects(exportLocation As String, _
                                Optional performExport As Boolean = True) As String
On Error GoTo Err_SAExportDatabaseObjects

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

Set db = appAccess.CurrentDb

StartTimer

'exportLocation = "C:\OPT\TestDBs\Export\" 'Do not forget the closing back slash! ie: C:\Temp\
If Right(exportLocation, 1) <> "\" Then
    exportLocation = exportLocation & "\"
End If

If Not (db Is Nothing) Then
    If processTables Then
        For Each td In db.TableDefs 'Tables
            If Left(td.Name, 4) <> "MSys" Then
                If performExport Then
                    'for each
                    If tableDataInXML Or SAContainsOleFields(td) Then
                    'If a table contains OLE objects, use XML (.xml)
                    'Otherwise stick with an XML schema (.xsd) + text data (.txt)
                        appAccess.ExportXML objectType:=acExportTable, _
                                              DataSource:=td.Name, _
                                              DataTarget:=exportLocation & PREFIX_TABLE & td.Name & FILE_EXT_TABLE_COMBINED, _
                                              OtherFlags:=acEmbedSchema
                    Else
                        appAccess.ExportXML objectType:=acExportTable, _
                                              DataSource:=td.Name, _
                                              SchemaTarget:=exportLocation & PREFIX_TABLE & td.Name & FILE_EXT_TABLE_SCHEMA
                        appAccess.DoCmd.TransferText acExportDelim, , td.Name, exportLocation & PREFIX_TABLE & td.Name & FILE_EXT_TABLE_DATA, True
                    End If
                End If
                tableCount = tableCount + 1
                CheckTimer
            End If
        Next td
    End If
    SAExportDatabaseObjects = tableCount & " tables" & vbCrLf

    If processForms Then
        Set c = db.Containers("Forms")
        For Each d In c.Documents
            If performExport Then
                appAccess.SaveAsText acForm, d.Name, exportLocation & PREFIX_FORM & d.Name & FILE_EXT_FORM
            End If
            formCount = formCount + 1
            CheckTimer
        Next d
    End If
    SAExportDatabaseObjects = SAExportDatabaseObjects & formCount & " forms" & vbCrLf
    
    If processClasses Then
        Set c = db.Containers("Classes")
        For Each d In c.Documents
            If performExport Then
                appAccess.SaveAsText acReport, d.Name, exportLocation & PREFIX_CLASS & d.Name & FILE_EXT_CLASS
            End If
            classCount = classCount + 1
            CheckTimer
        Next d
    End If
    SAExportDatabaseObjects = SAExportDatabaseObjects & classCount & " classes" & vbCrLf
    
    'If processPages ...
    'pageCount & " data access pages"
    
    If processReports Then
        Set c = db.Containers("Reports")
        For Each d In c.Documents
            If performExport Then
                appAccess.SaveAsText acReport, d.Name, exportLocation & PREFIX_REPORT & d.Name & FILE_EXT_REPORT
            End If
            reportCount = reportCount + 1
            CheckTimer
        Next d
    End If
    SAExportDatabaseObjects = SAExportDatabaseObjects & reportCount & " reports" & vbCrLf
    
    If processMacros Then
        Set c = db.Containers("Scripts")
        For Each d In c.Documents
            If performExport Then
                appAccess.SaveAsText acMacro, d.Name, _
                                     exportLocation & PREFIX_MACRO & d.Name & FILE_EXT_MACRO
            End If
            macroCount = macroCount + 1
            CheckTimer
        Next d
    End If
    SAExportDatabaseObjects = SAExportDatabaseObjects & macroCount & " macros" & vbCrLf
    
    If processModules Then
        Set c = db.Containers("Modules")
        For Each d In c.Documents
            If performExport Then
                appAccess.SaveAsText acModule, d.Name, _
                                     exportLocation & PREFIX_MODULE & d.Name & FILE_EXT_MODULE
            End If
            moduleCount = moduleCount + 1
            CheckTimer
        Next d
    End If
    SAExportDatabaseObjects = SAExportDatabaseObjects & moduleCount & " modules" & vbCrLf
    
    If processQueries Then
        For i = 0 To db.QueryDefs.Count - 1
            'Skip the embedded queries
            If Left(db.QueryDefs(i).Name, 1) <> "~" Then
                If performExport Then
                    appAccess.SaveAsText acQuery, db.QueryDefs(i).Name, _
                                         exportLocation & PREFIX_QUERY & db.QueryDefs(i).Name & FILE_EXT_QUERY
                End If
                queryCount = queryCount + 1
                CheckTimer
            End If
        Next i
    End If
    SAExportDatabaseObjects = SAExportDatabaseObjects & queryCount & " queries" & vbCrLf
    
End If

Set db = Nothing
Set c = Nothing

If False Then MsgBox "All database objects have been exported as text and XML files to " & exportLocation & vbCrLf & _
       "Total time taken: " & GetTimeString(CheckTimer), _
       vbInformation

Exit_SAExportDatabaseObjects:
    Exit Function
    
Err_SAExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_SAExportDatabaseObjects

End Function

'Imports all valid text files in the importFolder to the currentDB of appAccess.
Public Function SAImportDatabaseObjects(importFolder As String, _
                                      Optional importObjects As Boolean = True) _
                                      As String
On Error GoTo Err_SAImportDatabaseObjects

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
                appAccess.ImportXML importFolder & origFileName, acStructureAndData
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
                appAccess.ImportXML importFolder & origFileName, acStructureOnly
                'Also import the corresponding data file
                objectName = Mid(origFileName, Len(PREFIX_TABLE) + 1, _
                             Len(origFileName) - Len(PREFIX_TABLE) - Len(FILE_EXT_TABLE_SCHEMA))
                dataFileName = importFolder & PREFIX_TABLE & objectName & FILE_EXT_TABLE_DATA
                If GetFSO.FileExists(dataFileName) Then
                    appAccess.DoCmd.TransferText acImportDelim, , objectName, dataFileName, True
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
            appAccess.LoadFromText objectType, objectName, importFolder & origFileName
        End If
        
    End If
        
    objectType = ""
    objectName = ""
    origFileName = Dir
    ucFileName = UCase(origFileName)
    
Wend

'"Statistics for " & importFolder & ":" & vbCrLf & vbCrLf &
SAImportDatabaseObjects = _
       tableCount & " tables" & vbCrLf & _
       queryCount & " queries" & vbCrLf & _
       formCount & " forms" & vbCrLf & _
       moduleCount & " modules" & vbCrLf & _
       macroCount & " macros" & vbCrLf & _
       reportCount & " reports" & vbCrLf & _
       pageCount & " data access pages"
    
Exit_SAImportDatabaseObjects:
    Exit Function
    
Err_SAImportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_SAImportDatabaseObjects
    
End Function

Public Function SAContainsOleFields(td As TableDef) As Boolean
Dim f As Field
SAContainsOleFields = False
For Each f In td.Fields
    If f.Type = dbLongBinary Then
        SAContainsOleFields = True
        Exit For
    End If
Next f
End Function

Public Function SAExportThisDataBase() As Boolean

Dim exportLoc As String
exportLoc = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\src"
SAExportThisDataBase = False

Set appAccess = Application
MsgBox "EXPORTED:" & vbCrLf & SAExportDatabaseObjects(exportLoc, True)
SAExportThisDataBase = True

End Function

Public Function SAImportThisDataBase() As Boolean

Dim importLoc As String
importLoc = "G:\repos\MattsVCS\MattsVCS-Access\MattsVCS-Access-Addin\src"
SAImportThisDataBase = False

Set appAccess = Application
MsgBox "IMPORTED:" & vbCrLf & SAImportDatabaseObjects(importLoc, True)
SAImportThisDataBase = True

End Function