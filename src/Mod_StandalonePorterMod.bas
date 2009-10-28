''Author:    Matt Fisher
''Created:   13 May 2008
'
''TODO:
''   Make tables from external database files distinguishable once exported
''   Recreate external database files on import
''   Put linked tables in external files and relink to front db
''   Investigate additional DB info - startup, icon, properties, etc.
'
''   Figure out WTF is going on WRT no DB being open when we come into Export the first time
''   Figure out how to overwrite objects if they already exist
''   Stop database program running when opened
'
'Option Compare Database
'Option Explicit
'
'Const FILE_EXT_MODULE As String = ".bas"
'Const FILE_EXT_FORM As String = ".frm"
'Const FILE_EXT_CLASS As String = ".cls"
'Const FILE_EXT_REPORT As String = ".rpt"
'Const FILE_EXT_MACRO As String = ".mac"
'Const FILE_EXT_QUERY As String = ".qry"
'
'Const EXPORT_FILE_EXT As String = ".txt"
'Const TABLE_SCHEMA_FILE_EXT As String = ".xsd"
'Const TABLE_DATA_FILE_EXT As String = ".txt"
'Const TABLE_COMBINED_FILE_EXT As String = ".xml"
'
'Const QUERY_PREFIX As String = "Qry_"
'Const FORM_PREFIX As String = "Frm_"
'Const MODULE_PREFIX As String = "Mod_"
'Const MACRO_PREFIX As String = "Mac_"
'Const TABLE_PREFIX As String = "Tbl_"
'Const REPORT_PREFIX As String = "Rpt_"
'Const PAGE_PREFIX As String = "Pge_"
'
'Dim appAccess As Access.Application
'Const SADebug As Boolean = True
'
'Dim tableCount As Integer
'Dim queryCount As Integer
'Dim formCount As Integer
'Dim moduleCount As Integer
'Dim macroCount As Integer
'Dim reportCount As Integer
'Dim pageCount As Integer
'Dim processTables As Boolean
'Dim processQueries As Boolean
'Dim processForms As Boolean
'Dim processModules As Boolean
'Dim processMacros As Boolean
'Dim processReports As Boolean
'Dim processPages As Boolean
'
'Dim fso As Object
'
'
''Takes the exportFilename and exports all objects to files in the given exportLocation
'Public Function SAExportDatabaseObjects(exportLocation As String, _
'                                Optional performExport As Boolean = True) As String
'On Error GoTo Err_SAExportDatabaseObjects
'
'Dim db As Database
'Dim td As TableDef
'Dim d As Document
'Dim c As Container
'Dim i As Integer
''Dim exportLocation As String
'Dim tableDataInXML As Boolean
'tableDataInXML = False
'
'tableCount = 0
'queryCount = 0
'formCount = 0
'moduleCount = 0
'macroCount = 0
'reportCount = 0
'pageCount = 0
'
'processTables = True
'processQueries = True
'processForms = True
'processModules = True
'processMacros = True
'processReports = True
'processPages = True
'
'Set db = appAccess.CurrentDb
'
'StartTimer
'
''exportLocation = "C:\OPT\TestDBs\Export\" 'Do not forget the closing back slash! ie: C:\Temp\
'If Right(exportLocation, 1) <> "\" Then
'    exportLocation = exportLocation & "\"
'End If
'
'If Not (db Is Nothing) Then
'    If processTables Then
'        For Each td In db.TableDefs 'Tables
'            If Left(td.Name, 4) <> "MSys" Then
'                If performExport Then
'                    'for each
'                    If tableDataInXML Or SAContainsOleFields(td) Then
'                    'If a table contains OLE objects, use XML (.xml)
'                    'Otherwise stick with an XML schema (.xsd) + text data (.txt)
'                        appAccess.ExportXML objectType:=acExportTable, _
'                                              DataSource:=td.Name, _
'                                              DataTarget:=exportLocation & TABLE_PREFIX & td.Name & TABLE_COMBINED_FILE_EXT, _
'                                              OtherFlags:=acEmbedSchema
'                    Else
'                        appAccess.ExportXML objectType:=acExportTable, _
'                                              DataSource:=td.Name, _
'                                              SchemaTarget:=exportLocation & TABLE_PREFIX & td.Name & TABLE_SCHEMA_FILE_EXT
'                        appAccess.DoCmd.TransferText acExportDelim, , td.Name, exportLocation & TABLE_PREFIX & td.Name & TABLE_DATA_FILE_EXT, True
'                    End If
'                End If
'                tableCount = tableCount + 1
'                CheckTimer
'            End If
'        Next td
'    End If
'
'    If processForms Then
'        Set c = db.Containers("Forms")
'        For Each d In c.Documents
'            If performExport Then
'                appAccess.SaveAsText acForm, d.Name, _
'                                     exportLocation & FORM_PREFIX & d.Name & FILE_EXT_FORM
'            End If
'            formCount = formCount + 1
'            CheckTimer
'        Next d
'    End If
'
'    If processReports Then
'        Set c = db.Containers("Reports")
'        For Each d In c.Documents
'            If performExport Then
'                appAccess.SaveAsText acReport, d.Name, _
'                                     exportLocation & REPORT_PREFIX & d.Name & FILE_EXT_REPORT
'            End If
'            reportCount = reportCount + 1
'            CheckTimer
'        Next d
'    End If
'
'    If processMacros Then
'        Set c = db.Containers("Scripts")
'        For Each d In c.Documents
'            If performExport Then
'                appAccess.SaveAsText acMacro, d.Name, _
'                                     exportLocation & MACRO_PREFIX & d.Name & FILE_EXT_MACRO
'            End If
'            macroCount = macroCount + 1
'            CheckTimer
'        Next d
'    End If
'
'    If processModules Then
'        Set c = db.Containers("Modules")
'        For Each d In c.Documents
'            If performExport Then
'                appAccess.SaveAsText acModule, d.Name, _
'                                     exportLocation & MODULE_PREFIX & d.Name & FILE_EXT_MODULE
'            End If
'            moduleCount = moduleCount + 1
'            CheckTimer
'        Next d
'    End If
'
'    If processQueries Then
'        For i = 0 To db.QueryDefs.Count - 1
'            'Skip the embedded queries
'            If Left(db.QueryDefs(i).Name, 1) <> "~" Then
'                If performExport Then
'                    appAccess.SaveAsText acQuery, db.QueryDefs(i).Name, _
'                                         exportLocation & QUERY_PREFIX & db.QueryDefs(i).Name & FILE_EXT_QUERY
'                End If
'                queryCount = queryCount + 1
'                CheckTimer
'            End If
'        Next i
'    End If
'End If
'
'Set db = Nothing
'Set c = Nothing
'
'SAExportDatabaseObjects = _
'       tableCount & " tables" & vbCrLf & _
'       queryCount & " queries" & vbCrLf & _
'       formCount & " forms" & vbCrLf & _
'       moduleCount & " modules" & vbCrLf & _
'       macroCount & " macros" & vbCrLf & _
'       reportCount & " reports" & vbCrLf & _
'       pageCount & " data access pages"
'
'If False Then MsgBox "All database objects have been exported as text and XML files to " & exportLocation & vbCrLf & _
'       "Total time taken: " & GetTimeString(CheckTimer), _
'       vbInformation
'
'Exit_SAExportDatabaseObjects:
'    Exit Function
'
'Err_SAExportDatabaseObjects:
'    MsgBox Err.Number & " - " & Err.Description
'    Resume Exit_SAExportDatabaseObjects
'
'End Function
'
''Imports all valid text files in the importFolder to the currentDB of appAccess.
'Public Function SAImportDatabaseObjects(importFolder As String, _
'                                      Optional importObjects As Boolean = True) _
'                                      As String
'On Error GoTo Err_SAImportDatabaseObjects
'
''importFolder = "C:\OPT\TestDBs\Export\" 'Do not forget the closing back slash! ie: C:\Temp\
'
'If (Right(importFolder, 1) <> "\") Then
'    importFolder = importFolder & "\"
'End If
'
'Dim origFileName As String
'Dim ucFileName As String
'Dim objectType As String
'Dim objectName As String
'Dim dataFileName As String
'
'tableCount = 0
'queryCount = 0
'formCount = 0
'moduleCount = 0
'macroCount = 0
'reportCount = 0
'pageCount = 0
'
'processTables = True
'processQueries = True
'processForms = True
'processModules = True
'processMacros = True
'processReports = True
'processPages = True
'
'origFileName = Dir(importFolder, vbNormal)
'ucFileName = UCase(origFileName)
'
'While ucFileName <> ""
'    If Right(ucFileName, Len(TABLE_COMBINED_FILE_EXT)) = TABLE_COMBINED_FILE_EXT Then
'        '.xml file
'        If SADebug Then Debug.Print origFileName & " is a " & TABLE_COMBINED_FILE_EXT & " file"
'        If ((Left(ucFileName, Len(TABLE_PREFIX)) = TABLE_PREFIX) And processTables) Then
'            '"Tbl_" file
'            If SADebug Then Debug.Print origFileName & " is a combined table file"
'            tableCount = tableCount + 1
'
'            If importObjects Then
'                appAccess.ImportXML importFolder & origFileName, acStructureAndData
'            End If
'
'        End If
'    ElseIf Right(ucFileName, Len(TABLE_SCHEMA_FILE_EXT)) = TABLE_SCHEMA_FILE_EXT Then
'        '.xsd file
'        If SADebug Then Debug.Print origFileName & " is a " & TABLE_SCHEMA_FILE_EXT & " file"
'        If ((Left(ucFileName, Len(TABLE_PREFIX)) = TABLE_PREFIX) And processTables) Then
'            '"Tbl_" file
'            If SADebug Then Debug.Print origFileName & " is a table schema file"
'            tableCount = tableCount + 1
'
'            If importObjects Then
'                appAccess.ImportXML importFolder & origFileName, acStructureOnly
'                'Also import the corresponding data file
'                objectName = Mid(origFileName, Len(TABLE_PREFIX) + 1, _
'                             Len(origFileName) - Len(TABLE_PREFIX) - Len(TABLE_SCHEMA_FILE_EXT))
'                dataFileName = importFolder & TABLE_PREFIX & objectName & TABLE_DATA_FILE_EXT
'                If GetFSO.FileExists(dataFileName) Then
'                    appAccess.DoCmd.TransferText acImportDelim, , objectName, dataFileName, True
'                End If
'            End If
'
'        End If
'    ElseIf (Right(ucFileName, Len(EXPORT_FILE_EXT)) = EXPORT_FILE_EXT) Then
'        '.txt file
'        If SADebug Then Debug.Print origFileName & " is a " & EXPORT_FILE_EXT & " file"
'
'        If (Left(ucFileName, Len(QUERY_PREFIX)) = QUERY_PREFIX) And processQueries Then
'            '"Qry_" file
'            If SADebug Then Debug.Print origFileName & " is a query file"
'            objectName = Mid(origFileName, Len(QUERY_PREFIX) + 1, _
'                            Len(origFileName) - Len(QUERY_PREFIX) - Len(EXPORT_FILE_EXT))
'            'Skip embedded queries
'            If Left(objectName, 1) <> "~" Then
'                objectType = acQuery
'                queryCount = queryCount + 1
'            End If
'        ElseIf (Left(ucFileName, Len(MODULE_PREFIX)) = MODULE_PREFIX) And processModules Then
'            '"Mod_" file
'            If SADebug Then Debug.Print origFileName & " is a module file"
'            objectName = Mid(origFileName, Len(MODULE_PREFIX) + 1, _
'                            Len(origFileName) - Len(MODULE_PREFIX) - Len(EXPORT_FILE_EXT))
'            ' Don't overwrite yourself
'            If (objectName <> "StandalonePorterMod") Then
'                objectType = acModule
'                moduleCount = moduleCount + 1
'            End If
'        ElseIf (Left(ucFileName, Len(FORM_PREFIX)) = FORM_PREFIX) And processForms Then
'            '"Frm_" file
'            If SADebug Then Debug.Print origFileName & " is a form file"
'            objectType = acForm
'            formCount = formCount + 1
'            objectName = Mid(origFileName, Len(FORM_PREFIX) + 1, _
'                            Len(origFileName) - Len(FORM_PREFIX) - Len(EXPORT_FILE_EXT))
'        ElseIf (Left(ucFileName, Len(MACRO_PREFIX)) = MACRO_PREFIX) And processMacros Then
'            '"Mcr_" file
'            If SADebug Then Debug.Print origFileName & " is a macro file"
'            objectType = acMacro
'            macroCount = macroCount + 1
'            objectName = Mid(origFileName, Len(MACRO_PREFIX) + 1, _
'                            Len(origFileName) - Len(MACRO_PREFIX) - Len(EXPORT_FILE_EXT))
'        ElseIf (Left(ucFileName, Len(REPORT_PREFIX)) = REPORT_PREFIX) And processReports Then
'            '"Rpt_" file
'            If SADebug Then Debug.Print origFileName & " is a report file"
'            objectType = acReport
'            reportCount = reportCount + 1
'            objectName = Mid(origFileName, Len(REPORT_PREFIX) + 1, _
'                            Len(origFileName) - Len(REPORT_PREFIX) - Len(EXPORT_FILE_EXT))
'        ElseIf (Left(ucFileName, Len(PAGE_PREFIX)) = PAGE_PREFIX) And processPages Then
'            '"Pge_" file
'            If SADebug Then Debug.Print origFileName & " is a data access page file"
'            objectType = acPage
'            pageCount = pageCount + 1
'            objectName = Mid(origFileName, Len(PAGE_PREFIX) + 1, _
'                            Len(origFileName) - Len(PAGE_PREFIX) - Len(EXPORT_FILE_EXT))
'        Else
'            'Unknown file type.  Ignore it.
'        End If
'
'        If importObjects And (objectType <> "") Then
'            appAccess.LoadFromText objectType, objectName, importFolder & origFileName
'        End If
'
'    End If
'
'    objectType = ""
'    objectName = ""
'    origFileName = Dir
'    ucFileName = UCase(origFileName)
'
'Wend
'
''"Statistics for " & importFolder & ":" & vbCrLf & vbCrLf &
'SAImportDatabaseObjects = _
'       tableCount & " tables" & vbCrLf & _
'       queryCount & " queries" & vbCrLf & _
'       formCount & " forms" & vbCrLf & _
'       moduleCount & " modules" & vbCrLf & _
'       macroCount & " macros" & vbCrLf & _
'       reportCount & " reports" & vbCrLf & _
'       pageCount & " data access pages"
'
'Exit_SAImportDatabaseObjects:
'    Exit Function
'
'Err_SAImportDatabaseObjects:
'    MsgBox Err.Number & " - " & Err.Description
'    Resume Exit_SAImportDatabaseObjects
'
'End Function
'
'Public Function SAContainsOleFields(td As TableDef) As Boolean
'Dim f As Field
'SAContainsOleFields = False
'For Each f In td.Fields
'    If f.Type = dbLongBinary Then
'        SAContainsOleFields = True
'        Exit For
'    End If
'Next f
'End Function
'
'Public Function getDBFolderName() As String
'getDBFolderName = getFSO.GetParentFolderName(CurrentDb.Name)
'End Function
'
'Public Function SAExportThisDataBase() As Boolean
'
'Dim exportLoc As String
'exportLoc = getDBFolderName & "\src"
'SAExportThisDataBase = False
'
'Set appAccess = Application
'MsgBox "EXPORTED:" & vbCrLf & SAExportDatabaseObjects(exportLoc, True)
'SAExportThisDataBase = True
'
'End Function
'
'Public Function SAImportThisDataBase() As Boolean
'
'Dim importLoc As String
'importLoc = getDBFolderName & "\src"
'SAImportThisDataBase = False
'
'Set appAccess = Application
'MsgBox "IMPORTED:" & vbCrLf & SAImportDatabaseObjects(importLoc, True)
'SAImportThisDataBase = True
'
'End Function
'
'Public Function OpenMattsVCSForm() As Boolean
'Form_MattsVCSFrm.Visible = True
'End Function