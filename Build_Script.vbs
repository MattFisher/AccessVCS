' Usage:
'  WScript Build_Script.vbs <file> <path>

' Converts all modules, classes, forms and macros in a directory created by AccessVCS,
' then imports them into a Stub database, and composes then into an Access Database file (.mdb).
' Requires Microsoft Access.

Option Explicit

const acQuery = 1
const acForm = 2
const acReport = 3
const acMacro = 4
const acModule = 5

const filenamePrefixLength = 4

const acStructureOnly = 0
const acStructureAndData = 1
const acImportDelim = 0
const dbOpenSnapshot = 4

Const tableListName = "__TABLE_LIST__"

Const acCmdCompileAndSaveAllModules = &H7E
Const acCmdCompactDatabase = 4

' BEGIN CODE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

dim sMDBFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "Please enter the name of the database!", vbExclamation, "Error"
    Wscript.Quit()
End if
sMDBFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sPath
If (WScript.Arguments.Count = 1) then
    sPath = ""
else
    sPath = WScript.Arguments(1)
End If


importModulesTxt sMDBFilename, sPath

WScript.Echo "Build Succeeded!" & vbcrlf & sMDBFilename 

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If



Function importModulesTxt(sMDBFilename, sImportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    ' Build file and pathnames
    dim myType, myName, myPath, sStubMDBFilename
    myType = fso.GetExtensionName(sMDBFilename)
    myName = fso.GetBaseName(sMDBFilename)
    myPath = fso.GetParentFolderName(sMDBFilename)

    ' if no path was given as argument, use a relative directory
    If (sImportpath = "") then
        sImportpath = myPath & "\src\"
    End If
    sStubMDBFilename = myPath & "\" & myName & "_stub." & myType

    ' check for existing file and ask to overwrite with the stub
    if (fso.FileExists(sMDBFilename)) Then
        dim choice
        choice = MsgBox("The file: " & vbcrlf & sMDBFilename & vbcrlf & "already exists. Overwrite?", _
                        VbYesNo + VBQuestion, "File Already Exists!")
        if (choice <> VbYes) Then
            WScript.Quit
        end if

        fso.CopyFile sMDBFilename, sMDBFilename & ".bak"
    end if
    
    'WScript.Echo "Copying " & sStubMDBFilename & " to " & sMDBFilename

    fso.CopyFile sStubMDBFilename, sMDBFilename

    ' launch MSAccess
    'WScript.Echo "Starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    'WScript.Echo "opening " & sMDBFilename & " ..."
    'oApplication.OpenAccessProject sMDBFilename
    oApplication.OpenCurrentDatabase sMDBFilename
    oApplication.Visible = false

    Dim folder
    Set folder = fso.GetFolder(sImportpath)

    ' load each file from the import path into the stub
    Dim myFile, objectname, objecttype
    for each myFile in folder.Files
        objecttype = fso.GetExtensionName(myFile.Name)
        objectname = getObjectName(fso.GetBaseName(myFile.Name))
        'WScript.Echo "  " & objectname & " (" & objecttype & ")"

        'Table files with structure and data
        if (objecttype = "xml") then
            oApplication.ImportXML myFile.Path, acStructureAndData
        'Table files with structure only
        elseif (objecttype = "xsd") then
            oApplication.ImportXML myFile.Path, acStructureOnly
            'Also import the corresponding data file
            dim dataFileName
            dataFileName = folder.Path & "\" & objectName & ".txt"
            If fso.FileExists(dataFileName) Then
                oApplication.DoCmd.TransferText acImportDelim, , objectName, dataFileName, True
            End If
        'Forms
        elseif (objecttype = "frm") then
            oApplication.LoadFromText acForm, objectname, myFile.Path
        'Modules
        elseif (objecttype = "bas") then
            oApplication.LoadFromText acModule, objectname, myFile.Path
        'Macros
        elseif (objecttype = "mac") then
            oApplication.LoadFromText acMacro, objectname, myFile.Path
        'Reports
        elseif (objecttype = "rpt") then
            oApplication.LoadFromText acReport, objectname, myFile.Path
        'Queries
        elseif (objecttype = "qry") then
            oApplication.LoadFromText acQuery, objectname, myFile.Path
        end if

    next
    
    'WScript.Echo "Linked " & RelinkTables(oApplication.currentDb) & " tables."
    
    oApplication.RunCommand acCmdCompileAndSaveAllModules
    oApplication.RunCommand acCmdCompactDatabase
    oApplication.Quit
End Function

Private Function getObjectName(sFileBaseName)
    if (sFileBaseName <> tableListName) then
        getObjectName = mid(sFileBaseName, filenamePrefixLength + 1)
    else
        getObjectName = tableListName
    end if
End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function

Private Function RelinkTables(db)
'Iterate through the newly-created tablelist to relink all linked tables.

'Tbl_ID, Tbl_Name, Tbl_SourceTableName, Tbl_Connect, Tbl_ExportSchema, Tbl_ExportData, _
'Tbl_DispType, Tbl_Attributes, Tbl_System, Tbl_Hidden, Tbl_AttachedTable, Tbl_AttachedODBC, _
'Tbl_AttachSavePWD, Tbl_AttachExclusive, Tbl_ContainsBinary

'WScript.Echo db.name
Dim newTable
Dim TableCount
Dim TableList
Set TableList = db.OpenRecordset("SELECT * FROM " & tableListName, dbOpenSnapshot)
    If Not TableList.EOF Then
        TableList.MoveFirst
        While Not TableList.EOF
            If ( TableList("Tbl_AttachedTable") or TableList("Tbl_AttachedODBC") ) Then
                'Create new TableDef and connect it
                'WScript.Echo "Creating new TableDef" & vbcrlf & _
                '             "Tbl_Name: " & TableList("Tbl_Name") & vbcrlf &  _
                '             "Tbl_Attributes: " & TableList("Tbl_Attributes") & vbcrlf & _
                '             "Tbl_SourceTableName: " & TableList("Tbl_SourceTableName") & vbcrlf & _
                '             "Tbl_Connect: " & TableList("Tbl_Connect")
                set newTable = db.CreateTableDef(TableList("Tbl_Name"), 0, _
                                             TableList("Tbl_SourceTableName"), _
                                             TableList("Tbl_Connect"))
                                             '(Attributes: TableList("Tbl_Attributes"), _
                'WScript.Echo "New Table Created: " & newTable.name & vbcrlf & _
                '""
                'WScript.Echo "Current Table Count: " & db.TableDefs.Count
                'WScript.Echo "About to refresh link"
                'newTable.RefreshLink
                'WScript.Echo "Just refreshed link"
                'newTable.CreateField "TestField"
            'Somewhere around here we need to check if this fails and provide a method to locate the
            'missing database (ie FilePicker)
                db.TableDefs.Append newTable
                'WScript.Echo "New Table Count: " & db.TableDefs.Count
                TableCount = TableCount + 1
            End If
            TableList.MoveNext
        Wend
    End If
TableList.Close

RelinkTables = TableCount

End Function
