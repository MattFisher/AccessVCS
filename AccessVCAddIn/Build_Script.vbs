' Usage:
'  WScript Build_Script.vbs <file> <path>

' Composes all modules, classes, forms and macros in a directory created
' by Matt's Version Control System into an Access Database file (.mdb).
' This includes relinking external tables, restoring database properties,
' and re-establishing relationships between tables.
' Requires Microsoft Access.

Option Explicit

Const acTable = 0
const acQuery = 1
const acForm = 2
const acReport = 3
const acMacro = 4
const acModule = 5

Const msoFileDialogFilePicker = 3
Const msoFileDialogViewDetails = 2

const filenamePrefixLength = 4

const acStructureOnly = 0
const acStructureAndData = 1
const acImportDelim = 0
const dbOpenSnapshot = 4

Const dbBoolean = 1
Const dbByte = 2
Const dbInteger = 3
Const dbLong = 4
Const dbCurrency = 5
Const dbSingle = 6
Const dbDouble = 7
Const dbDate = 8
Const dbBinary = 9
Const dbText = 10
Const dbLongBinary = 11
Const dbMemo = 12
Const dbGUID = 15

Const tableListTableName = "__TABLES__"
Const propertiesTableName = "__PROPERTIES__"
Const relationsTableName = "__RELATIONS__"
Const referencesTableName = "__REFERENCES__"

Const acCmdCompileAndSaveAllModules = &H7E
Const acCmdCompactDatabase = 4

'BEGIN CODE
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim newDb
    
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
'END CODE

Function restoreStubIfPresent(sDbPath, sDbName, sDbType)
    Dim sStubMDBFilename
    Dim sNewMDBFilename
    sStubMDBFilename = sDbPath & "\" & sDbName & "_stub." & sDbType
    sNewMDBFilename = sDbPath & "\" & sDbName & "." & sDbType
    if (fso.FileExists(sStubMDBFilename)) Then
        WScript.Echo "Restoring Stub:" & vbcrlf & "Copying " & sStubMDBFilename & " to " & sMDBFilename
        fso.CopyFile sStubMDBFilename, sNewMDBFilename  
    end if
End Function

Function importModulesTxt(sMDBFilename, sImportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    ' Build file and pathnames
    dim myType, myName, myPath
    myType = fso.GetExtensionName(sMDBFilename)
    myName = fso.GetBaseName(sMDBFilename)
    myPath = fso.GetParentFolderName(sMDBFilename)

    ' if no path was given as argument, use a relative directory
    If (sImportpath = "") then
        sImportpath = myPath & "\src\"
    End If

    ' check for existing file and ask to overwrite with the stub
    if (fso.FileExists(sMDBFilename)) Then
        dim choice
        choice = MsgBox("The file: " & vbcrlf & sMDBFilename & vbcrlf & "already exists. Overwrite?", _
                        VbYesNo + VBQuestion, "File Already Exists!")
        if (choice <> VbYes) Then
            WScript.Quit
        end if
        
        fso.CopyFile sMDBFilename, sMDBFilename & ".bak"
        fso.DeleteFile sMDBFilename
    end if
    
    restoreStubIfPresent myPath, myName, myType

    ' launch MSAccess
    'WScript.Echo "Starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    'WScript.Echo "opening " & sMDBFilename & " ..."
    'oApplication.OpenAccessProject sMDBFilename
    if (fso.FileExists(sMDBFilename)) Then
        oApplication.OpenCurrentDatabase sMDBFilename
    else
        'WScript.echo "Creating NEW DATABASE!: " & sMDBFilename
        oApplication.NewCurrentDatabase sMDBFilename
    end if
    oApplication.Visible = false
    set newDb = oApplication.CurrentDb
    
    Dim folder
    Set folder = fso.GetFolder(sImportpath)
    
    ' Import __REFERENCES__ table
    if fso.FileExists(sImportPath & "\" & referencesTableName & ".xml") then
        oApplication.ImportXML sImportPath & "\" & referencesTableName & ".xml", acStructureAndData
        ' Restore the references
        RestoreReferences oApplication, newDb
        ' Delete __REFERENCES__ table
        'WScript.echo "Deleting References table!"
        deleteTable oApplication, newDb, referencesTableName
    end if
    
    ' Import __PROPERTIES__ table
    if fso.FileExists(sImportPath & "\" & propertiesTableName & ".xml") then
        oApplication.ImportXML sImportPath & "\" & propertiesTableName & ".xml", acStructureAndData
        ' Restore the properties
        RestoreAllProperties newDb
        ' Delete __PROPERTIES__ table
        'WScript.echo "Deleting Properties table!"
        deleteTable oApplication, newDb, propertiesTableName
    end if
    
    ' Import __TABLES__ table
    if fso.FileExists(sImportPath & "\" & tableListTableName & ".xml") then
        oApplication.ImportXML sImportPath & "\" & tableListTableName & ".xml", acStructureAndData
        ' Relink tables
        Dim numRelinkedTables
        'WScript.echo "About to relink tables!"
        numRelinkedTables = RelinkTables(oApplication)
        'WScript.echo "Relinked " & numRelinkedTables & " tables"
        ' Delete __TABLES__ table
        'WScript.echo "Deleting TableList table!"
        deleteTable oApplication, newDb, tableListTableName
    end if
    
    ' load each file from the import path into the stub
    Dim myFile, objectname, objecttype
    for each myFile in folder.Files
        objecttype = LCase(fso.GetExtensionName(myFile.Name))
        objectname = getObjectName(fso.GetBaseName(myFile.Name))
        'WScript.Echo "  " & objectname & " (" & objecttype & ")"
        
        if (objectName = tableListTableName) or _
           (objectName = propertiesTableName) or _
           (objectName = referencesTableName) or _
           (objectName = relationsTableName) then
           'Ignore file.
           'Wscript.echo "Ignoring " & objectName
        'Table files with structure and data
        ElseIf (objecttype = "xml") then
            oApplication.ImportXML myFile.Path, acStructureAndData
        'Table files with structure only
        ElseIf (objecttype = "xsd") then
            oApplication.ImportXML myFile.Path, acStructureOnly
            'Also import the corresponding data file
            dim dataFileName
            dataFileName = folder.Path & "\" & "Tbl_" & objectName & ".txt"
            'WScript.Echo "Importing data from " & dataFileName
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
    
    ' Import __RELATIONS__ table
    if fso.FileExists(sImportPath & "\" & relationsTableName & ".xml") then
        oApplication.ImportXML sImportPath & "\" & relationsTableName & ".xml", acStructureAndData
        ' Re-relate tables
        Dim numRelations
        numRelations = RerelateTables(oApplication)
        'WScript.echo "Created " & numRelations & " relations"
        ' Delete __RELATIONS__ table
        'WScript.echo "Deleting Relationships table!"
        deleteTable oApplication, newDb, relationsTableName
    end if
    
    oApplication.Visible = false
    oApplication.RunCommand acCmdCompileAndSaveAllModules
    oApplication.Visible = false
    oApplication.RunCommand acCmdCompactDatabase
    oApplication.Visible = True
    'oApplication.Quit
    'set oApplication = Nothing
End Function

Private Function deleteTable(app, db, tableNameToDelete)
    db.TableDefs.Refresh
    deleteTable = False
    Dim attemptsLeft
    attemptsLeft = 20
    On Error Resume Next
    Do While (attemptsLeft > 0)
        'WScript.echo "About to check if " & tableNameToDelete & " exists"
        If (db.TableDefs(tableNameToDelete).Name <> tableNameToDelete) Then
            'Error Thrown - Table doesn't exist! Awesome!
            'WScript.echo "It worked!"
            attemptsLeft = -1
            deleteTable = True
        Else
            'WScript.echo "About to try to delete " & tableNameToDelete
            attemptsLeft = attemptsLeft - 1
            app.DoCmd.DeleteObject acTable, tableNameToDelete
        End If
    Loop
End Function

Private Function getObjectName(sFileBaseName)
    if (sFileBaseName <> tableListTableName) and _
       (sFileBaseName <> propertiesTableName) and _
       (sFileBaseName <> referencesTableName) and _
       (sFileBaseName <> relationsTableName) then
        getObjectName = mid(sFileBaseName, filenamePrefixLength + 1)
    else
        getObjectName = sFileBaseName
    end if
End Function

Private Function RelinkTables(app)
    'Iterate through the newly-created tablelist to relink all linked tables.
    
    'Tbl_ID, Tbl_Name, Tbl_SourceTableName, Tbl_Connect, Tbl_ExportSchema, Tbl_ExportData, _
    'Tbl_DispType, Tbl_Attributes, Tbl_System, Tbl_Hidden, Tbl_AttachedTable, Tbl_AttachedODBC, _
    'Tbl_AttachSavePWD, Tbl_AttachExclusive, Tbl_ContainsBinary
    
    'WScript.Echo db.name
    Dim db
    set db = app.CurrentDb
    Dim newTable
    Dim TableCount
    Dim TableList
    Dim TableName
    Set TableList = db.OpenRecordset("SELECT * FROM " & tableListTableName, dbOpenSnapshot)
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
                    TableName = TableList("Tbl_Name")
                    set newTable = db.CreateTableDef(TableName, 0, _
                                                 TableList("Tbl_SourceTableName"), _
                                                 TableList("Tbl_Connect"))
                                                 '(Attributes: TableList("Tbl_Attributes"), _
                    'Check if this fails and provide a method to locate the
                    'missing database (ie FilePicker)
                    On Error Resume Next
                    db.TableDefs.Append newTable
                    
                    'Expected Error:  ODBC--connection to '[DB Filename]' failed'
                    '         Code:   800A0C4F
                    '         Source: DAO.TableDef
                    'Example ConnectString:
                    ' "ODBC;DSN=FB_TEST_DB;Driver=Firebird/InterBase(r) driver;Dbname=C:\Documents and Settings\Matt\My Documents\Projects\FB_TEST_DB.FDB;CHARSET=ASCII;"
                    'OR       Error:  Could not find file '[MDB Path and filename]'.
                    '         Code:   800A0BD0
                    '         Source: DAO.TableDef
                    'Example ConnectString:
                    ' ";DATABASE=C:\Documents and Settings\Matt\My Documents\Projects\MattsVCS-Access\MattsVCS-Access-Addin\test\AnotherTestDB.mdb"
                    Dim cancelClicked
                    cancelClicked = false
                    Dim newConnString
                    newConnString = ""
                    Do while (db.TableDefs(TableName).Name <> TableName)
                    'This should throw an error if the table doesn't exist (because of a linking failure)
                        If not cancelClicked Then
                            Wscript.echo "Could not re-establish link to '" & TableName & "' table" & vbcrlf & _
                                         "The connection command was:" & vbcrlf & _
                                         TableList("Tbl_Connect")
                            newConnString = getNewConnectionString(TableList("Tbl_Connect"), app, TableName)
                            If (newConnString <> "") Then
                                newTable.Connect = newConnString
                                db.TableDefs.Append newTable
                            Else
                                cancelClicked = true
                            End If
                        Else
                            Exit Do
                        End if
                    Loop
                    
                    On Error Resume Next
                    If (db.TableDefs(TableName).Name <> TableName) Then
                        Wscript.echo "Skipping " & TableName
                    Else
                        TableCount = TableCount + 1
                    End If
                    On Error Goto 0
                    
                End If
                TableList.MoveNext
            Wend
        End If
    TableList.Close
    
    RelinkTables = TableCount

End Function

Private Function RestoreReferences(app, db)
    'Iterate through the newly-created reference list to restore them.
    
    'RelName, RelTable, RelForeignTable, RelFieldName, RelForeignFieldName, RelAttributes, RelPartialReplica

    Dim refCount
    Dim RefList
    
    Set RefList = db.OpenRecordset("SELECT * FROM " & referencesTableName, dbOpenSnapshot)
        If Not RefList.EOF Then
            RefList.MoveFirst
            While Not RefList.EOF
                'Create new Reference
                'WScript.Echo "Creating new Reference" & vbcrlf & _
                '             "Ref_Name: " & RefList("Ref_Name") & vbcrlf &  _
                '             "Ref_GUID: " & RefList("Ref_GUID") & vbcrlf & _
                '             "Ref_Major: " & RefList("Ref_Major") & vbcrlf & _
                '             "Ref_Minor: " & RefList("Ref_Minor")
                On Error Resume Next
                app.References.AddFromGuid RefList("Ref_GUID"), _
                                           RefList("Ref_Major"), _
                                           RefList("Ref_Minor")
                refCount = refCount + 1
                On Error Goto 0
                RefList.MoveNext
            Wend
        End If
    RefList.Close
    RestoreReferences = refCount

End Function

Private Function RerelateTables(app)
    'Iterate through the newly-created relationslist to re-relate all related tables.
    
    'RelName, RelTable, RelForeignTable, RelFieldName, RelForeignFieldName, RelAttributes, RelPartialReplica
    
    Dim db
    set db = app.CurrentDb
    Dim newRel
    Dim newField
    Dim relCount
    Dim RelList
    
    Set RelList = db.OpenRecordset("SELECT * FROM " & relationsTableName, dbOpenSnapshot)
        If Not RelList.EOF Then
            RelList.MoveFirst
            While Not RelList.EOF
                'Create new Relation and append it
                'WScript.Echo "Creating new Relation" & vbcrlf & _
                '             "RelName: " & RelList("RelName") & vbcrlf &  _
                '             "RelTable: " & RelList("RelTable") & vbcrlf & _
                '             "RelForeignTable: " & RelList("RelForeignTable") & vbcrlf & _
                '             "RelFieldName: " & RelList("RelFieldName") & vbcrlf & _
                '             "RelForeignFieldName: " & RelList("RelForeignFieldName") & vbcrlf & _
                '             "RelAttributes: " & RelList("RelAttributes") & vbcrlf & _
                '             "RelPartialReplica: " & RelList("RelPartialReplica")
                set newRel = db.CreateRelation(RelList("RelName"), _
                                               RelList("RelTable"), _
                                               RelList("RelForeignTable"), _
                                               RelList("RelAttributes"))
                'WScript.Echo "New Relation Created: " & newRel.name
                'WScript.Echo "Current Relation Count: " & db.Relations.Count
                set newField = newRel.CreateField(RelList("RelFieldName"))
                newRel.Fields.Append newField
                newField.ForeignName = RelList("RelForeignFieldName")
                'newRel.PartialReplica = RelList("RelPartialReplica")
                
                'Access creates hidden indexes automatically for any fields ending with ID.
                'Also, if two tables were related in the original database, there will be
                'an index called '[ForeignFieldName][FieldName]' in the Foreign Table.
                'Either way will throw error
                ' 3284 Index already exists
                'when trying to append the new relationship.
                'Therefore, delete it first if it exists.
                'More information can be found in this link:
                'http://groups.google.ie/group/microsoft.public.access/browse_thread/thread/ca58ce291bdc62df?hl=en&ie=UTF-8&q=create+relation+3284+Index+already+exists
                On Error Resume Next
                db.TableDefs(RelList("RelForeignTable")).Indexes.Delete RelList("RelName")
                On Error Goto 0
                
                db.Relations.Append newRel
                relCount = relCount + 1
                    
                RelList.MoveNext
            Wend
        End If
    RelList.Close
    RerelateTables = relCount

End Function

public sub RestoreAllProperties(db)
    RestoreProperties db, "StartupProperties"
    RestoreProperties db, "SummaryInfo"
End sub

Public Sub RestoreProperties(db, collectionName)
    ' CollectionName must be "StartupProperties" or "SummaryInfo"
    
    Dim obj
    Dim c
    Dim PropList
    Dim SQLStr
    
    Dim propertyValue
    Dim propertyCount
    
    Select Case collectionName
        Case "SummaryInfo"
            Set c = db.Containers("Databases")
            Set obj = c.Documents("SummaryInfo")
        Case "StartupProperties"
            Set obj = db
    End Select
    
    SQLStr = "SELECT * FROM " & propertiesTableName & _
             " WHERE PropCollection = '" & collectionName & "'"
    
    Set PropList = db.OpenRecordset(SQLStr, dbOpenSnapshot)
        If Not PropList.EOF Then
            PropList.MoveFirst
        End If
        While Not PropList.EOF
            Select Case PropList("PropType")
                Case dbBinary: propertyValue = PropList("PropValueBinary")
                Case dbBoolean: propertyValue = PropList("PropValueBoolean")
                Case dbByte: propertyValue = PropList("PropValueByte")
                Case dbCurrency: propertyValue = PropList("PropValueCurrency")
                Case dbDate: propertyValue = PropList("PropValueDate")
                Case dbDouble: propertyValue = PropList("PropValueDouble")
                Case dbGUID: propertyValue = PropList("PropValueGUID")
                Case dbInteger: propertyValue = PropList("PropValueInteger")
                Case dbLong: propertyValue = PropList("PropValueLong")
                Case dbLongBinary: propertyValue = PropList("PropValueBinary")
                Case dbMemo: propertyValue = PropList("PropValueMemo")
                Case dbSingle: propertyValue = PropList("PropValueSingle")
                Case dbText: propertyValue = PropList("PropValueText")
                Case Else: MsgBox "Unexpected Property Type: " & PropList("PropType") & "!"
            End Select
            SetProperty obj, PropList("PropName"), propertyValue, PropList("PropType")
            propertyCount = propertyCount + 1
        PropList.MoveNext
        Wend
    PropList.Close
    Exit Sub
End Sub

Sub SetProperty(objParent, strName, varValue, lngType)
    Dim prpNew
    'Title, Author and Company work first time. All others must be created.
    
    ' Attempt to set the specified property.
    On Error Resume next
        'Wscript.echo "Setting " & strName & " to [" & varValue & "]"
        objParent.Properties(strName) = varValue
        
        ' If that failed, create a new property and append it.
        'WScript.echo "It worked! " & strName & ": " & objParent.Properties(strName)
        If (objParent.Properties(strName) <> varValue) Then
            'WScript.echo "It didn't work so we have to create the property " & strName
            'WScript.echo strName & ": " & objParent.Properties(strName)
            Set prpNew = objParent.CreateProperty(strName, _
                lngType, varValue)
            objParent.Properties.Append prpNew
            'WScript.echo "It worked that time! " & strName & ": " & objParent.Properties(strName)
        else
            'Wscript.echo "Yep, it worked."
        End If
    Exit Sub
End Sub

Private Function getNewConnectionString(currentConnectString, app, tableName)
    ' "ODBC;DSN=FB_TEST_DB;Driver=Firebird/InterBase(r) driver;Dbname=C:\Documents and Settings\Matt\My Documents\Projects\FB_TEST_DB.FDB;CHARSET=ASCII;"
    ' ";DATABASE=C:\Documents and Settings\Matt\My Documents\Projects\MattsVCS-Access\MattsVCS-Access-Addin\test\AnotherTestDB.mdb"
    
    Dim startPos
    Dim endPos
    Dim oldFilename
    Dim newFilename
    
    If (left(currentConnectString, 5) = "ODBC;") Then
    'ODBC Database - we don't change the DSN Name but it still works. (0_o)
        startPos = InStr(currentConnectString, "Dbname=")
        If (startPos > 0) Then
            endPos = InStr(startPos, currentConnectString, ";")
            oldFilename = mid(currentConnectString, startPos + 7, endPos - startPos - 7)
            newFilename = pickFile(app, oldFileName, tableName)
            If (newFilename <> "") Then
                getNewConnectionString = _
                    Replace(currentConnectString, oldFileName, newFilename)
            Else
                getNewConnectionString = ""
            End If
        End If
    ElseIf (left(currentConnectString, 1) = ";") Then
    'MDB Database
        oldFilename = mid(currentConnectString, 11)
        newFilename = pickFile(app, oldFileName, tableName)
        If (newFilename <> "") Then
            getNewConnectionString = _
                Replace(currentConnectString, oldFileName, newFilename)
        Else
            getNewConnectionString = ""
        End If
    Else
        'Unrecognised Connection String
        WScript.Echo "Unrecognised database type in connection string:" & vbcrlf & _
                     currentConnectString
        getNewConnectionString = "" '?????
    End If
    wscript.echo "Old/New Connection Strings:" & vbcrlf & _
                 currentConnectString & vbcrlf & _
                 getNewConnectionString
End Function

Private Function pickFile(app, oldFileName, tableName)
    pickFile = ""
    Dim fd
    Set fd = app.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Filters.Add "Databases", "*." & fso.GetExtensionName(oldFileName), 1
    fd.InitialView = msoFileDialogViewDetails
    Dim folderName
    folderName = fso.GetParentFolderName(oldFileName)
    If fso.FolderExists(folderName) then
        fd.InitialFileName = folderName
    else
        fd.InitialFileName = fso.GetParentFolderName(app.currentDb.Name)
    End If
    fd.Title = "Select the " & fso.GetExtensionName(oldFileName) & " database containing the '" & _
               tableName & "' table (previously " & fso.GetBaseName(oldFileName) & ")"
    
    Dim vrtSelectedItem
    'Use the Show method to display the File Picker dialog box and return the user's action.
    If fd.Show = -1 Then
        'The user pressed the action button.
        'Get the first string in the FileDialogSelectedItems collection.
        If fd.SelectedItems(1) <> "" Then
            pickFile = fd.SelectedItems(1)
        End If
    Else
        'The user pressed Cancel.
    End If
    
    Set fd = Nothing
    app.Visible = False
end Function
