Option Compare Database
Option Explicit

Public Const PROPERTY_LIST_TABLENAME As String = "__PROPERTIES__BE"
Public Const PROPERTY_LIST_FILENAME As String = "__PROPERTIES__.xml"
'Those are underscores

'Properties are associated with a database.
'Options are associated with the Application itself (Access)
'Thus it may cause problems if changes are automatically made to options.

Private Function getStartupPropertyCollection() As Collection

Dim startupProperties As New Collection '(All strings)
startupProperties.Add "AppTitle" 'Application Title
startupProperties.Add "AppIcon" 'Application Icon
startupProperties.Add "StartupForm" 'Display Form/Page
startupProperties.Add "StartupShowDBWindow" 'Display Database Window
startupProperties.Add "StartupShowStatusBar" 'Display Status Bar
startupProperties.Add "StartupMenuBar" 'Menu Bar
startupProperties.Add "StartupShortcutMenuBar" 'Shortcut Menu Bar
startupProperties.Add "AllowFullMenus" 'Allow Full Menus
startupProperties.Add "AllowShortcutMenus" 'Allow Default Shortcut Menus
startupProperties.Add "AllowBuiltInToolbars" 'Allow Built-In Toolbars
startupProperties.Add "AllowToolbarChanges" 'Allow Toolbar/Menu Changes
startupProperties.Add "AllowBreakIntoCode" 'Allow Viewing Code After Error
startupProperties.Add "AllowSpecialKeys" 'Use Access Special Keys

Set getStartupPropertyCollection = startupProperties
End Function

'?Replicable
'?ReplicationConflictFunction

Private Function getSummaryPropertyCollection() As Collection

Dim summaryProperties As New Collection '(All strings)
summaryProperties.Add "Title"
summaryProperties.Add "Author"
summaryProperties.Add "Company"
summaryProperties.Add "Comments"
summaryProperties.Add "Subject"
summaryProperties.Add "Keywords"
summaryProperties.Add "Manager"
summaryProperties.Add "Category"
summaryProperties.Add "Hyperlink base"

Set getSummaryPropertyCollection = summaryProperties
End Function

Private Function ListProperties(ByRef objectWithProps As Object, _
                                ByRef propertyCollection As Collection, _
                                ByRef propertyCollectionName As String, _
                                ByRef PropList As DAO.Recordset) As Integer
Dim prop As Property
Dim propertyCount As Integer
Dim propertyName As Variant

For Each propertyName In propertyCollection
    Set prop = getProperty(propertyName, objectWithProps)
    If Not (prop Is Nothing) Then
        PropList.AddNew
            PropList("PropCollection") = propertyCollectionName
            PropList("PropName") = prop.Name
            PropList("PropType") = prop.Type
            Select Case prop.Type
                Case dbBinary: PropList("PropValueBinary") = prop.Value
                Case dbBoolean: PropList("PropValueBoolean") = prop.Value
                Case dbByte: PropList("PropValueByte") = prop.Value
                Case dbCurrency: PropList("PropValueCurrency") = prop.Value
                Case dbDate: PropList("PropValueDate") = prop.Value
                Case dbDouble: PropList("PropValueDouble") = prop.Value
                Case dbGUID: PropList("PropValueGUID") = prop.Value
                Case dbInteger: PropList("PropValueInteger") = prop.Value
                Case dbLong: PropList("PropValueLong") = prop.Value
                Case dbLongBinary: PropList("PropValueBinary") = prop.Value
                Case dbMemo: PropList("PropValueMemo") = prop.Value
                Case dbSingle: PropList("PropValueSingle") = prop.Value
                Case dbText: PropList("PropValueText") = prop.Value
                Case Else: MsgBox "Unexpected Property Type: " & prop.Type & "!"
            End Select
        PropList.Update
        propertyCount = propertyCount + 1
    End If
Next propertyName

End Function

Public Function ListAllProperties() As Integer
On Error GoTo ErrProc
Dim db As Database
Dim obj As Object
Dim c As Container
Dim PropList As DAO.Recordset
Dim i As Integer
Dim propertyName As Variant
Dim propertyCount As Integer
Dim propertyCollection As Collection
Set propertyCollection = getSummaryPropertyCollection()

Set db = Access.CurrentDb
Log "Taking stock of datatabase properties"

CodeDb.Execute "DELETE * FROM " & PROPERTY_LIST_TABLENAME

Set PropList = CodeDb.OpenRecordset(PROPERTY_LIST_TABLENAME, dbOpenDynaset)
    If Not PropList.EOF Then
        PropList.MoveFirst
    End If
    Set c = db.Containers("Databases")
    Set obj = c.Documents("SummaryInfo")
    propertyCount = propertyCount + ListProperties(obj, _
                        propertyCollection, "SummaryInfo", PropList)
    Set obj = db
    Set propertyCollection = getStartupPropertyCollection()
    propertyCount = propertyCount + ListProperties(obj, _
                        propertyCollection, "StartupProperties", PropList)

PropList.Close

ListAllProperties = propertyCount
Exit Function
ErrProc:
DispErrMsgGSb Error$, "document the SummaryInfo properties"
ListAllProperties = -1
End Function

Public Sub RestoreSummaryAndStartupProperties()
RestoreProperties "StartupProperties"
RestoreProperties "SummaryInfo"
End Sub

Public Sub RestoreProperties(collectionName)
' CollectionName must be "StartupProperties" or "SummaryInfo"
On Error GoTo ErrProc

Dim db As Database
Dim obj As Object
Dim c As Container
Dim PropList As DAO.Recordset
Dim SQLStr As String

Dim propertyValue As Variant
Dim propertyCount As Integer

Set db = Access.CurrentDb

Select Case collectionName
    Case "SummaryInfo"
        Set c = db.Containers("Databases")
        Set obj = c.Documents("SummaryInfo")
    Case "StartupProperties"
        Set obj = db
End Select

SQLStr = "SELECT * FROM " & PROPERTY_LIST_TABLENAME & _
         " WHERE PropCollection = " & collectionName

Set PropList = CodeDb.OpenRecordset(PROPERTY_LIST_TABLENAME, dbOpenDynaset)
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
ErrProc:
    DispErrMsgGSb Error$, "restore the database properties from a table."
End Sub

Sub SetProperty(objParent As Object, strName As String, _
                varValue As Variant, lngType As Long)

    Dim prpNew As Property
    Dim errLoop As Error
    Dim varType As String
    
    ' Attempt to set the specified property.
    On Error GoTo Err_Property
        objParent.Properties(strName) = varValue
    On Error GoTo 0
    
    Exit Sub

Err_Property:
    ' Error 3270 means that the property was not found.
    If DBEngine.Errors(0).number = 3270 Then
    ' Create property, set its value, and append it to the
    ' Properties collection.
        Set prpNew = objParent.CreateProperty(strName, _
            lngType, varValue)
        objParent.Properties.Append prpNew
        Resume Next
    Else
    ' If different error has occurred, display message.
        For Each errLoop In DBEngine.Errors
            MsgBox "Error number: " & errLoop.number & vbCr & _
                errLoop.Description
        Next errLoop
        End
    End If

End Sub

Private Function getProperty(ByVal propName As String, ByRef obj As Object) As Property
Set getProperty = Nothing

' Attempt to get the specified property.
On Error Resume Next
Set getProperty = obj.Properties(propName)

ExitProc:
End Function

Public Sub test_ListProperties()
Debug.Print "Documented " & ListAllProperties() & " properties."
End Sub

Public Function ListAllOptions() As Integer
On Error GoTo ErrProc
Dim db As Database
Dim OptionList As DAO.Recordset
Dim i As Integer
Dim optionName As Variant
Dim optionValue As Variant
Dim optionCount As Integer
Dim optionCollection As Collection
Set optionCollection = getOptionCollection()

'Set db = Access.CurrentDb
Log "Taking stock of datatabase options"

Set OptionList = CodeDb.OpenRecordset(PROPERTY_LIST_TABLENAME, dbOpenDynaset)
    If Not OptionList.EOF Then
        OptionList.MoveLast
    End If
    For Each optionName In optionCollection
    optionValue = Application.GetOption(optionName)
    OptionList.AddNew
        OptionList("PropCollection") = "ApplicationOptions"
        OptionList("PropName") = optionName
        OptionList("PropType") = varType(optionValue)
        Select Case varType(optionValue)
            Case vbInteger: OptionList("PropValueInteger") = optionValue
            Case vbLong: OptionList("PropValueLong") = optionValue
            Case vbString: OptionList("PropValueText") = optionValue
            Case Else: MsgBox "Unexpected Option Type: " & varType(optionValue) & "!"
        End Select
    OptionList.Update
    optionCount = optionCount + 1
Next optionName
OptionList.Close

ListAllOptions = optionCount
Exit Function
ErrProc:
DispErrMsgGSb Error$, "document the Database Options"
ListAllOptions = optionCount
End Function


Private Function getOptionCollection() As Collection
'LIST OF ALL DATABASE OPTIONS:

Dim options As New Collection '(All strings)
'View Tab
    options.Add "Show Status Bar" ' Show, Status bar
    options.Add "Show Startup Dialog Box" ' Show, Startup Task Pane
    options.Add "Show New Object Shortcuts" ' Show, New object shortcuts
    options.Add "Show Hidden Objects" ' Show, Hidden objects
    options.Add "Show System Objects" ' Show, System objects
    options.Add "ShowWindowsInTaskbar" ' Show, Windows in Taskbar
    options.Add "Show Macro Names Column" ' Show in Macro Design, Names column
    options.Add "Show Conditions Column" ' Show in Macro Design, Conditions column
    options.Add "Database Explorer Click Behavior" ' Click options in database window
'General Tab
    options.Add "Left Margin" ' Print margins, Left margin
    options.Add "Right Margin" ' Print margins, Right margin
    options.Add "Top Margin" ' Print margins, Top margin
    options.Add "Bottom Margin" ' Print margins, Bottom margin
    options.Add "Four-Digit Year Formatting" ' Use four-year digit year formatting, This database
    options.Add "Four-Digit Year Formatting All Databases" ' Use four-year digit year formatting, All databases
    options.Add "Track Name AutoCorrect Info" ' Name AutoCorrect, Track name AutoCorrect info
    options.Add "Perform Name AutoCorrect" ' Name AutoCorrect, Perform name AutoCorrect
    options.Add "Log Name AutoCorrect Changes" ' Name AutoCorrect, Log name AutoCorrect changes
    options.Add "Enable MRU File List" ' Recently used file list
    options.Add "Size of MRU File List" ' Recently used file list, (number of files)
    options.Add "Provide Feedback with Sound" ' Provide feedback with sound
    options.Add "Auto Compact" ' Compact on Close
    options.Add "New Database Sort Order" ' New database sort order
    options.Add "Remove Personal Information" ' Remove personal information from file properties on save
    options.Add "Default Database Directory" ' Default database folder
'Edit/Find Tab
    options.Add "Default Find/Replace Behavior" ' Default find/replace behavior
    options.Add "Confirm Record Changes" ' Confirm, Record changes
    options.Add "Confirm Document Deletions" ' Confirm, Document deletions
    options.Add "Confirm Action Queries" ' Confirm, Action queries
    options.Add "Show Values in Indexed" ' Show list of values in, Local indexed fields
    options.Add "Show Values in Non-Indexed" ' Show list of values in, Local nonindexed fields
    options.Add "Show Values in Remote" ' Show list of values in, ODBC fields
    options.Add "Show Values in Snapshot" ' Show list of values in, Records in local snapshot
    options.Add "Show Values in Server" ' Show list of values in, Records at server
'Datasheet Tab
    options.Add "Default Font Color" ' Default colors, Font
    options.Add "Default Background Color" ' Default colors, Background
    options.Add "Default Gridlines Color" ' Default colors, Gridlines
    options.Add "Default Gridlines Horizontal" ' Default gridlines showing, Horizontal
    options.Add "Default Gridlines Vertical" ' Default gridlines showing, Vertical
    options.Add "Default Column Width" ' Default column width
    options.Add "Default Font Name" ' Default font, Font
    options.Add "Default Font Weight" ' Default font, Weight
    options.Add "Default Font Size" ' Default font, Size
    options.Add "Default Font Underline" ' Default font, Underline
    options.Add "Default Font Italic" ' Default font, Italic
    options.Add "Default Cell Effect" ' Default cell effect
    options.Add "Show Animations" ' Show animations
    options.Add "Show Smart Tags on Datasheets" ' Show Smart Tags on Datasheets
'Keyboard Tab
    options.Add "Move After Enter" ' Move after enter
    options.Add "Behavior Entering Field" ' Behavior entering field
    options.Add "Arrow Key Behavior" ' Arrow key behavior
    options.Add "Cursor Stops at First/Last Field" ' Cursor stops at first/last field
    options.Add "Ime Autocommit" ' Auto commit
    options.Add "Datasheet Ime Control" ' Datasheet IME control
'Tables/Queries Tab
    options.Add "Default Text Field Size" ' Table design, Default field sizes - Text
    options.Add "Default Number Field Size" ' Table design, Default field sizes - Number
    options.Add "Default Field Type" ' Table design, Default field type
    options.Add "AutoIndex on Import/Create" ' Table design, AutoIndex on Import/Create
    options.Add "Show Table Names" ' Query design, Show table names
    options.Add "Output All Fields" ' Query design, Output all fields
    options.Add "Enable AutoJoin" ' Query design, Enable AutoJoin
    options.Add "Run Permissions" ' Query design, Run permissions
    options.Add "ANSI Query Mode" ' Query design, SQL Server Compatible Syntax (ANSI 92) - This database
    options.Add "ANSI Query Mode Default" ' Query design, SQL Server Compatible Syntax (ANSI 92) - Default for new databases
    options.Add "Query Design Font Name" ' Query design, Query design font, Font
    options.Add "Query Design Font Size" ' Query design, Query design font, Size
    options.Add "Show Property Update Options buttons" ' Show Property Update Options buttons
'Forms/Reports Tab
    options.Add "Selection Behavior" ' Selection behavior
    options.Add "Form Template" ' Form template
    options.Add "Report Template" ' Report template
    options.Add "Always Use Event Procedures" ' Always use event procedures
    options.Add "Show Smart Tags on Forms" ' Show Smart Tags on Forms
    options.Add "Themed Form Controls" ' Show Windows Themed Controls on Forms
'Advanced Tab
    options.Add "Ignore DDE Requests" ' DDE operations, Ignore DDE requests
    options.Add "Enable DDE Refresh" ' DDE operations, Enable DDE refresh
    options.Add "Default File Format" ' Default File Format
    options.Add "Default Open Mode for Databases" ' Default open mode
    options.Add "Command-Line Arguments" ' Command-line arguments
    options.Add "OLE/DDE Timeout (sec)" ' OLE/DDE timeout (sec)
    options.Add "Default Record Locking" ' Default record locking
    options.Add "Refresh Interval (sec)" ' Refresh interval (sec)
    options.Add "Number of Update Retries" ' Number of update retries
    options.Add "ODBC Refresh Interval (sec)" ' ODBC refresh interval (sec)
    options.Add "Update Retry Interval (msec)" ' Update retry interval (msec)
    options.Add "Use Row Level Locking" ' Open databases using record-level locking
'Pages Tab
    options.Add "Section Indent" ' Default Designer Properties, Section Indent
    options.Add "Alternate Row Color" ' Default Designer Properties, Alternate Row Color
    options.Add "Caption Section Style" ' Default Designer Properties, Caption Section Style
    options.Add "Footer Section Style" ' Default Designer Properties, Footer Section Style
    options.Add "Use Default Page Folder" ' Default Database/Project Properties, Use Default Page Folder
    options.Add "Default Page Folder" ' Default Database/Project Properties, Default Page Folder
    options.Add "Use Default Connection File" ' Default Database/Project Properties, Use Default Connection File
    options.Add "Default Connection File" ' Default Database/Project Properties, Default Connection File
'Spelling Tab
    options.Add "Spelling dictionary language" ' Dictionary Language
    options.Add "Spelling add words to" ' Add words to
    options.Add "Spelling suggest from main dictionary only" ' Suggest from main dictionary only
    options.Add "Spelling ignore words in UPPERCASE" ' Ignore words in UPPERCASE
    options.Add "Spelling ignore words with number" ' Ignore words with numbers
    options.Add "Spelling ignore Internet and file addresses" ' Ignore Internet and file addresses
    options.Add "Spelling use German post-reform rules" ' Language-specific, German: Use post-reform rules
    options.Add "Spelling combine aux verb/adj" ' Language-specific, Korean: Combine aux verb/adj.
    options.Add "Spelling use auto-change list" ' Language-specific, Korean: Search misused word list
    options.Add "Spelling process compound nouns" ' Language-specific, Korean: Process compound nouns
    options.Add "Spelling Hebrew modes" ' Language-specific, Hebrew modes
    options.Add "Spelling Arabic modes" ' Language-specific, Arabic modes
'International Tab
    options.Add "Default direction" ' Right-to-Left, Default direction
    options.Add "General alignment" ' Right-to-Left, General alignment
    options.Add "Cursor movement" ' Right-to-Left, Cursor movement
    options.Add "Use Hijri Calendar" ' Use Hijri Calendar
'Error Checking Tab
    options.Add "Enable Error Checking" ' Settings, Enable error checking
    options.Add "Error Checking Indicator Color" ' Settings, Error indicator color
    options.Add "Unassociated Label and Control Error Checking" ' Form/Report Design Rules, Unassociated label and control
    'This one doesn't work for some reason.
    'options.Add "New Unassociated Label Error Checking" ' Form/Report Design Rules, New unassociated labels
    options.Add "Keyboard Shortcut Errors Error Checking" ' Form/Report Design Rules, Keyboard shortcut errors
    options.Add "Invalid Control Properties Error Checking" ' Form/Report Design Rules, Invalid control properties
    options.Add "Common Report Errors Error Checking" ' Form/Report Design Rules, Common report errors

Set getOptionCollection = options

End Function

Public Sub test_show_dbConsts()
Debug.Print dbBoolean '1
Debug.Print dbByte '2
Debug.Print dbInteger '3
Debug.Print dbLong '4
Debug.Print dbCurrency '5
Debug.Print dbSingle '6
Debug.Print dbDouble '7
Debug.Print dbDate '8
Debug.Print dbBinary '9
Debug.Print dbText '10
Debug.Print dbLongBinary '11
Debug.Print dbMemo '12
Debug.Print dbGUID '15
End Sub