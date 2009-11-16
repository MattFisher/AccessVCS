' Usage:
'  CScript decompose.vbs <input file> <path>

' Converts all modules, classes, forms and macros from an Access Database file (.mdb) <input file> to
' text and saves the results in separate files to <path>.  Requires Microsoft Access.
'

Option Explicit

const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3

' BEGIN CODE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

dim sADPFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "Please include the arguments!", vbExclamation, "Error"
    Wscript.Quit()
End if
sADPFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sExportpath
If (WScript.Arguments.Count = 1) then
    sExportpath = ""
else
    sExportpath = WScript.Arguments(1)
End If


exportModulesTxt sADPFilename, sExportpath

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If

Function exportModulesTxt(sADPFilename, sExportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    dim myType, myName, myPath, sStubADPFilename
    myType = fso.GetExtensionName(sADPFilename)
    myName = fso.GetBaseName(sADPFilename)
    myPath = fso.GetParentFolderName(sADPFilename)

    If (sExportpath = "") then
        sExportpath = myPath & "\Source\"
    End If
    sStubADPFilename = sExportpath & myName & "_stub." & myType

    WScript.Echo "Copy stub from" & vbcrlf & _
				 sADPFilename & vbcrlf & _
				 "to" & vbcrlf & _
				 sStubADPFilename & "..."
    On Error Resume Next
        fso.CreateFolder(sExportpath)
    On Error Goto 0
    fso.CopyFile sADPFilename, sStubADPFilename

    WScript.Echo "starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "opening " & sStubADPFilename & " ..."
    'oApplication.OpenAccessProject sStubADPFilename
	oApplication.OpenCurrentDatabase sStubADPFilename

    oApplication.Visible = false

    dim dctDelete
    Set dctDelete = CreateObject("Scripting.Dictionary")
    WScript.Echo "exporting..."
    Dim myObj
	dim myForm
	Dim myProject
	set myProject = oApplication.currentProject
	WScript.Echo oApplication.currentDB.Containers("Forms").Documents.count & " forms found."
    dim myCollection
	set myCollection = myProject.AllForms
	'OR
	'set myCollection = oApplication.currentDB.Containers("Forms").Documents
	For Each myForm In myCollection
        WScript.Echo "  " & myForm.Name
        oApplication.SaveAsText acForm, myForm.name, sExportpath & "\" & myForm.name & ".form"
        oApplication.DoCmd.Close acForm, myForm.name
        dctDelete.Add "FO" & myForm.name, acForm
    Next
	WScript.Echo myProject.AllModules.count & " modules found."
    For Each myObj In myProject.AllModules
        WScript.Echo "  " & myObj.name
        oApplication.SaveAsText acModule, myObj.name, sExportpath & "\" & myObj.name & ".bas"
        dctDelete.Add "MO" & myObj.name, acModule
    Next
	WScript.Echo myProject.AllMacros.count & " macros found."
    For Each myObj In myProject.AllMacros
        WScript.Echo "  " & myObj.fullname
        oApplication.SaveAsText acMacro, myObj.fullname, sExportpath & "\" & myObj.fullname & ".mac"
        dctDelete.Add "MA" & myObj.fullname, acMacro
    Next
	WScript.Echo myProject.AllReports.count & " reports found."
    For Each myObj In myProject.AllReports
        WScript.Echo "  " & myObj.fullname
        oApplication.SaveAsText acReport, myObj.fullname, sExportpath & "\" & myObj.fullname & ".report"
        dctDelete.Add "RE" & myObj.fullname, acReport
    Next
	For i = 0 To oApplication.CurrentDB.QueryDefs.Count - 1
		oApplication.SaveAsText acQuery, oApplication.CurrentDB.QueryDefs(i).Name, sExportpath & "\" & oApplication.CurrentDB.QueryDefs(i).Name & ".txt"
	Next

    WScript.Echo "deleting..."
    dim sObjectname
    For Each sObjectname In dctDelete
        WScript.Echo "  " & Mid(sObjectname, 3)
        oApplication.DoCmd.DeleteObject dctDelete(sObjectname), Mid(sObjectname, 3)
    Next

    oApplication.CloseCurrentDatabase
    oApplication.CompactRepair sStubADPFilename, sStubADPFilename & "_"
    oApplication.Quit

    fso.CopyFile sStubADPFilename & "_", sStubADPFilename
    fso.DeleteFile sStubADPFilename & "_"


End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function
