' Script to copy AccessVCS AddIn file into AddIn dir
' Author Matt Fisher mrpfisher@gmail.com
' Created 28 Oct 2009
' ----------------------------------------------------------
Option Explicit

Dim objFSO, objFileCopy
Dim strFilePath, strDestination, strAppDataPath
Set objFSO = CreateObject("Scripting.FileSystemObject")

strFilePath = ".\AccessVCAddIn.mdb"

Const APPLICATION_DATA = &H1a&
Dim objShell, objFolder
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(APPLICATION_DATA)
strAppDataPath = objFolder.Self.Path
strDestination = strAppDataPath & "\Microsoft\AddIns\AccessVCAddIn.mda"

if objFSO.FileExists(strFilePath) then
	Set objFileCopy = objFSO.GetFile(strFilePath)
	setLibraryLocation objFileCopy.Path, strDestination
	if objFSO.FileExists(strDestination) then
		Set objFileCopy = objFSO.GetFile(strDestination)
		objFileCopy.Copy (strDestination & ".bak") 'Back up any existing file
	end if
	objFileCopy.Copy (strDestination)
else
	WScript.Echo strFilePath & " does not exist."
end if

WScript.Echo "Matt's AccessVCS Successfully Installed!" & vbcrlf & vbcrlf & _
			 "You'll now find it in Access under:" & vbcrlf & _
			 "Tools -> Add-Ins -> Add-In Manager"
WScript.Quit


sub setLibraryLocation(dbFilePath, targetLocation)
dim SQLStr
	SQLStr = "UPDATE USysRegInfo SET USysRegInfo.Value = '" & targetLocation & "'" & _
			 " WHERE (USysRegInfo.ValName = 'Library')"
	Dim oApplication
	Set oApplication = CreateObject("Access.Application")
	'wscript.echo SQLStr
	oApplication.OpenCurrentDatabase dbFilePath
	oApplication.CurrentDb.Execute SQLStr
	oApplication.Quit 1 '(acQuitSaveAll)
end sub
