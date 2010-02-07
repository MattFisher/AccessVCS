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
	if objFSO.FileExists(strDestination) then
		Set objFileCopy = objFSO.GetFile(strDestination)
		objFileCopy.Copy (strDestination & ".bak") 'Back up any existing file
	end if
	Set objFileCopy = objFSO.GetFile(strFilePath)
	objFileCopy.Copy (strDestination)
else
	WScript.Echo strFilePath & " does not exist."
end if

'WSCript.Echo objFileCopy.Name & " copied to " & strDestination
Wscript.Quit
