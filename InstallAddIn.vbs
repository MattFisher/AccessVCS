' Script to copy Matt's Access VCS AddIn file into AddIn dir
' Author Matt Fisher
' Created 28 Oct 2009
' ----------------------------------------------------------
Option Explicit

Dim objFSO, objFileCopy
Dim strFilePath, strDestination, strAppDataPath
Set objFSO = CreateObject("Scripting.FileSystemObject")

strFilePath = ".\AccessVCAddIn.mdb"
strAppDataPath = "C:\Documents and Settings\Matt\"
'strAppDataPath = "H:\profile\"
strDestination = strAppDataPath & "Application Data\Microsoft\AddIns\AccessVCAddIn.mda"

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
