Option Compare Database
Option Explicit
Dim ShellStr As String

'Modify the following constants to point to your installations of TortoiseGit and TortoiseSVN
Const TortoiseSvnPath = "C:\Program Files\TortoiseSVN\bin\TortoiseProc.exe"
Const TortoiseGitPath = "C:\Program Files\TortoiseGit\bin\TortoiseProc.exe"

Public Sub Commit2Subversion(fileOrFolder As String)
ShellStr = """" & TortoiseSvnPath & """" & " /command:commit" & _
                                           " /path:""" & fileOrFolder & """" & _
                                           " /closeonend:1" '& _
                                           " /logmsg:""test log message"""
GetShell.Run ShellStr, vbNormalFocus, False 'Don't wait for completion
End Sub

Public Sub UpdateFromSubversion(fileOrFolder As String)
ShellStr = """" & TortoiseSvnPath & """" & " /command:update" & _
                                           " /path:""" & fileOrFolder & """" & _
                                           " /closeonend:1"
'Debug.Print shellStr
GetShell.Run ShellStr, vbNormalFocus, True 'Wait for completion
End Sub

Public Sub Commit2Git(fileOrFolder As String)
ShellStr = """" & TortoiseGitPath & """" & " /command:commit" & _
                                           " /path:""" & fileOrFolder & """" & _
                                           " /closeonend:0" '& _
                                           " /logmsg:""test log message"""
GetShell.Run ShellStr, vbNormalFocus, True 'Wait for completion
End Sub

Public Sub FetchFromGit(fileOrFolder As String)
'Update From Svn doesn't really have an exact equivalent in Git, because the Working Path
'in Git contains the appropriate versions of the files.  'Fetch' is the closest equivalent.
'Though 'Pull', which is the same as fetch except merges local changes, is close.
ShellStr = """" & TortoiseGitPath & """" & " /command:fetch" & _
                                           " /path:""" & fileOrFolder & """" & _
                                           " /closeonend:1"
'Debug.Print shellStr
GetShell.Run ShellStr, vbNormalFocus, True 'Wait for completion
End Sub