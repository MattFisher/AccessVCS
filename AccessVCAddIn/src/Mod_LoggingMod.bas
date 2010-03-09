Option Compare Database
Option Explicit
Dim watchingListBox As ListBox

Public Sub Log(message As String, Optional writeToFile = False)
Dim LogRS As DAO.Recordset
Set LogRS = CodeDb.OpenRecordset("SELECT * FROM __LogTbl__", dbOpenDynaset, dbAppendOnly)
    LogRS.AddNew
        LogRS("Log_Timestamp") = Now()
        LogRS("Log_Message") = message
    LogRS.Update
LogRS.Close

If Not (watchingListBox Is Nothing) Then
    If (watchingListBox.Parent.Parent.Parent.CurrentView <> 0) Then 'Not in design view
        watchingListBox.Requery
        watchingListBox.Selected(watchingListBox.ListCount - 1) = True
        watchingListBox.Parent.Parent.Parent.Repaint
    End If
End If
  
End Sub

Public Sub ClearLog()
    CodeDb.Execute ("DELETE * FROM __LogTbl__")
End Sub

Public Sub SetLogWatchingListBox(aListBox As ListBox)
    Set watchingListBox = aListBox
End Sub