Option Compare Database
Option Explicit

Public Const REFERENCES_TABLENAME As String = "__REFERENCES__BE"
Public Const REFERENCES_FILENAME As String = "__REFERENCES__.xml"
'Those are underscores

Public Function ListReferences() As Integer
'Ref_Name, Ref_GUID, Ref_Major, Ref_Minor
On Error GoTo ErrProc
Dim Ref As Reference
Dim refCount As Integer
Dim Refs As DAO.Recordset

CodeDb.Execute "DELETE * FROM " & REFERENCES_TABLENAME

Set Refs = CodeDb.OpenRecordset(REFERENCES_TABLENAME, dbOpenDynaset)
    If Not Refs.EOF Then
        Refs.MoveFirst
    End If
    For Each Ref In Application.References
        Refs.AddNew
            Refs("Ref_Name") = Ref.Name
            Refs("Ref_GUID") = Ref.Guid
            Refs("Ref_Major") = Ref.Major
            Refs("Ref_Minor") = Ref.Minor
        Refs.Update
        refCount = refCount + 1
    Next Ref
Refs.Close

Exit Function
ErrProc:
DispErrMsgGSb Error$, "document the library references"
ListReferences = -1

End Function