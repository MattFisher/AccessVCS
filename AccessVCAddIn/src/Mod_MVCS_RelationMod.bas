Option Compare Database
Option Explicit

Public Const RELATIONS_TABLENAME As String = "__RELATIONS__BE"
Public Const RELATIONS_FILENAME As String = "__RELATIONS__.xml"
'Those are underscores

Public Function ListRelationships() As Integer
                                   
'RelName, RelTable, RelForeignTable, RelFieldName, RelForeignFieldName, (All <= 64 char strings)
'RelAttributes, (long)
'RelPartialReplica (Boolean)

On Error GoTo ErrProc
Dim rel As Relation
Dim relCount As Integer
Dim db As DAO.Database
Dim Rels As DAO.Recordset

Log "Taking stock of database relationships"
Set db = Access.CurrentDb

CodeDb.Execute "DELETE * FROM " & RELATIONS_TABLENAME

Set Rels = CodeDb.OpenRecordset(RELATIONS_TABLENAME, dbOpenDynaset)
    If Not Rels.EOF Then
        Rels.MoveFirst
    End If
    For Each rel In db.Relations
        Rels.AddNew
            Rels("RelName") = rel.Name
            Rels("RelTable") = rel.Table
            Rels("RelForeignTable") = rel.ForeignTable
            Rels("RelFieldName") = rel.Fields(0).Name
            Rels("RelForeignFieldName") = rel.Fields(0).ForeignName
            Rels("RelAttributes") = rel.Attributes
            Rels("RelPartialReplica") = rel.PartialReplica
        Rels.Update
        relCount = relCount + 1
    Next rel
Rels.Close

Exit Function
ErrProc:
DispErrMsgGSb Error$, "document the Table Relations"
ListRelationships = -1

End Function

'Relation Attributes
'1    &H1    dbRelationUnique        - The relationship is one-to-one.
'2    &H2    dbRelationDontEnforce   - The relationship isn't enforced
'                                      (no referential integrity).
'4    &H4    dbRelationInherited     - The relationship exists in a non-current
'                                      database that contains the two linked tables.
'256  &H100  dbRelationUpdateCascade - Updates will cascade.
'4096 &H1000 dbRelationDeleteCascade - Deletions will cascade.
'16777216    dbRelationLeft          - Microsoft Access only. In Design view, display
'&H1000000                             a LEFT JOIN as the default join type.
'33554432    dbRelationRight         - Microsoft Access only. In Design view, display
'&H2000000                             a RIGHT JOIN as the default join type.