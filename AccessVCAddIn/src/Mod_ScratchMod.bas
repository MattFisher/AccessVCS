Option Compare Database

Sub CreateRelationX()

   Dim dbsNorthwind As Database
   Dim tdfEmployees As TableDef
   Dim tdfNew As TableDef
   Dim idxNew As Index
   Dim relNew As Relation
   Dim idxLoop As Index

   Set dbsNorthwind = OpenDatabase("Northwind.mdb")

   With dbsNorthwind
      ' Add new field to Employees table.
      Set tdfEmployees = .TableDefs!Employees
      tdfEmployees.Fields.Append _
         tdfEmployees.CreateField("DeptID", dbInteger, 2)

      ' Create new Departments table.
      Set tdfNew = .CreateTableDef("Departments")

      With tdfNew
         ' Create and append Field objects to Fields
         ' collection of the new TableDef object.
         .Fields.Append .CreateField("DeptID", dbInteger, 2)
         .Fields.Append .CreateField("DeptName", dbText, 20)

         ' Create Index object for Departments table.
         Set idxNew = .CreateIndex("DeptIDIndex")
         ' Create and append Field object to Fields
         ' collection of the new Index object.
         idxNew.Fields.Append idxNew.CreateField("DeptID")
         ' The index in the primary table must be Unique in
         ' order to be part of a Relation.
         idxNew.Unique = True
         .Indexes.Append idxNew
      End With

      .TableDefs.Append tdfNew

      ' Create EmployeesDepartments Relation object, using
      ' the names of the two tables in the relation.
      Set relNew = .CreateRelation("EmployeesDepartments", _
         tdfNew.Name, tdfEmployees.Name, _
         dbRelationUpdateCascade)

      ' Create Field object for the Fields collection of the
      ' new Relation object. Set the Name and ForeignName
      ' properties based on the fields to be used for the
      ' relation.
      relNew.Fields.Append relNew.CreateField("DeptID")
      relNew.Fields!DeptID.ForeignName = "DeptID"
      .Relations.Append relNew

      ' Print report.
      Debug.Print "Properties of " & relNew.Name & _
         " Relation"
      Debug.Print "  Table = " & relNew.Table
      Debug.Print "  ForeignTable = " & _
         relNew.ForeignTable
      Debug.Print "Fields of " & relNew.Name & " Relation"

      With relNew.Fields!DeptID
         Debug.Print "  " & .Name
         Debug.Print "    Name = " & .Name
         Debug.Print "    ForeignName = " & .ForeignName
      End With

      Debug.Print "Indexes in " & tdfEmployees.Name & _
         " TableDef"
      For Each idxLoop In tdfEmployees.Indexes
         Debug.Print "  " & idxLoop.Name & _
            ", Foreign = " & idxLoop.Foreign
      Next idxLoop

      ' Delete new objects because this is a demonstration.
      .Relations.Delete relNew.Name
      .TableDefs.Delete tdfNew.Name
      tdfEmployees.Fields.Delete "DeptID"
      .Close
   End With

End Sub