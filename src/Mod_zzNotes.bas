'Notes on construction of Matts Version Control System
'
'Form options:
'   Set TortoiseSVN/Git program directory
'   Choose between SVN or Git
'   Export DB and commit to Git/SVN



'Program flow:
'   Must select what to export &/ commit / export data from
'       Tables:
'           __TABLE_LIST__ - yes / yes / yes
'           System - no / no / no - The information from these tables might have to be retained in the stub.
'                                   Or it may be generated automatically as indexes etc when other objects are imported.
'           Hidden - ? / maybe / maybe
'           ODBC - no / no / no
'           Linked - no / no / no
'           All others - yes / yes / maybe
'       Table Data Options:
'           Seed data (should be present in fresh version of app)
'           Test data (should not be present in fresh version of app, but should be in test version)
'       Queries, Forms, Reports, Modules, Macros should all go.
'       Must check how to get Relationships

            