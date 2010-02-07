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

            
'TODO:
'Hook up svn/git radio buttons
'   DONE    Get rid of current debug msgboxes
'Provide more infomation about what's to be/has been exported.
'   DONE    Produce VBScript file that will rebuilt the database.
'   DONE    Export database into folder named for db (to allow multiple dbs in the same place)
'   NO      Decide whether to recurse (ie do workpads/BEs/etc as well)
'           These are not editable within the DB, therefore should not be exported
'Allow initialisation of Git Repo?
'   NO      Make Git also commit the stub db?
'           Stub DB removed entirely
'   NO      Add button to create Stub DB and remove it from the Export procedures?
'           Stub DB removed entirely
'   DONE    Include function to create the build script
'   DONE    Fix exporting tables to the wrong place
'   DONE    Consider what to do when a no-longer existant module is (not) exported
'   DONE    Modify export so Git only modified entities are updated
'Add Fetch/Pull and Push buttons to interface
'   DONE    Save and restore relationships
'Implement Git2Svn functionality
'   DONE    Implement notify/select on failed attempt to re-link a table
'   DONE    CheckAndBuildFolder on click of export not form opening
'   DONE    Implement get/set Options and Startup Properties
'On completion of build, open newly built db and open MVCS form?
'   DONE    Add "Build new Database" button
'   DONE    Delete __TABLE_LIST__ on completion of build? Or leave TABLE_LIST in addin DB permanently?
'   DONE    Saving relationships, options and startup properties may allow us to eliminate the stub database entirely
'   DONE    Import properties on MAKE - do it FIRST so you can delete the table afterwards.
'   DONE    Reconsider placement of __TABLE_LIST__ again! (It's needed for re-linking tables)
'Make Git Commit MAKE file, build script