Attribute VB_Name = "CreatVBAModule"
'REQUIRES: MS Access <version> Object Library

Dim dbpath            As String              'Path of the database to add the module to
Dim moduleName        As String              'Name of the new module to be created in the database
Dim strCode           As String              'String for the code to be added to new module
Dim defaultModuleName As String              'Default module name chosen by Access
Dim ObjAccess         As Access.Application  'Access Application Object

'Set database path
dbpath = "C:\PATH\To\MY\database.accdb"

'Define Code String with line breaks (vbCrLf)
strCode = "Sub foo()" & vbCrLf & _
          "  Debug.Print ""FOO""" & vbCrLf & _
          "End Sub"

'(English 'Module1', German 'Modul1', ...)
defaultModuleName = "Module1"

'Set new module name
moduleName = "MyModule"

'Initialize Access Application Object
Set ObjAccess = New Access.Application

'Open Database and add blank default module, rename afterwards
ObjAccess.OpenCurrentDatabase dbpath, True
ObjAccess.DoCmd.RunCommand acCmdNewObjectModule
ObjAccess.DoCmd.Save acModule, defaultModuleName
ObjAccess.DoCmd.Rename moduleName, acModule, defaultModuleName
ObjAccess.DoCmd.Save acModule, moduleName

'Loop Modules in Database
For i = 0 To ObjAccess.Modules.Count - 1
    'Find module name
    If ObjAccess.Modules(i).Name = moduleName Then
         'add code to module
         ObjAccess.Modules(i).AddFromString strCode
         Exit For
    End If
Next

'Save and Close
ObjAccess.DoCmd.Save acModule, moduleName
ObjAccess.CloseCurrentDatabase
ObjAccess.Quit

