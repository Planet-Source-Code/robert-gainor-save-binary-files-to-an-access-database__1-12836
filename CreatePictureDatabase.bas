Attribute VB_Name = "CreatePictureDatabase"
Option Explicit
Dim Newdb As Database


Public Function CreatedNewDB(ByVal sDestDBPath As String, ByVal sDestDBPassword As String) As Boolean

'''''''''''''''''''''''''''''''''''''''
'Database Tables, Fields & Indexes
'Copied from: D:\Visual Basic Projects\Picture Database and Example projects\Example of saving a byte array to an access database\Picture.mdb
'On: 11/16/00 7:24:18 AM
'Copied via rcSmithDBCopy ver:1.6
'REQUIRES:  Reference to MS DAO in VB project
'NOTE NOTE NOTE:  Code does *not* check Validity of Destination Path!!
'''''''''''''''''''''''''''''''''''''''

CreatedNewDB = False

On Error GoTo Err_Handler

If sDestDBPassword <> "" Then
     Set Newdb = Workspaces(0).CreateDatabase(sDestDBPath, dbLangGeneral & ";pwd=" & sDestDBPassword)
Else
     Set Newdb = Workspaces(0).CreateDatabase(sDestDBPath, dbLangGeneral)
End If

'Now call the functions for each table

Dim b As Boolean

b = CreatedNewTBLTable1
If b = False Then
     CreatedNewDB = False
     Newdb.Close
     Set Newdb = Nothing
     Exit Function
End If

Newdb.Close
Set Newdb = Nothing
CreatedNewDB = True

Exit Function

Err_Handler:
     If Err.Number <> 0 Then
          'Alert & Close Objects, could be altered to Raise the error
          MsgBox "Error Creating Copy Database." & vbCr & Err.Number & vbCr & Err.Description
               CreatedNewDB = False
               Newdb.Close

               Set Newdb = Nothing

               Exit Function
     End If
End Function

Private Function CreatedNewTBLTable1() As Boolean

'''''''''''''''''''''''''''''''''''''''
'Database Table:Table1
'Copied from: D:\Visual Basic Projects\Picture Database and Example projects\Example of saving a byte array to an access database\Picture.mdb
'On: 11/16/00 7:24:19 AM
'Copied via rcSmithDBCopy ver:1.6
'REQUIRES:  Reference to MS DAO in VB project
'NOTE NOTE NOTE:  Code does *not* check Validity of Destination Path!!
'''''''''''''''''''''''''''''''''''''''

Dim TempTDef As TableDef
Dim TempField As Field
Dim TempIdx As Index

CreatedNewTBLTable1 = False

On Error GoTo Err_Handler

Set TempTDef = Newdb.CreateTableDef("Table1")
     Set TempField = TempTDef.CreateField("Picture", 11)
          TempField.Attributes = 2
          TempField.Required = False
          TempField.OrdinalPosition = 1
     TempTDef.Fields.Append TempField
     TempTDef.Fields.Refresh

     Set TempField = TempTDef.CreateField("Type", 10)
          TempField.Attributes = 2
          TempField.Required = False
          TempField.OrdinalPosition = 2
          TempField.Size = 50
          TempField.AllowZeroLength = False
     TempTDef.Fields.Append TempField
     TempTDef.Fields.Refresh

     Set TempField = TempTDef.CreateField("Name", 10)
          TempField.Attributes = 2
          TempField.Required = False
          TempField.OrdinalPosition = 3
          TempField.Size = 50
          TempField.AllowZeroLength = False
     TempTDef.Fields.Append TempField
     TempTDef.Fields.Refresh

Newdb.TableDefs.Append TempTDef
Newdb.TableDefs.Refresh

'Done, Close the objects
     Set TempTDef = Nothing
     Set TempField = Nothing
     Set TempIdx = Nothing

CreatedNewTBLTable1 = True

Exit Function

Err_Handler:
     If Err.Number <> 0 Then
          'Alert & Close Objects, could be altered to Raise the error
               MsgBox "Error Creating Database Table: Table1" & vbCr & Err.Number & vbCr & Err.Description
     Set TempTDef = Nothing
     Set TempField = Nothing
     Set TempIdx = Nothing

     CreatedNewTBLTable1 = False
     Exit Function
     End If
End Function

