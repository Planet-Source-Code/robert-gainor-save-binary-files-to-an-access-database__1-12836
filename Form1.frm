VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpendb 
         Caption         =   "Open Database"
      End
      Begin VB.Menu mnuNewdb 
         Caption         =   "Create New Database"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Picture"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save To Database"
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "Extract Picture"
      End
      Begin VB.Menu mnuViewPicture 
         Caption         =   "View picture in Database"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------
'- Name: Robert Gainor
'- Email: robertgainor@hotmail.com

'- Date/Time: 11/16/00 8:58:46 AM
'----------------------------------------
'- Notes:
'This application shows how open a binary file put it into a  byte array and then
'save that  byte array into an Access database
'using the OLE object field type
'also how to extract the data from the field and convert it back into a file
'this application uses graphic files as an example but with minor modifications
'to the code any file type can be saved into a database OLE object field
'I kind of threw this together in a couple of hours, and I know there are
'better ways of doing some of the basic stuff within the program and a few bugs might crop up
'and alot of error handling is missing
'but this is only meant to be an example not a finished Product.
'some of the minor functions were thrown together on the spur of the moment so
'forgive the lack of comments
'This version was created with Visual Basic 6 SP4 using a reference to DAO 3.6 but I'm sure with
'minor modifications you can use Visual Basic 5
'The CreatePictureDatabase mod was made with "Jet Database Copier"
'from smithvoice.com, thanks guys it saves me alot of time.
'if you find this usefull let me know
'also if you know of a better way of doing this let me know
'
'----------------------------------------

Option Explicit
Dim dbPictures As Database
Public rsPictures As Recordset
Public strCurrentPicture As String
Dim TheBytes() As Byte
Dim strfile As String


Private Sub Form_Resize()
'resize the picturebox to the form
Picture1.Width = Me.ScaleWidth
Picture1.Height = Me.ScaleHeight
'set picturebox position in the form
Picture1.Left = Me.ScaleLeft
Picture1.Top = Me.ScaleTop

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'close any open recordset and database
rsPictures.Close
Set rsPictures = Nothing
dbPictures.Close
Set dbPictures = Nothing

End Sub

Private Sub mnuExtract_Click()
On Error Resume Next

Dim newfile As String
Dim strBytes As String

'Show a list of pictures in the database
'This can also be a list of any type of file that you have stored in the
'Database

Form2.Show vbModal

With CommonDialog1
     .FileName = ""
     .DialogTitle = "Extract to"
     .CancelError = True
     .Filter = "All (*." & Getfileext(newfile) & ")|*." & Getfileext(newfile)
     .FileName = strCurrentPicture
     .ShowSave
     If Err.Number > 0 Then
          Exit Sub
     End If
     newfile = .FileName
End With
ReDim TheBytes(0)
rsPictures.MoveFirst
'find the picture in the database
Do Until rsPictures.EOF = True
     If rsPictures("name") & "." & rsPictures("type") = strCurrentPicture Then
     'we found the picture
          'set the byte array to the size needed to hold the picture data
          ReDim TheBytes(rsPictures("Picture").FieldSize - 1)
          'put the picture data from the database in the array
          TheBytes() = rsPictures("Picture").GetChunk(0, rsPictures("Picture").FieldSize)
          'open the newfile and put the byte array in it
          Open newfile For Binary Access Write As #1
               Put #1, , strBytes
          Close #1
          
          'We are done with the loop so exit
          Exit Do
     End If
     rsPictures.MoveNext
Loop


End Sub

Private Sub mnuNewdb_Click()
On Error Resume Next

'Select a place for the new database on the computer
With CommonDialog1
     .FileName = ""
     .DialogTitle = "Create Database"
     .CancelError = True
     .Filter = "All mdb files (*.mdb)|*.mdb"
     .ShowSave
     'if the cancel button is clicked exit the sub
     If Err.Number > 0 Then
          GoTo ErrorHandle
     End If
     strfile = .FileName
End With
'close open recordset and set to nothing
'for the new recordset

rsPictures.Close
Set rsPictures = Nothing
'Close open database and set to nothing
'for the new database
dbPictures.Close
Set dbPictures = Nothing
If Err.Number > 0 Then
     'clear any errors that have occured in the above statements
     Err.Clear
End If
'create the new database
Dim bNewdb As Boolean

bNewdb = CreatedNewDB(strfile, "")
'check to see if it was created
If bNewdb = True Then
     'if it was created open it
     Set dbPictures = Workspaces(0).OpenDatabase(strfile)
     'open the recordset
     Set rsPictures = dbPictures.OpenRecordset("Table1", dbOpenDynaset)
End If
Me.Caption = "Database - " & strfile
Exit Sub
ErrorHandle:



End Sub

Private Sub mnuopen_Click()
On Error GoTo ErrorHandle
'select a picture file to open

With CommonDialog1
     .FileName = ""
     .DialogTitle = "Open File"
     .CancelError = True
     .Filter = "All Bitmaps or Jpeg files (*.bmp)(*.jpg)|*.bmp;*.jpg"
     .ShowOpen
     strfile = .FileName
End With
'load the file into the picture box
Picture1.Picture = LoadPicture(strfile)

ErrorHandle:
End Sub

Private Sub mnuOpendb_Click()
On Error Resume Next


'select the database that you want to open
With CommonDialog1
     .FileName = ""
     .DialogTitle = "Open File"
     .CancelError = True
     .Filter = "All mdb files (*.mdb)|*.mdb"
     .ShowOpen
     'if the cancel button is pressed then exit the sub
     If Err.Number > 0 Then
          GoTo ErrorHandle:
     End If
     strfile = .FileName
End With
'need to close any open database and recordset
rsPictures.Close
Set rsPictures = Nothing
dbPictures.Close
Set dbPictures = Nothing
'open up the selected database and recordset
Set dbPictures = Workspaces(0).OpenDatabase(strfile)
Set rsPictures = dbPictures.OpenRecordset("table1", dbOpenDynaset)

Me.Caption = "Database - " & strfile

Exit Sub
ErrorHandle:

End Sub

Private Sub mnuSave_Click()
On Error Resume Next

'open the file to get the data
Open strfile For Binary Access Read As #1
     'set up the byte array to hold the file information
     ReDim TheBytes(FileLen(strfile) - 1)
     
     'put the file into the array
     Get #1, , TheBytes()
'close the file
Close #1
'move to the last record in the pics recordset
rsPictures.MoveLast
'add the picture
rsPictures.AddNew
'put the byte array in the pic field
rsPictures("Picture").AppendChunk TheBytes
'put the file type in the type field
rsPictures("type") = Getfileext(strfile)
'put the file name without the extension in the name field
rsPictures("name") = GetFileNameWOExt(strfile)
'update the database
rsPictures.Update
'move to the first record
rsPictures.MoveFirst
'clear picture out
Picture1.Picture = LoadPicture
'clear the file string
strfile = ""

End Sub

Private Sub mnuViewPicture_Click()
Dim tempfile As String

'show the list of pictures
Form2.Show vbModal

'get the picture from the recordset
rsPictures.MoveFirst
Do Until rsPictures.EOF = True
     If rsPictures("name") & "." & rsPictures("type") = strCurrentPicture Then
          'we found the picture put it in the byte array
          
          'reset the byte array
          ReDim TheBytes(rsPictures("picture").FieldSize - 1)
          'put the picture in the array
          TheBytes() = rsPictures("picture").GetChunk(0, rsPictures("picture").FieldSize)
     
          'set up the tempfile
          tempfile = App.Path & "\" & strCurrentPicture

          'open the tempfile
          Open tempfile For Binary Access Write As #1
          'put the byte array in the tempfile
          Put #1, , TheBytes()
          'close the file
          Close #1
          'load the tempfile into the picturebox
          Picture1.Picture = LoadPicture(tempfile)
          'delete the tempfile
          Kill tempfile
          'exit the loop because we're done with it
          Exit Do
     End If
     'move to the next record
     rsPictures.MoveNext
Loop

End Sub


Private Function GetFileName(strFullPath As String) As String
Dim iposition As Integer

iposition = InStrRev(strFullPath, "\")
GetFileName = Right(strFullPath, Len(strFullPath) - iposition)



ErrorHandle:
End Function

Private Function Getfileext(strpath As String) As String
Dim iposition As Integer
iposition = InStrRev(strpath, ".")
Getfileext = Right(strpath, Len(strpath) - iposition)

End Function

Private Function GetFileNameWOExt(strpath As String) As String
Dim iposition As Integer
'get the file name from the fullpath
strpath = GetFileName(strpath)
'get the position of the . in the file name
iposition = InStrRev(strpath, ".")
'get everything to the left of the . in the file name
GetFileNameWOExt = Left(strpath, iposition - 1)


End Function
