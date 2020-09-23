VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pictures in Database"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Load the list box with pictures from the picture recordset
Form1.rsPictures.MoveFirst
Do Until Form1.rsPictures.EOF = True
     List1.AddItem Form1.rsPictures("Name") & "." & Form1.rsPictures("type")
     Form1.rsPictures.MoveNext
Loop
If List1.ListCount = 0 Then
     Unload Me
End If

End Sub

Private Sub Form_Resize()
'size the listbox to the form
List1.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight

End Sub


Private Sub List1_DblClick()
'load the current strpicture in form1 with the picture selected
Form1.strCurrentPicture = List1.Text
'unload the list
List1.Clear
'unload the form
Unload Me
End Sub
