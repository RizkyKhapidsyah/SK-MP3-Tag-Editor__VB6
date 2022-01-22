VERSION 5.00
Begin VB.Form frmChoose 
   Caption         =   "Select MP3 file"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox fileMP3 
      Height          =   4380
      Left            =   3120
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   0
      Width           =   2685
   End
   Begin VB.DirListBox dirMP3 
      Height          =   3915
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   3075
   End
   Begin VB.DriveListBox drvMP3 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dirMP3_Change()
  fileMP3 = dirMP3
End Sub

Private Sub drvMP3_Change()
  dirMP3 = drvMP3
  fileMP3 = dirMP3
End Sub

Private Sub fileMP3_DblClick()
  
  If Len(fileMP3.Path) > 3 Then
    MP3FileName = fileMP3.Path & "\"
  Else
    MP3FileName = fileMP3.Path
  End If
  MP3FileName = MP3FileName & fileMP3.FileName
Clipboard.Clear
Clipboard.SetText fileMP3.FileName
  DoEvents
  frmTag.Show vbModal


 
End Sub
