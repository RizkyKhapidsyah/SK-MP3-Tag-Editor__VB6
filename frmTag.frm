VERSION 5.00
Begin VB.Form frmTag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MPEG File Info Box + ID3 Tag Editor"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3450
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   2340
      Width           =   2835
   End
   Begin VB.Frame frmButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   3255
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1080
         TabIndex        =   24
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove ID3"
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame frmID3 
      Caption         =   "Tag info"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Comment 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox Title 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox Artist 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Album 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Year 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox Genre 
         Height          =   315
         ItemData        =   "frmTag.frx":0000
         Left            =   1920
         List            =   "frmTag.frx":027F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label labGenre 
         AutoSize        =   -1  'True
         Caption         =   "Genre"
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   1365
         Width           =   435
      End
      Begin VB.Label labYear 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   345
         TabIndex        =   11
         Top             =   1365
         Width           =   450
      End
      Begin VB.Label labComent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   1725
         Width           =   660
      End
      Begin VB.Label labTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   195
         Left            =   375
         TabIndex        =   9
         Top             =   300
         Width           =   420
      End
      Begin VB.Label labAlbum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Album"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label labArtist 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Artist"
         Height          =   195
         Left            =   330
         TabIndex        =   7
         Top             =   660
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add ID3"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame frmNoTag 
      Caption         =   "Tag info"
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   3975
      Begin VB.Label labNoTag 
         AutoSize        =   -1  'True
         Caption         =   "This MP3 doesn't contain ID3 Tag"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   2430
      End
   End
   Begin VB.Frame frmMPEG 
      Caption         =   "MPEG info"
      Height          =   2415
      Left            =   4200
      TabIndex        =   13
      Top             =   120
      Width           =   2175
      Begin VB.Label labEmphasis 
         AutoSize        =   -1  'True
         Caption         =   "Emphasis:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1815
         Width           =   720
      End
      Begin VB.Label labOriginal 
         AutoSize        =   -1  'True
         Caption         =   "Original:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label labCopy 
         AutoSize        =   -1  'True
         Caption         =   "Copyrighted:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labBitRate 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   825
         Width           =   45
      End
      Begin VB.Label labLayer 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   615
         Width           =   45
      End
      Begin VB.Label labCRC 
         AutoSize        =   -1  'True
         Caption         =   "CRCs:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1215
         Width           =   450
      End
      Begin VB.Label labFreqChan 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1020
         Width           =   45
      End
      Begin VB.Label labLength 
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   540
      End
      Begin VB.Label labSize 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame frmNoMPEG 
      Caption         =   "MPEG info"
      Height          =   2415
      Left            =   4200
      TabIndex        =   20
      Top             =   120
      Width           =   2175
      Begin VB.Label labNoMPEG 
         AutoSize        =   -1  'True
         Caption         =   "Probably not a MP3 file..."
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Dim MP3Length As Long
Dim MP3File As String
Dim MP3Size As Long
Dim locArtist As String * 30
Dim locTitle As String * 30
Dim locAlbum As String * 30
Dim locYear As String * 4
Dim locComment As String * 30
Dim locGenre As String * 1
Option Explicit

Private Sub cmdAdd_Click()
  Dim emptyStr As String * 124
  
  frmID3.Visible = True
  frmButtons.Visible = True
  MP3Size = FileLen(MP3File)
  Open MP3File For Binary As #1
    Put #1, MP3Size, "Dejvi"
    Put #1, MP3Size, Chr$(0) & "TAG" & emptyStr & Chr$(255)
  Close #1
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  Dim tmpStr As String
  
  Me.Caption = "Removing Tag"
  MP3Size = FileLen(MP3File)
  tmpStr = Space(MP3Size - 128)
  Open MP3File For Binary As #1
    Get #1, 1, tmpStr
  Close #1
  Kill MP3File
  Open MP3File For Binary As #1
    Put #1, 1, tmpStr
  Close #1
  Unload Me
End Sub

Private Sub cmdSave_Click()
  MP3Size = FileLen(MP3File)
  locTitle = Title
  locArtist = Artist
  locAlbum = Album
  locYear = Year
  locComment = Comment
  If Genre.ListIndex = -1 Then
    locGenre = Chr(255)
  Else
    locGenre = Chr(Genre.ItemData(Genre.ListIndex))
  End If
  Open MP3File For Binary As #1
    Put #1, MP3Size + 1 - 128, "TAG" & locTitle & locArtist & locAlbum & locYear & locComment & locGenre
  Close #1
  Unload Me
End Sub

Private Sub Form_Load()
  GetTagInf
  GetMP3Inf
End Sub

Private Sub GetMP3Inf()
  Dim accMP3Info As MP3Info
  
  getMP3Info MP3FileName, accMP3Info
  
  labSize = labSize & " " & accMP3Info.SIZE
  labLength = labLength & " " & accMP3Info.LENGTH
  labLayer = accMP3Info.MPEG & " " & accMP3Info.LAYER
  labBitRate = accMP3Info.BITRATE
  labFreqChan = accMP3Info.FREQ & " " & accMP3Info.CHANNELS
  labCRC = labCRC & accMP3Info.CRC
  labCopy = labCopy & accMP3Info.COPYRIGHT
  labEmphasis = labEmphasis & accMP3Info.EMPHASIS
  labOriginal = labOriginal & accMP3Info.ORIGINAL
End Sub

Private Sub GetTagInf()
  Dim Buf As String * 128
  Dim tmpStr As String
  Dim i As Byte
  
  MP3File = MP3FileName
  'Get the size of mp3 file(in bytes)
  MP3Size = FileLen(MP3File)
  
  'labLength = labLength & mp3Length & " seconds"
  
  'Open the file for binary access in order to get the ID3 Tag
  Open MP3File For Binary As #1
    'Get last 128 bytes of the file. The size of file is reduced by 127 bytes, because
    'the last byte in file is in fact the size of file
    Get #1, MP3Size - 127, Buf
    'Check if the file has a tag
    If Format(Left(Buf, 3), "<") <> "tag" Then
      frmID3.Visible = False
      frmButtons.Visible = False
    Else
      'If it has a tag the separate the info obtained in the buffer string
      Title = Trim(Mid(Buf, 4, 30))
      Artist = Trim(Mid(Buf, 34, 30))
      Album = Trim(Mid(Buf, 64, 30))
      Year = Trim(Mid(Buf, 94, 4))
      Comment = Trim(Mid(Buf, 98, 30))
      For i = 0 To 148
        If Genre.ItemData(i) = Trim(Asc(Mid$(Buf, 128, 1))) Then Exit For
      Next i
      If i < 149 Then
        Genre.ListIndex = i
      End If
    End If
  Close #1
End Sub

