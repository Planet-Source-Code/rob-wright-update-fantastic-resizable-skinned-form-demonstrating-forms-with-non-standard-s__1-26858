VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0000
      Top             =   1200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":03BF
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   7320
      Top             =   720
      Width           =   195
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   7
      Left            =   2400
      MousePointer    =   8  'Size NW SE
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   6
      Left            =   2160
      MousePointer    =   6  'Size NE SW
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   5
      Left            =   1920
      MousePointer    =   6  'Size NE SW
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   4
      Left            =   1680
      MousePointer    =   8  'Size NW SE
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   3
      Left            =   720
      MousePointer    =   7  'Size N S
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   2
      Left            =   720
      MousePointer    =   7  'Size N S
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   1
      Left            =   480
      MousePointer    =   9  'Size W E
      Top             =   120
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   0
      Left            =   240
      MousePointer    =   9  'Size W E
      Top             =   120
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1770
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0609
      Top             =   480
      Width           =   195
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0853
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0A9D
      Top             =   0
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   5880
      Picture         =   "Form1.frx":0CE7
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   6240
      Picture         =   "Form1.frx":1431
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "Form1.frx":1B7B
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   6960
      Picture         =   "Form1.frx":22C5
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5880
      Picture         =   "Form1.frx":2A0F
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   6240
      Picture         =   "Form1.frx":3159
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "Form1.frx":38A3
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   6960
      Picture         =   "Form1.frx":3FED
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************
' * WARNING                                               *
' * =======                                               *
' *                                                       *
' * This code relies heavily on the Z-Index (arrangement, *
' * send to back etc.) of the elements that make up the   *
' * skinned window.  Therefore, if you use this in your   *
' * own programs, you might have to fiddle about sending  *
' * things to back, front and the like before it works.   *
' *********************************************************

Dim Temp
Dim flgResize As Boolean
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI

Private Sub Form_Load()
    MakeWindow Me, True
    'AlwaysOnTop Me, True

' Make the Maximize/Restore button have the Maximize image
   imgTitleMaxRestore.Picture = imgTitleMaximize.Picture
End Sub

Private Sub imgTitleClose_Click()
    Unload Me
    End
End Sub

Private Sub imgTitleHelp_Click()
    MsgBox "You can insert code here for loading a help file."
End Sub

Private Sub imgTitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMaxRestore_Click()
    ChangeState Me
End Sub

Private Sub imgTitleMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub imgTitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub lblTitle_DblClick()
    ChangeState Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub Resizer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    flgResize = True
    Temp = GetCursorPos(OldCursorPos)
End Sub

Private Sub Resizer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' COMMENT THE FOLLOWING CODE OUT FOR THE FLICKER-FREE (sort of!) RESIZING
   ' If flgResize Then
   '     Temp = GetCursorPos(NewCursorPos)
   '     ResizeForm Me, OldCursorPos, NewCursorPos, Index
   '     OldCursorPos = NewCursorPos
   ' End If
End Sub

Private Sub Resizer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    flgResize = False
    Temp = GetCursorPos(NewCursorPos)
    ResizeForm Me, OldCursorPos, NewCursorPos, Index
End Sub

Private Sub imgTitleMain_DblClick()
    ChangeState Me
End Sub

