VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbResizeMode 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   6000
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox cmbColourChange 
      Height          =   315
      ItemData        =   "Form1.frx":002C
      Left            =   2640
      List            =   "Form1.frx":003C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblResizeMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resizing Mode:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4440
      TabIndex        =   13
      Top             =   997
      Width           =   1455
   End
   Begin VB.Shape shaBullet2 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   6990
      Width           =   135
   End
   Begin VB.Label lblResize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There are now two methods of resizing."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   6960
      Width           =   2865
   End
   Begin VB.Label lblColour 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The demo project (this one) now supports dynamic colour scheme changing."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   6690
      Width           =   5460
   End
   Begin VB.Shape shaBullet1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   135
   End
   Begin VB.Label lblChange 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Colour Scheme:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   990
      Width           =   2295
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":005A
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0419
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   1680
      Top             =   360
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
   Begin VB.Line lneWhatsNew 
      BorderColor     =   &H00800000&
      X1              =   16
      X2              =   496
      Y1              =   440
      Y2              =   440
   End
   Begin VB.Label lbWhatsNew 
      BackStyle       =   0  'Transparent
      Caption         =   "What's New In This Release?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label lblPara1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0663
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   7455
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Version 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label lblPara4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":087A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   7455
   End
   Begin VB.Label lblVote 
      BackStyle       =   0  'Transparent
      Caption         =   "Please vote and leave comments on this code at Planet Source Code!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6000
      Width           =   6735
   End
   Begin VB.Label lblPara3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0983
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   7455
   End
   Begin VB.Label lblPara2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0A77
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   7455
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Version 3"
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
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2085
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0B2C
      Top             =   480
      Width           =   195
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0D76
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0FC0
      Top             =   0
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   5880
      Picture         =   "Form1.frx":120A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   6240
      Picture         =   "Form1.frx":1954
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "Form1.frx":209E
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   6960
      Picture         =   "Form1.frx":27E8
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5880
      Picture         =   "Form1.frx":2F32
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   6240
      Picture         =   "Form1.frx":367C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "Form1.frx":3DC6
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   6960
      Picture         =   "Form1.frx":4510
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

Private Sub cmbColourChange_Click()
' ****************************************************
' * THANKS TO STEVEN GERHARDT FOR THE FOLLOWING CODE *
' ****************************************************

On Error Resume Next
    For Each ctl In Me.Controls
        If TypeOf ctl Is Image Then
            Debug.Print ctl.Name
            ctl.Picture = LoadPicture(App.Path & "\Images\" & cmbColourChange.Text & "\" & ctl.Name & ".gif")
        End If
    Next
    SetStateBtn Me, Me.WindowState
    For Each ctl In Me.Controls
        ctl.Refresh
    Next
End Sub

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
    If cmbResizeMode.Text = "As cursor drags" Then
        If flgResize Then
            Temp = GetCursorPos(NewCursorPos)
            ResizeForm Me, OldCursorPos, NewCursorPos, Index
            OldCursorPos = NewCursorPos
        End If
    End If
End Sub

Private Sub Resizer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    flgResize = False
    Temp = GetCursorPos(NewCursorPos)
    ResizeForm Me, OldCursorPos, NewCursorPos, Index
End Sub

Private Sub imgTitleMain_DblClick()
    ChangeState Me
End Sub

