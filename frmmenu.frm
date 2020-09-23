VERSION 5.00
Begin VB.Form frmmenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tanks V4.6"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ControlBox      =   0   'False
   Icon            =   "frmmenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdoptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   23
      Top             =   4800
      Width           =   1575
   End
   Begin VB.PictureBox picp6 
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   2280
      Width           =   855
   End
   Begin VB.PictureBox picp8 
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox picp5 
      Height          =   855
      Left            =   1920
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.PictureBox picp7 
      Height          =   855
      Left            =   1920
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox picp4 
      Height          =   855
      Left            =   6960
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox picp3 
      Height          =   855
      Left            =   720
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox picp2 
      Height          =   855
      Left            =   6960
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.PictureBox picp1 
      Height          =   855
      Left            =   720
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   5640
      Width           =   4455
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Layout of players mimics that of game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   22
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   21
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   20
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "V4.6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6300
      TabIndex        =   13
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Game Configuration"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2400
      TabIndex        =   4
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tanks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdabout_Click()
frmmenu.Visible = False
frmabout.Visible = True
End Sub

Private Sub cmdexit_Click()
End

End Sub

Private Sub cmdoptions_Click()
frmoptions.Visible = True
frmmenu.Visible = False
End Sub

Private Sub cmdstart_Click()
leavenow = 0
If slowdown < 10 Then slowdown = 1010
  For a = 1 To 8
    ph(a) = armour
     pfire(a) = 0
    upon(a) = 0
    downon(a) = 0
    ps(a) = 0
    pbs(a) = 0
 Next a
If pstate(1) = 1 Then pdir(1) = 1
If pstate(1) = 2 Then pdir(1) = 1
If pstate(2) = 0 Then pdir(2) = 0
If pstate(2) = 1 Then pdir(2) = 1
If pstate(2) = 2 Then pdir(2) = 1
If pstate(3) = 0 Then pdir(3) = 0
If pstate(3) = 1 Then pdir(3) = 19
If pstate(4) = 0 Then pdir(4) = 0
If pstate(4) = 1 Then pdir(4) = 19
If pstate(5) = 0 Then pdir(5) = 0
If pstate(5) = 1 Then pdir(5) = 1
If pstate(6) = 0 Then pdir(6) = 0
If pstate(6) = 1 Then pdir(6) = 1
If pstate(7) = 0 Then pdir(7) = 0
If pstate(7) = 1 Then pdir(7) = 19
If pstate(8) = 0 Then pdir(8) = 0
If pstate(8) = 1 Then pdir(8) = 19
frmmenu.Visible = False
frmmain.Visible = True


End Sub

Private Sub Form_Load()
  Open App.Path & "\info.dat" For Input As #1
    Input #1, gspeed
    Input #1, reloadspeed
    Input #1, dback
    Input #1, armour
  Close #1
'  reloadspeed = 300
  
  ns = 1500
  frmmenu.Left = (Screen.Width - frmmenu.Width) / 2
  frmmenu.Top = (Screen.Height - frmmenu.Height) / 2
  pstate(1) = 2
  pstate(2) = 2
  pstate(3) = 1
  pstate(4) = 1
  pstate(5) = 1
  pstate(6) = 1
  pstate(7) = 1
  pstate(8) = 1
  picp1.Picture = LoadPicture(App.Path & "\pics\p1h.bmp")
  picp2.Picture = LoadPicture(App.Path & "\pics\p2h.bmp")
  picp3.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
  picp4.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
  picp5.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
  picp6.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
  picp7.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
  picp8.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
End Sub






Private Sub picp1_Click()
If pstate(1) = 2 Then
  pstate(1) = 1
  picp1.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
Else
  picp1.Picture = LoadPicture(App.Path & "\pics\p1h.bmp")
  pstate(1) = 2
End If
End Sub

Private Sub picp2_Click()
If pstate(2) = 1 Then
  pstate(2) = 2
  picp2.Picture = LoadPicture(App.Path & "\pics\p2h.bmp")
Else
  pstate(2) = 1
  picp2.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
End If
End Sub

Private Sub picp3_Click()
If pstate(3) = 0 Then
  pstate(3) = 1
  picp3.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
Else
  pstate(3) = 0
  picp3.Picture = LoadPicture(App.Path & "\pics\none.bmp")
End If

End Sub

Private Sub picp4_Click()
If pstate(4) = 0 Then
  pstate(4) = 1
  picp4.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
Else
  pstate(4) = 0
  picp4.Picture = LoadPicture(App.Path & "\pics\none.bmp")
End If

End Sub

Private Sub picp5_Click()
If pstate(5) = 0 Then
  pstate(5) = 1
  picp5.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
Else
  pstate(5) = 0
  picp5.Picture = LoadPicture(App.Path & "\pics\none.bmp")
End If

End Sub

Private Sub picp6_Click()
If pstate(6) = 0 Then
  pstate(6) = 1
  picp6.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
Else
  pstate(6) = 0
  picp6.Picture = LoadPicture(App.Path & "\pics\none.bmp")
End If

End Sub

Private Sub picp7_Click()
If pstate(7) = 0 Then
  pstate(7) = 1
  picp7.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
Else
  pstate(7) = 0
  picp7.Picture = LoadPicture(App.Path & "\pics\none.bmp")
End If

End Sub

Private Sub picp8_Click()
If pstate(8) = 0 Then
  pstate(8) = 1
  picp8.Picture = LoadPicture(App.Path & "\pics\comp.bmp")
Else
  pstate(8) = 0
  picp8.Picture = LoadPicture(App.Path & "\pics\none.bmp")
End If

End Sub
