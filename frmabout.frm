VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Tanks"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Version 4.6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "TheVBGod@Hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Programmed by Nick Thompson in Visual Basic 5. Please send any comments or upgrade suggestions to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
frmabout.Visible = False
frmmenu.Visible = True
Unload frmabout

End Sub

Private Sub Form_Load()
frmabout.Left = (Screen.Width - frmabout.Width) / 2
frmabout.Top = (Screen.Height - frmabout.Height) / 2

End Sub
