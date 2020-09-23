VERSION 5.00
Begin VB.Form frmoptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Options"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reload Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblrspeed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Game Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdgspeed 
         Caption         =   "Change"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblgspeed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Armour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   3135
      Begin VB.CommandButton Command2 
         Caption         =   "Change"
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
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblarmour 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   3135
      Begin VB.OptionButton optb 
         Caption         =   "ON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optb 
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdback_Click()
frmoptions.Visible = False
frmmenu.Visible = True
End Sub

Private Sub cmdgspeed_Click()
  If gspeed >= 300 Then
    gspeed = 20
  Else
    gspeed = gspeed + 10
  End If
  lblgspeed.Caption = gspeed
End Sub

Private Sub cmdsave_Click()
  dodraw = 1
  Open App.Path & "\info.dat" For Output As #1
    Write #1, gspeed
    Write #1, reloadspeed
    Write #1, dback
    Write #1, armour
  Close #1
  frmoptions.Visible = False
  frmmenu.Visible = True

End Sub

Private Sub Command1_Click()
  If reloadspeed = 1 Then
    reloadspeed = 5
  ElseIf reloadspeed < 50 Then
    reloadspeed = reloadspeed + 5
  ElseIf reloadspeed >= 50 Then
    reloadspeed = 1
  End If
  lblrspeed.Caption = reloadspeed
End Sub

Private Sub Command2_Click()
If armour < 200 Then
  armour = armour + 20
Else
  armour = 20
End If
lblarmour.Caption = armour

End Sub

Private Sub Form_Load()
lblrspeed.Caption = reloadspeed
lblarmour.Caption = armour
lblgspeed.Caption = gspeed
'If gofast = 1 Then
'  optgs(1).Value = True
'Else
'  optgs(0).Value = True
'End If
If dback = 1 Then
  optb(0).Value = True
Else
  optb(1).Value = True
End If
End Sub

Private Sub optb_Click(Index As Integer)
If optb(0).Value = True Then
  dback = 1
Else
  dback = 0
End If

End Sub

Private Sub optgs_Click(Index As Integer)
'If optgs(0).Value = True Then
'  gofast = 0
'Else
'  gofast = 1
'End If
End Sub
