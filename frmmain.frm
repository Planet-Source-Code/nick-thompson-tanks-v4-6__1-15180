VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tanks V4.6 - Get Ready"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picbackgnd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   615
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer timermain 
      Interval        =   1000
      Left            =   540
      Top             =   60
   End
   Begin VB.Frame framewin 
      Height          =   1335
      Left            =   3120
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.PictureBox picwinner 
         AutoRedraw      =   -1  'True
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdcontinue 
         Caption         =   "Continue"
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
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblwin 
         Alignment       =   2  'Center
         Caption         =   "Player 1 Wins"
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
         TabIndex        =   7
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.PictureBox picshells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picshellsm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer timerfps 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   60
      Top             =   60
   End
   Begin VB.PictureBox picback 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   3420
      ScaleHeight     =   3840
      ScaleWidth      =   6000
      TabIndex        =   2
      Top             =   1380
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox pictanksm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox pictanks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  If leavenow > 0 Then Exit Sub 'leave if this foorm has accidentally been caled eg by a timer - shouldnt happen but just in case
  domove = 1 'how often it does move information - always set to 1
  If dodraw = 0 Then dodraw = 1 'dodraw is how often it draws to screen, if 0 then it too is set to 1
  fps = 0 'current speed of moving
  fpsdraw = 0 'current speed of moving
  slowdown = 0 'slows the entire program down if fpsdraw > 60
  dochangeai = 0 'countdown (up) to AI target change
  doturn = 0 'countdown (up) to processing turns
  docheck = 0 'countdown (up) to checking for a winner
  framewin.Left = (frmmain.Width - framewin.Width) / 2 'centre win frame
  framewin.Top = (frmmain.Height - framewin.Height) / 2
  pxpos(1) = 200 'set start positions
  pypos(1) = 100
  pxpos(2) = 600
  pypos(2) = 101
  pxpos(3) = 201
  pypos(3) = 500
  pxpos(4) = 601
  pypos(4) = 501
  pxpos(5) = 300
  pypos(5) = 100
  pxpos(6) = 500
  pypos(6) = 101
  pxpos(7) = 301
  pypos(7) = 500
  pxpos(8) = 501
  pypos(8) = 501
  picback.Width = 800 * 15 'set width of picback
  picback.Height = 600 * 15 'set height of picback
  pictanks.Picture = LoadPicture(App.Path & "\pics\tankso.bmp") 'load pictures
  pictanksm.Picture = LoadPicture(App.Path & "\pics\tanks.bmp")
  picshells.Picture = LoadPicture(App.Path & "\pics\shell.bmp")
  picshellsm.Picture = LoadPicture(App.Path & "\pics\shellm.bmp")
  If dback = 1 Then picbackgnd.Picture = LoadPicture(App.Path & "\pics\back.bmp") 'only load background if needed
  For i = 0 To ns 'each shells direction = 0
    sdir(i) = 0
  Next i
  For i = 1 To 8
    preloaded(i) = 1 'each tank has reloaded
    ps(i) = 0 'tanks speed = 0
  Next i
End Sub

Private Sub timermain_Timer()
  If framewin.Visible = True Then Exit Sub 'Leave if game has ended
  Call changetarget 'set start targets for AI
  timerfps.Enabled = True 'start the frames per second timer
  Do
    moveopto = moveopto + 0.1
    If domove <= moveopto Then 'when domove = moveopto all move information is done
      fps = fps + 1 'tell fps timer how many fps have been done
      moveopto = 0 'reset moveopto
      doturn = doturn + 1 'many subs are not done every time we throught this area
      docheck = docheck + 1 'thes variables count up and when reaching a certain vale
      dochangeai = dochangeai + 1 'the subs are executed
      Call reloadtanks 'reload is done after a certain time not at exact time intervals like the others
      If dochangeai = 150 Then 'so the AI change
        dochangeai = 0
        Call changetarget
      End If
      If doturn = 2 Then 'turn the tanks
        doturn = 0
        Call turntanks
      End If
      If docheck = 10 Then 'check for winner
        docheck = 0
        Call checkforwin
      End If
      '**** Calls Sub in module to move tanks ****
      Call movetanks 'move tanks, speed up, slow down etc
      Call collide 'detect tanks collisions
      Call regshells 'allocate shells, deallocated them and detect hits
      Call aicontrol 'AI Control
    End If
        
    drawopto = drawopto + 0.1
    If dodraw <= drawopto Then
      For i = 1 To 500 'In my experence computers need between 1 and 200 doevents or the dont process key presses correctly, 500 should (hopefully be safe for all computers)
        DoEvents
      Next i
      For i = 1 To slowdown 'to stop the fps running way to fast and using all system resources it is limited to 60 by slowdown
        DoEvents
      Next i
      fpsdraw = fpsdraw + 1 'tell fps timer the fps
      drawopto = 0 'reset drawopto
      If dback = 1 Then BitBlt picback.hDC, 0, 0, 800, 600, picbackgnd.hDC, 0, 0, SRCCOPY 'draw background if needed
      '**** Draws Tanks on screen ****
      For i = 1 To 8
        If pdir(i) > 0 Then '0 if not in game
          success = BitBlt(picback.hDC, pxpos(i) - 29, pypos(i) - 29, 58, 58, pictanks.hDC, (pdir(i) - 1) * 58, (i - 1) * 58, SRCAND)
          success = BitBlt(picback.hDC, pxpos(i) - 29, pypos(i) - 29, 58, 58, pictanksm.hDC, (pdir(i) - 1) * 58, (i - 1) * 58, SRCPAINT)
        End If
      Next i
      '**** draw shells to screen
      For i = 0 To ns
        If sdir(i) > 0 Then
          success = BitBlt(picback.hDC, sxpos(i) - 2, sypos(i) - 2, 4, 4, picshells.hDC, 0, 0, SRCAND)
          success = BitBlt(picback.hDC, sxpos(i) - 2, sypos(i) - 2, 4, 4, picshellsm.hDC, 0, 0, SRCPAINT)
        End If
      Next i
      'copy all to display
      success = BitBlt(frmmain.hDC, 0, 0, 800, 600, picback.hDC, 0, 0, SRCCOPY)
      If dback = 0 Then picback.Cls 'clear for next frame (not needed if background if drawn
    End If
  Loop Until leavenow > 0 'leavenow = 2 when the winframe is desplayed - so it doesn't close the form
  If leavenow = 1 Then
    frmmain.Visible = False
    frmmenu.Visible = True
    Unload frmmain
  End If
  leavenow = 0
End Sub

Private Sub timerfps_Timer()
  If framewin.Visible = True Then Exit Sub
  checkspeed = checkspeed + 1
  If checkspeed = 2 Then 'this makes it only do it every second - you can change this to your personal preferance
    frmmain.Caption = "Tanks V4.6 " & "- Frames per Second = " & fpsdraw & " - Game Speed = " & fps & " - Target = " & gspeed 'show info to user
    If fps > gspeed - 10 And fpsdraw > 65 Then 'if both are very high the slow down whole program
      slowdown = slowdown + 40
    ElseIf slowdown > 40 And fpsdraw < 50 Then 'if slowdown is on and the drawing fps is low decrease slowdown
      slowdown = slowdown - 40
    End If
    checkspeed = 0
    If fps < (gspeed - 5) Then 'if the moving fps is too low then decrease the speed of the drawing
      dodraw = dodraw + 0.1
    ElseIf fps > (gspeed + 5) Then 'opposite of above
      dodraw = dodraw - 0.1
    End If
    fps = 0 'reset both
    fpsdraw = 0
  End If
End Sub

Private Sub cmdcontinue_Click()
 For i = 1 To 8 'this is on the win frame
    lefton(i) = 0
    righton(i) = 0
    upon(i) = 0
    downon(i) = 0
    ps(i) = 0
    pbs(i) = 0
    pbdir(i) = 0
    pfire(i) = 0
  Next i
  For i = 0 To ns
    sdir(i) = 0
  Next i
  frmmenu.Visible = True
  frmmain.Visible = False
  Unload frmmain
  Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If framewin.Visible = True Then Exit Sub
  If KeyCode = 27 Then
    leavenow = 1
    frmmenu.Visible = True
    frmmain.Visible = False
    Unload frmmain
    Exit Sub
  End If
  If pdir(1) > 0 And pstate(1) = 2 Then
    If KeyCode = 37 Then lefton(1) = 1
    If KeyCode = 38 Then upon(1) = 1
    If KeyCode = 39 Then righton(1) = 1
    If KeyCode = 40 Then downon(1) = 1
    If KeyCode = 13 Then pfire(1) = 1
  End If
  If pdir(2) > 0 And pstate(2) = 2 Then
    If KeyCode = 65 Then lefton(2) = 1
    If KeyCode = 87 Then upon(2) = 1
    If KeyCode = 68 Then righton(2) = 1
    If KeyCode = 83 Then downon(2) = 1
    If KeyCode = 32 Then pfire(2) = 1
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If framewin.Visible = True Then Exit Sub
  If pdir(1) > 0 Then
    If KeyCode = 37 Then lefton(1) = 0
    If KeyCode = 38 Then upon(1) = 0
    If KeyCode = 39 Then righton(1) = 0
    If KeyCode = 40 Then downon(1) = 0
    If KeyCode = 13 Then pfire(1) = 0
  End If
  If pdir(2) > 0 Then
    If KeyCode = 65 Then lefton(2) = 0
    If KeyCode = 87 Then upon(2) = 0
    If KeyCode = 68 Then righton(2) = 0
    If KeyCode = 83 Then downon(2) = 0
    If KeyCode = 32 Then pfire(2) = 0
  End If
End Sub

