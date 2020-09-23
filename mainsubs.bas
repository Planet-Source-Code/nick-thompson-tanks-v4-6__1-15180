Attribute VB_Name = "mainsubs"
Public Sub aicontrol()
 For a = 1 To 8
  If pdir(a) > 0 And pstate(a) = 1 Then
    If pxpos(pt(a)) < pxpos(a) And pypos(pt(a)) < pypos(a) Then
      If pdir(a) <= 32 And pdir(a) >= 14 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) < 14 Or pdir(a) > 33 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    ElseIf pxpos(pt(a)) < pxpos(a) And pypos(pt(a)) > pypos(a) Then
      If pdir(a) <= 24 And pdir(a) >= 5 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) < 5 Or pdir(a) > 25 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    ElseIf pxpos(pt(a)) > pxpos(a) And pypos(pt(a)) < pypos(a) Then
       If pdir(a) <= 5 Or pdir(a) >= 25 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) > 6 And pdir(a) < 25 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    ElseIf pxpos(pt(a)) > pxpos(a) And pypos(pt(a)) > pypos(a) Then
       If pdir(a) < 15 Or pdir(a) >= 34 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) > 15 And pdir(a) < 34 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    Else
      pfire(a) = 0
      righton(a) = 0
      lefton(a) = 0
    End If
    upon(a) = 1
  End If
Next a
End Sub

Public Sub changetarget()
  '**** Ends if all but one tank has been destroyed
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) >= 1 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) >= 1 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) >= 1 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) >= 1 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) >= 1 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) >= 1 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) >= 1 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) >= 1 Then Exit Sub
  

For a = 1 To 8 Step 1
  If pstate(a) = 1 And pdir(a) > 0 Then
doh:
    Randomize Timer
    ran = Rnd
     b = a + 1
     If b = 9 Then b = 1
    For ran2 = 0 To 0.875 Step 0.125
      If ran >= ran2 And ran < ran2 + 0.125 Then
        pt(a) = b
       End If
        b = b + 1
        If b = 9 Then b = 1
   
    Next ran2
  If pt(a) = a Then GoTo doh:
  If pdir(pt(a)) = 0 Then GoTo doh
    
  End If
Next a


End Sub
Public Sub movetanks()
  '**** Speeds up or slows down tanks ****
  For a = 1 To 8
    If upon(a) = 1 And downon(a) = 0 Then
      If ps(a) < 101 Then ps(a) = ps(a) + 1
    ElseIf downon(a) = 1 And upon(a) = 0 Then
      If ps(a) > -51 Then ps(a) = ps(a) - 1
    End If
  Next a
 
  For a = 1 To 8
    If pdir(a) = 1 Then
      pypos(a) = pypos(a) - ps(a) / 20 - (pbs(a) / 20)
    End If
    If pdir(a) = 2 Then
      pypos(a) = pypos(a) - ps(a) / 22 - (pbs(a) / 22)
      pxpos(a) = pxpos(a) + ps(a) / 100 + (pbs(a) / 100)
    End If
    If pdir(a) = 3 Then
      pypos(a) = pypos(a) - ps(a) / 25 - (pbs(a) / 25)
      pxpos(a) = pxpos(a) + ps(a) / 60 + (pbs(a) / 60)
    End If
    If pdir(a) = 4 Then
      pypos(a) = pypos(a) - ps(a) / 27 - (pbs(a) / 27)
      pxpos(a) = pxpos(a) + ps(a) / 50 + (pbs(a) / 50)
    End If
    If pdir(a) = 5 Then
      pypos(a) = pypos(a) - ps(a) / 29 - (pbs(a) / 29)
      pxpos(a) = pxpos(a) + ps(a) / 35 + (pbs(a) / 35)
    End If
    If pdir(a) = 6 Then
      pypos(a) = pypos(a) - ps(a) / 35 - (pbs(a) / 35)
      pxpos(a) = pxpos(a) + ps(a) / 29 + (pbs(a) / 29)
    End If
    If pdir(a) = 7 Then
      pypos(a) = pypos(a) - ps(a) / 50 - (pbs(a) / 50)
      pxpos(a) = pxpos(a) + ps(a) / 27 + (pbs(a) / 27)
    End If
    If pdir(a) = 8 Then
      pypos(a) = pypos(a) - ps(a) / 60 - (pbs(a) / 60)
      pxpos(a) = pxpos(a) + ps(a) / 25 + (pbs(a) / 25)
    End If
    If pdir(a) = 9 Then
      pypos(a) = pypos(a) - ps(a) / 100 - (pbs(a) / 100)
      pxpos(a) = pxpos(a) + ps(a) / 22 + (pbs(a) / 22)
    End If
    If pdir(a) = 10 Then
      pxpos(a) = pxpos(a) + ps(a) / 20 + (pbs(a) / 20)
    End If
    If pdir(a) = 11 Then
      pypos(a) = pypos(a) + ps(a) / 100 + (pbs(a) / 100)
      pxpos(a) = pxpos(a) + ps(a) / 22 + (pbs(a) / 22)
    End If
    If pdir(a) = 12 Then
      pypos(a) = pypos(a) + ps(a) / 60 + (pbs(a) / 60)
      pxpos(a) = pxpos(a) + ps(a) / 25 + (pbs(a) / 25)
    End If
    If pdir(a) = 13 Then
      pypos(a) = pypos(a) + ps(a) / 50 + (pbs(a) / 50)
      pxpos(a) = pxpos(a) + ps(a) / 27 + (pbs(a) / 27)
    End If
    If pdir(a) = 14 Then
      pypos(a) = pypos(a) + ps(a) / 35 + (pbs(a) / 35)
      pxpos(a) = pxpos(a) + ps(a) / 29 + (pbs(a) / 29)
    End If
    If pdir(a) = 15 Then
      pypos(a) = pypos(a) + ps(a) / 29 + (pbs(a) / 29)
      pxpos(a) = pxpos(a) + ps(a) / 35 + (pbs(a) / 35)
    End If
    If pdir(a) = 16 Then
      pypos(a) = pypos(a) + ps(a) / 27 + (pbs(a) / 27)
      pxpos(a) = pxpos(a) + ps(a) / 50 + (pbs(a) / 50)
    End If
    If pdir(a) = 17 Then
      pypos(a) = pypos(a) + ps(a) / 25 + (pbs(a) / 25)
      pxpos(a) = pxpos(a) + ps(a) / 60 + (pbs(a) / 60)
    End If
    If pdir(a) = 18 Then
      pypos(a) = pypos(a) + ps(a) / 22 + (pbs(a) / 22)
      pxpos(a) = pxpos(a) + ps(a) / 100 + (pbs(a) / 100)
    End If
    If pdir(a) = 19 Then
      pypos(a) = pypos(a) + ps(a) / 20 + (pbs(a) / 20)
    End If
      If pdir(a) = 20 Then
      pypos(a) = pypos(a) + ps(a) / 22 + (pbs(a) / 22)
      pxpos(a) = pxpos(a) - ps(a) / 100 - (pbs(a) / 100)
    End If
    If pdir(a) = 21 Then
      pypos(a) = pypos(a) + ps(a) / 25 + (pbs(a) / 25)
      pxpos(a) = pxpos(a) - ps(a) / 60 - (pbs(a) / 60)
    End If
    If pdir(a) = 22 Then
      pypos(a) = pypos(a) + ps(a) / 27 + (pbs(a) / 27)
      pxpos(a) = pxpos(a) - ps(a) / 50 - (pbs(a) / 50)
    End If
    If pdir(a) = 23 Then
      pypos(a) = pypos(a) + ps(a) / 29 + (pbs(a) / 29)
      pxpos(a) = pxpos(a) - ps(a) / 35 - (pbs(a) / 35)
    End If
    If pdir(a) = 24 Then
      pypos(a) = pypos(a) + ps(a) / 35 + (pbs(a) / 35)
      pxpos(a) = pxpos(a) - ps(a) / 29 - (pbs(a) / 29)
    End If
    If pdir(a) = 25 Then
      pypos(a) = pypos(a) + ps(a) / 50 + (pbs(a) / 50)
      pxpos(a) = pxpos(a) - ps(a) / 27 - (pbs(a) / 27)
    End If
    If pdir(a) = 26 Then
      pypos(a) = pypos(a) + ps(a) / 60 + (pbs(a) / 60)
      pxpos(a) = pxpos(a) - ps(a) / 25 - (pbs(a) / 25)
    End If
    If pdir(a) = 27 Then
      pypos(a) = pypos(a) + ps(a) / 100 + (pbs(a) / 100)
      pxpos(a) = pxpos(a) - ps(a) / 22 - (pbs(a) / 22)
    End If
    If pdir(a) = 28 Then
      pxpos(a) = pxpos(a) - ps(a) / 20 - (pbs(a) / 20)
    End If
    If pdir(a) = 29 Then
      pypos(a) = pypos(a) - ps(a) / 100 - (pbs(a) / 100)
      pxpos(a) = pxpos(a) - ps(a) / 22 - (pbs(a) / 22)
    End If
    If pdir(a) = 30 Then
      pypos(a) = pypos(a) - ps(a) / 60 - (pbs(a) / 60)
      pxpos(a) = pxpos(a) - ps(a) / 25 - (pbs(a) / 25)
    End If
    If pdir(a) = 31 Then
      pypos(a) = pypos(a) - ps(a) / 50 - (pbs(a) / 50)
      pxpos(a) = pxpos(a) - ps(a) / 27 - (pbs(a) / 27)
    End If
    If pdir(a) = 32 Then
      pypos(a) = pypos(a) - ps(a) / 35 - (pbs(a) / 35)
      pxpos(a) = pxpos(a) - ps(a) / 29 - (pbs(a) / 29)
    End If
    If pdir(a) = 33 Then
      pypos(a) = pypos(a) - ps(a) / 29 - (pbs(a) / 29)
      pxpos(a) = pxpos(a) - ps(a) / 35 - (pbs(a) / 35)
    End If
    If pdir(a) = 34 Then
      pypos(a) = pypos(a) - ps(a) / 27 - (pbs(a) / 27)
      pxpos(a) = pxpos(a) - ps(a) / 50 - (pbs(a) / 50)
    End If
    If pdir(a) = 35 Then
      pypos(a) = pypos(a) - ps(a) / 25 - (pbs(a) / 25)
      pxpos(a) = pxpos(a) - ps(a) / 60 - (pbs(a) / 60)
    End If
    If pdir(a) = 36 Then
      pypos(a) = pypos(a) - ps(a) / 22 - (pbs(a) / 22)
      pxpos(a) = pxpos(a) - ps(a) / 100 - (pbs(a) / 100)
    End If
  Next a
  'slow done after bounce
  For a = 1 To 8
    If pbs(a) > 0 Then pbs(a) = pbs(a) - 1
    If pbs(a) < 0 Then pbs(a) = pbs(a) + 1
  Next a


End Sub
Public Sub moveshells(s As Long)
    If sdir(s) = 1 Then
      sypos(s) = sypos(s) - ss(s) / 20
    End If
    If sdir(s) = 2 Then
      sypos(s) = sypos(s) - ss(s) / 22
      sxpos(s) = sxpos(s) + ss(s) / 100
    End If
    If sdir(s) = 3 Then
      sypos(s) = sypos(s) - ss(s) / 25
      sxpos(s) = sxpos(s) + ss(s) / 60
    End If
    If sdir(s) = 4 Then
      sypos(s) = sypos(s) - ss(s) / 27
      sxpos(s) = sxpos(s) + ss(s) / 50
    End If
    If sdir(s) = 5 Then
      sypos(s) = sypos(s) - ss(s) / 29
      sxpos(s) = sxpos(s) + ss(s) / 35
    End If
    If sdir(s) = 6 Then
      sypos(s) = sypos(s) - ss(s) / 35
      sxpos(s) = sxpos(s) + ss(s) / 29
    End If
    If sdir(s) = 7 Then
      sypos(s) = sypos(s) - ss(s) / 50
      sxpos(s) = sxpos(s) + ss(s) / 27
    End If
    If sdir(s) = 8 Then
      sypos(s) = sypos(s) - ss(s) / 60
      sxpos(s) = sxpos(s) + ss(s) / 25
    End If
    If sdir(s) = 9 Then
      sypos(s) = sypos(s) - ss(s) / 100
      sxpos(s) = sxpos(s) + ss(s) / 22
    End If
    If sdir(s) = 10 Then
      sxpos(s) = sxpos(s) + ss(s) / 20
    End If
    If sdir(s) = 11 Then
      sypos(s) = sypos(s) + ss(s) / 100
      sxpos(s) = sxpos(s) + ss(s) / 22
    End If
    If sdir(s) = 12 Then
      sypos(s) = sypos(s) + ss(s) / 60
      sxpos(s) = sxpos(s) + ss(s) / 25
    End If
    If sdir(s) = 13 Then
      sypos(s) = sypos(s) + ss(s) / 50
      sxpos(s) = sxpos(s) + ss(s) / 27
    End If
    If sdir(s) = 14 Then
      sypos(s) = sypos(s) + ss(s) / 35
      sxpos(s) = sxpos(s) + ss(s) / 29
    End If
    If sdir(s) = 15 Then
      sypos(s) = sypos(s) + ss(s) / 29
      sxpos(s) = sxpos(s) + ss(s) / 35
    End If
    If sdir(s) = 16 Then
      sypos(s) = sypos(s) + ss(s) / 27
      sxpos(s) = sxpos(s) + ss(s) / 50
    End If
    If sdir(s) = 17 Then
      sypos(s) = sypos(s) + ss(s) / 25
      sxpos(s) = sxpos(s) + ss(s) / 60
    End If
    If sdir(s) = 18 Then
      sypos(s) = sypos(s) + ss(s) / 22
      sxpos(s) = sxpos(s) + ss(s) / 100
    End If
    If sdir(s) = 19 Then
      sypos(s) = sypos(s) + ss(s) / 20
    End If
      If sdir(s) = 20 Then
      sypos(s) = sypos(s) + ss(s) / 22
      sxpos(s) = sxpos(s) - ss(s) / 100
    End If
    If sdir(s) = 21 Then
      sypos(s) = sypos(s) + ss(s) / 25
      sxpos(s) = sxpos(s) - ss(s) / 60
    End If
    If sdir(s) = 22 Then
      sypos(s) = sypos(s) + ss(s) / 27
      sxpos(s) = sxpos(s) - ss(s) / 50
    End If
    If sdir(s) = 23 Then
      sypos(s) = sypos(s) + ss(s) / 29
      sxpos(s) = sxpos(s) - ss(s) / 35
    End If
    If sdir(s) = 24 Then
      sypos(s) = sypos(s) + ss(s) / 35
      sxpos(s) = sxpos(s) - ss(s) / 29
    End If
    If sdir(s) = 25 Then
      sypos(s) = sypos(s) + ss(s) / 50
      sxpos(s) = sxpos(s) - ss(s) / 27
    End If
    If sdir(s) = 26 Then
      sypos(s) = sypos(s) + ss(s) / 60
      sxpos(s) = sxpos(s) - ss(s) / 25
    End If
    If sdir(s) = 27 Then
      sypos(s) = sypos(s) + ss(s) / 100
      sxpos(s) = sxpos(s) - ss(s) / 22
    End If
    If sdir(s) = 28 Then
      sxpos(s) = sxpos(s) - ss(s) / 20
    End If
    If sdir(s) = 29 Then
      sypos(s) = sypos(s) - ss(s) / 100
      sxpos(s) = sxpos(s) - ss(s) / 22
    End If
    If sdir(s) = 30 Then
      sypos(s) = sypos(s) - ss(s) / 60
      sxpos(s) = sxpos(s) - ss(s) / 25
    End If
    If sdir(s) = 31 Then
      sypos(s) = sypos(s) - ss(s) / 50
      sxpos(s) = sxpos(s) - ss(s) / 27
    End If
    If sdir(s) = 32 Then
      sypos(s) = sypos(s) - ss(s) / 35
      sxpos(s) = sxpos(s) - ss(s) / 29
    End If
    If sdir(s) = 33 Then
      sypos(s) = sypos(s) - ss(s) / 29
      sxpos(s) = sxpos(s) - ss(s) / 35
    End If
    If sdir(s) = 34 Then
      sypos(s) = sypos(s) - ss(s) / 27
      sxpos(s) = sxpos(s) - ss(s) / 50
    End If
    If sdir(s) = 35 Then
      sypos(s) = sypos(s) - ss(s) / 25
      sxpos(s) = sxpos(s) - ss(s) / 60
    End If
    If sdir(s) = 36 Then
      sypos(s) = sypos(s) - ss(s) / 22
      sxpos(s) = sxpos(s) - ss(s) / 100
    End If
 
End Sub

Public Sub reloadtanks()
  For i = 0 To 7
    If preloaded(i + 1) = 0 Then
      doreload(i) = doreload(i) + 1
    End If
  Next i
  For i = 0 To 7
    If doreload(i) >= reloadspeed Then
      doreload(i) = 0
      preloaded(i + 1) = 1
    End If
  Next i
End Sub

Public Sub turntanks()
  For i = 1 To 8
    If lefton(i) = 1 And righton(i) = 0 And pdir(i) > 0 Then
      pdir(i) = pdir(i) - 1
      If pdir(i) = 0 Then pdir(i) = 36
    ElseIf righton(i) = 1 And lefton(i) = 0 And pdir(i) > 0 Then
      pdir(i) = pdir(i) + 1
      If pdir(i) = 37 Then pdir(i) = 1
    End If
  Next i
End Sub

Public Sub checkforwin()
  '**** Ends if all but one tank has been destroyed
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "It was a Draw"
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) >= 1 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "Player 1 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 0, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) = 0 And pdir(2) >= 1 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "Player 2 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 58, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) >= 1 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "Player 3 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 116, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) >= 1 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "Player 4 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 174, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) >= 1 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "Player 5 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 232, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) >= 1 And pdir(7) = 0 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "Player 6 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 290, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) >= 1 And pdir(8) = 0 Then
    frmmain.lblwin.Caption = "Player 7 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 348, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) >= 1 Then
    frmmain.lblwin.Caption = "Player 8 Wins"
    success = BitBlt(frmmain.picwinner.hDC, 0, 0, frmmain.picwinner.Width / 15, frmmain.picwinner.Height / 15, frmmain.pictanks.hDC, 0, 406, SRCCOPY)
    frmmain.framewin.Visible = True
    frmmain.picwinner.SetFocus
    leavenow = 2
  End If
End Sub

Public Sub collide()
  '**** Detects collisions with edges of screen ****
  For a = 1 To 8
    If pypos(a) < 29 Or pxpos(a) < 29 Or pypos(a) > 600 - 29 Or pxpos(a) > 800 - 29 Then
       If pypos(a) < 29 Then pypos(a) = 29
       If pxpos(a) < 29 Then pxpos(a) = 29
       If pypos(a) > 600 - 29 Then pypos(a) = 600 - 29
       If pxpos(a) > 800 - 29 Then pxpos(a) = 800 - 29
       pbs(a) = -ps(a)
       ps(a) = 0
    End If
  Next a
      
      
  '**** Detect collisions with other tanks ****
  For b = 1 To 8
    For a = 1 To 8
      If a <> b And pdir(a) > 0 And pdir(b) > 0 Then
        If pxpos(b) > pxpos(a) - 40 And pxpos(b) < pxpos(a) + 40 And pypos(b) > pypos(a) - 40 And pypos(b) < pypos(a) + 40 Then
      '    pbdir(b) = pdir(b)
          pbs(b) = -ps(b)
          ps(b) = 0
          pypos(b) = ptypos(b)
          pxpos(b) = ptxpos(b)
        Else
          ptxpos(b) = pxpos(b)
          ptypos(b) = pypos(b)
        End If
      End If
    Next a
  Next b
End Sub

Public Sub regshells()
  '**** Find spare shell entry and registers it ****
  For b = 1 To 8
    If pfire(b) = 1 And preloaded(b) = 1 Then
      For a = 0 To ns
        If sdir(a) = 0 Then
          sdir(a) = pdir(b)
          sxpos(a) = pxpos(b)
          sypos(a) = pypos(b)
          ss(a) = ps(b) + 180
          sown(a) = b
          preloaded(b) = 0
          Exit For
        End If
      Next a
    End If
  Next b
    '**** deallocates shells off screen,
  '     Also detects hits on tanks
      For a = 0 To ns
        If sdir(a) > 0 Then
          Call moveshells(a) 'moves shells
          For b = 1 To 8
            If b <> sown(a) And pdir(b) > 0 Then
              If sxpos(a) > pxpos(b) - 22 And sxpos(a) < pxpos(b) + 22 And sypos(a) > pypos(b) - 22 And sypos(a) < pypos(b) + 22 Then
                sdir(a) = 0
                ph(b) = ph(b) - 1
                If ph(b) = 0 Then
                  pdir(b) = 0
                  ps(b) = 0
                End If
              End If
            End If
          Next b
      'deallocate if off screen
      If sxpos(a) < 0 Or sxpos(a) > 800 Or sypos(a) < 0 Or sypos(a) > 600 Then sdir(a) = 0
    End If
  Next a

End Sub

