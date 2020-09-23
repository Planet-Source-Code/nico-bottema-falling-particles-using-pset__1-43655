VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ParticleX(1800), ParticleY(1800), oldparticlex(1800), oldparticley(1800), ParticleC(1800), particles

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If particles < 1770 Then
For i = 1 To Int(Rnd * 30)
Xpar = Int(Rnd * 40)
Ypar = Int(Rnd * 40)
If Xpar < 20 Then Xpar = X - Xpar Else Xpar = X + Xpar - 20
If Ypar < 20 Then Ypar = Y - Ypar Else Ypar = Y + Ypar - 20
    Call AddParticle(Xpar, Ypar)
Next i
End If
End If
End Sub

Private Sub AddParticle(Xpos, YPos)
particles = particles + 1
ParticleX(particles) = Xpos
ParticleY(particles) = YPos
ParticleC(particles) = Int(Rnd * 15)


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If particles < 1770 Then
For i = 1 To Int(Rnd * 30)
Xpar = Int(Rnd * 40)
Ypar = Int(Rnd * 40)
If Xpar < 20 Then Xpar = X - Xpar Else Xpar = X + Xpar - 20
If Ypar < 20 Then Ypar = Y - Ypar Else Ypar = Y + Ypar - 20
    Call AddParticle(Xpar, Ypar)
Next i
End If
End If
End Sub

Private Sub Timer1_Timer()
For i = 1 To particles
DoEvents
xmover = Int(Rnd * 2)
ymover = Int(Rnd * 2)
PSet (oldparticlex(i), oldparticley(i)), QBColor(0)
If xmover = 1 Then ParticleX(i) = ParticleX(i) - Int(Rnd * 60) Else ParticleX(i) = ParticleX(i) + Int(Rnd * 60)
If ymover = 1 Then ParticleY(i) = ParticleY(i) - Int(Rnd * 20) Else ParticleY(i) = ParticleY(i) + Int(Rnd * 100)
oldparticlex(i) = ParticleX(i)
oldparticley(i) = ParticleY(i)
PSet (ParticleX(i), ParticleY(i)), QBColor(ParticleC(i))
If ParticleY(i) > 6000 Then Call RemoveParticle(i)
Next i
End Sub

Private Sub RemoveParticle(ParNr)
For par = ParNr To particles - 1
ParticleX(par) = ParticleX(par + 1)
ParticleY(par) = ParticleY(par + 1)
oldparticlex(par) = oldparticlex(par + 1)
oldparticley(par) = oldparticley(par + 1)
ParticleC(par) = ParticleC(par + 1)
Next par
particles = particles - 1

End Sub
