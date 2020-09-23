VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   2610
   ClientTop       =   2115
   ClientWidth     =   7605
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Sparks.frx":0000
   MousePointer    =   2  'Cross
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   15
      ScaleHeight     =   4515
      ScaleWidth      =   3405
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   3405
      Begin VB.CheckBox Check2 
         Caption         =   "Collect"
         Height          =   315
         Left            =   105
         TabIndex        =   21
         Top             =   4215
         Width           =   825
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cls"
         Height          =   315
         Left            =   135
         TabIndex        =   20
         Top             =   3885
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "?"
         Height          =   330
         Left            =   2925
         TabIndex        =   19
         Top             =   345
         Width           =   345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Sparks.frx":0442
         Left            =   1575
         List            =   "Sparks.frx":0455
         TabIndex        =   14
         Text            =   "Choose One"
         Top             =   945
         Width           =   1305
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   585
         Left            =   1905
         TabIndex        =   12
         Top             =   3900
         Width           =   1395
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1575
         TabIndex        =   10
         Text            =   "120"
         Top             =   360
         Width           =   1260
      End
      Begin VB.Frame Frame2 
         Caption         =   "ParticleStyle"
         Height          =   2355
         Left            =   75
         TabIndex        =   5
         Top             =   1485
         Width           =   3180
         Begin VB.OptionButton Option7 
            Caption         =   "Burst"
            Height          =   285
            Left            =   105
            TabIndex        =   28
            Top             =   1980
            Width           =   1230
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1770
            TabIndex        =   26
            Text            =   "1"
            Top             =   1290
            Width           =   1065
         End
         Begin VB.CheckBox Check4 
            Caption         =   "All Together"
            Height          =   255
            Left            =   1755
            TabIndex        =   25
            Top             =   1935
            Value           =   1  'Checked
            Width           =   1260
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Shower"
            Height          =   285
            Left            =   105
            TabIndex        =   24
            Top             =   1700
            Width           =   1230
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Glitter"
            Height          =   285
            Left            =   105
            TabIndex        =   22
            Top             =   1417
            Width           =   1230
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            Height          =   315
            Left            =   1860
            ScaleHeight     =   255
            ScaleWidth      =   840
            TabIndex        =   17
            Top             =   630
            Width           =   900
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Colour"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   18
               Top             =   15
               Width           =   840
            End
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Sparks.frx":048C
            Left            =   1785
            List            =   "Sparks.frx":0499
            TabIndex        =   16
            Text            =   "Group 1"
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Zig Zag"
            Height          =   285
            Left            =   105
            TabIndex        =   9
            Top             =   1134
            Width           =   1230
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Spew Right"
            Height          =   285
            Left            =   105
            TabIndex        =   8
            Top             =   851
            Width           =   1230
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Spew Left"
            Height          =   285
            Left            =   105
            TabIndex        =   7
            Top             =   568
            Width           =   1230
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Squiggle"
            Height          =   285
            Left            =   105
            TabIndex        =   6
            Top             =   285
            Width           =   1230
         End
         Begin VB.Label Label3 
            Caption         =   "Spew Amount"
            Height          =   225
            Left            =   1800
            TabIndex        =   27
            Top             =   1080
            Width           =   1065
         End
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Text            =   "5"
         ToolTipText     =   "Number of bounces"
         Top             =   780
         Width           =   930
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bounce"
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bouncing"
         Height          =   1395
         Left            =   45
         TabIndex        =   4
         Top             =   75
         Width           =   1410
         Begin VB.CheckBox Check3 
            Caption         =   "Side B"
            Height          =   210
            Left            =   90
            TabIndex        =   23
            Top             =   1080
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Bounces"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   13
            Top             =   510
            Width           =   675
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Placment Option"
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   705
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Particles 1 - 520"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   11
         Top             =   105
         Width           =   1290
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Height          =   150
      Left            =   -45
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -60
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ************
'    SPARKS
'  ************
'
' This is one of my cooler projects. I've been
' Programming since I was 9 in qbasic. I'm now 14 and in vb5.
' Its fairly simple to use. just run it, and move your mouse
' around the screen.The particles are divided into 3 groups
'each with it's own settings. When your ready to do a bit
' of customizing, just click that tiny grey thing
' at the top of the screen. here you can fiddle with the
' particles to no end. Double-click exits if you must! :)
' Note: Less particles = faster peformance    Default = 120
Public Ro As Integer
Public X1, Y1, E, T, NX As Integer, NY As Integer, X, Y, Vel, R As Boolean, Col As Boolean
Private ParticleX(0 To 520) As Long, St(0 To 520) As Boolean, SV(0 To 520)
Private ParticleY(0 To 520) As Long
Private Trail1X(0 To 520) As Long
Private Trail1Y(0 To 520) As Long, SVC(0 To 520) As Integer, SS As Integer
Private Trail2X(0 To 520) As Long, CT As Long, Bst As Integer, BG As Integer
Private Trail2Y(0 To 520) As Long, Bat As Integer
Private Trail3X(0 To 520) As Long
Private Trail3Y(0 To 520) As Long
Private Trail4X(0 To 520) As Long
Private Trail4Y(0 To 520) As Long
Private Trail5X(0 To 520) As Long
Private Trail5Y(0 To 520) As Long
Private Trail6X(0 To 520) As Long
Private Trail6Y(0 To 520) As Long
Private Trail7X(0 To 520) As Long
Private Trail7Y(0 To 520) As Long
Private Trail8X(0 To 520) As Long
Private Trail8Y(0 To 520) As Long
Private Trail9X(0 To 520) As Long
Private Trail9Y(0 To 520) As Long
Private Trail10X(0 To 520) As Long
Private Trail10Y(0 To 520) As Long
Private Bounces(0 To 520) As Long
Private V(0 To 520) As Long, SB As Boolean, Dx, Dy
Private Theta
Private Const Inc = 3.14159265 / 180
Private EMax, Part(1 To 3) As Integer, B As Boolean, NoB, P As Integer, Grp As Integer, Lgrp, C(1 To 3) As Integer, lc, Sy As Boolean, Sx As Boolean

Private Sub Combo2_Click()
Let Lgrp = Combo2.ListIndex + 1
Select Case Part(Lgrp)
Case 1
Option1.Value = True
Case 2
Option2.Value = True
Case 3
Option3.Value = True
Case 4
Option4.Value = True
Case 5
Option5.Value = True
Case 6
Option6.Value = True
Case 7
Option7.Value = True
End Select
Select Case C(Lgrp)
Case 1
Label2.ForeColor = RGB(250, 250, 0)
Case 2
Label2.ForeColor = RGB(0, 0, 250)
Case 3
Label2.ForeColor = RGB(250, 0, 250)
End Select
End Sub

Private Sub Command1_Click()
Let Picture1.Visible = True
Lgrp = 1
End Sub

Private Sub Command2_Click()
If Check4.Value = 1 Then
If Option1.Value = True Then Let Part(1) = 1
If Option2.Value = True Then Let Part(1) = 2
If Option3.Value = True Then Let Part(1) = 3
If Option4.Value = True Then Let Part(1) = 4
If Option5.Value = True Then Let Part(1) = 5
If Option6.Value = True Then Let Part(1) = 6
If Option7.Value = True Then Let Part(1) = 7
Let SS = Text3
Let Part(2) = Part(1)
Let Part(3) = Part(2)
If T = 0 Or T = 6 Then Let T = 1
C(1) = T
C(2) = T
C(3) = T
End If
If Int(Text1.Text) > 30 Then MsgBox "Invalid Bounce Number" Else Let NoB = Int(Text1.Text)
If Int(Text2.Text) > 520 Then MsgBox "Invalid Particle Number" Else Let EMax = Int(Text2.Text)

Let B = Check1.Value
Let Col = Check2.Value
Let SB = Check3.Value
If Combo1.ListIndex <> -1 Then Let P = Combo1.ListIndex
Let Picture1.Visible = False
End Sub

Private Sub Command3_Click()
MsgBox "Programed by michael pote, 14, (Who says South Africans Can't Program!?)" & Chr(13) & "Send any querys to mikepote@mailcity.com"
End Sub

Private Sub Command5_Click()
Form1.Cls
End Sub

Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Load() 'where evrything happens
Show
'setup default settings
EMax = 120
SVC(0) = -1
Part(1) = 1
Part(2) = 1
Part(3) = 1
C(1) = 1
C(2) = 1
C(3) = 1
SS = 1
NoB = 5
P = 0
B = True
SB = True
Bst = 10
Do
DoEvents
'loop through all particles
Let E = E + 1
If E >= EMax + 1 Then
Let E = 0
End If
'setup groups
If E <= EMax / 3 Then Let Grp = 1
If E <= EMax / 3 * 2 And E >= EMax / 3 Then Let Grp = 2
If E <= EMax And E >= EMax / 3 * 2 Then Let Grp = 3
If E <= EMax / 90 Then Let BG = 1
'If Part(Grp) = 7 Then
'For S = 2 To 90
'If E <= EMax / 90 * S And E >= EMax / 90 * S - 1 Then Let BG = S
'Next
'End If
'set velocity
Let V(E) = V(E) + 1
'update trails
Let Trail10X(E) = Trail9X(E)
Let Trail10Y(E) = Trail9Y(E)
Let Trail9X(E) = Trail8X(E)
Let Trail9Y(E) = Trail8Y(E)
Let Trail8X(E) = Trail7X(E)
Let Trail8Y(E) = Trail7Y(E)
Let Trail7X(E) = Trail6X(E)
Let Trail7Y(E) = Trail6Y(E)
Let Trail6X(E) = Trail5X(E)
Let Trail6Y(E) = Trail5Y(E)
Let Trail5X(E) = Trail4X(E)
Let Trail5Y(E) = Trail4Y(E)
Let Trail4X(E) = Trail3X(E)
Let Trail4Y(E) = Trail3Y(E)
Let Trail3X(E) = Trail2X(E)
Let Trail3Y(E) = Trail2Y(E)
Let Trail2X(E) = Trail1X(E)
Let Trail2Y(E) = Trail1Y(E)
Let Trail1X(E) = ParticleX(E)
Let Trail1Y(E) = ParticleY(E)
'move particle sideways
If Part(Grp) <> SVC(E) Then
Select Case Part(Grp)
Case 2
Let SV(E) = SS
Case 3
Let SV(E) = -SS
Case 6
If St(E) = False Then Let SV(E) = (Int(Rnd * 50) - 24) / 2: Let St(E) = True
Case 7
Theta = Inc * (360 / EMax) * (EMax - E) + 10
Dx = X1 + Cos(Theta) * 50
Dy = Y1 + Sin(Theta) * 10
Let Dx = X1 - Dx
Let Dy = Y1 - Dy
Let SV(E) = Dy
Let V(E) = Dx
End Select
End If
Let SVC(E) = Part(Grp)
If Part(Grp) = 1 Then Let SV(E) = Int(Rnd * 5) - 2.4
If Part(Grp) = 4 Then Let SV(E) = Sin(V(E)) * 5
If Part(Grp) = 5 Then Let SV(E) = Sin(V(E)) * V(E)

If (ParticleX(E) >= ScaleWidth Or ParticleX(E) <= 0) And SB = True Then Let SV(E) = -SV(E)
Let ParticleX(E) = ParticleX(E) + SV(E)
'move particle down
Let ParticleY(E) = ParticleY(E) + Int((V(E) + Rnd * 10) / 5)
'bounce particle
If ParticleY(E) >= ScaleHeight Then
If B = False Then GoTo D
Let Bounces(E) = Bounces(E) + 1
Let V(E) = -((NoB - Bounces(E)) * (V(E) / 5))
If Bounces(E) = NoB Then
D:
Bounces(E) = 0
Let St(E) = False
Let SV(E) = 0
Let SVC(E) = -1
CurrentX = X1
CurrentY = Y1
Select Case P
Case 0
CurrentX = X1
CurrentY = Y1
ParticleX(E) = X1: ParticleY(E) = Y1
Case 1
Let ParticleY(E) = Rnd * ScaleHeight
Let ParticleX(E) = Rnd * ScaleWidth
Case 2
Let ParticleY(E) = 0
Let ParticleX(E) = Rnd * ScaleWidth
Case 3

If R = False Then
Vel = 0
Let NX = Rnd * ScaleWidth
Let NY = Rnd * ScaleHeight
R = True
If Y1 < NY Then
Let Sy = False
ElseIf Y1 > NY Then
Let Sy = True
End If
If X1 < NX Then
Let Sx = False
ElseIf X1 > NX Then
Let Sx = True
End If
End If
Vel = Vel + 1
If Y1 < NY Then
Let Y1 = Y1 + (Vel / 100)
ElseIf Y1 > NY Then
Let Y1 = Y1 - (Vel / 100)
End If
If X1 < NX Then
Let X1 = X1 + (Vel / 100)
ElseIf X1 > NX Then
Let X1 = X1 - (Vel / 100)
End If
If X1 <= NX + 10 And X1 >= NX - 10 And Y1 <= NY + 10 And Y1 >= NY - 10 Then Let R = False
ParticleX(E) = X1: ParticleY(E) = Y1
Case 4
Ro = Ro + 5
If Ro >= 361 Then Let Ro = 0
Theta = Inc * Ro
Dx = (ScaleWidth / 2) + Cos(Theta) * 100
Dy = (ScaleHeight / 2) + Sin(Theta) * 100
ParticleX(E) = Dx: ParticleY(E) = Dy
End Select
If Col = True Then PSet (Trail10X(E), Trail10Y(E)), RGB(250, 250, 0)
Let V(E) = 0
If Col = True Then GoTo ByPass
End If
End If
'draw particle
DoEvents
Select Case C(Grp)
Case 1
PSet (ParticleX(E), ParticleY(E)), RGB(250, 250, 0)
PSet (Trail1X(E), Trail1Y(E)), RGB(250, 200, 0)
PSet (Trail2X(E), Trail2Y(E)), RGB(190, 160, 0)
PSet (Trail3X(E), Trail3Y(E)), RGB(190, 150, 0)
PSet (Trail4X(E), Trail4Y(E)), RGB(180, 140, 0)
PSet (Trail5X(E), Trail5Y(E)), RGB(180, 130, 0)
PSet (Trail6X(E), Trail6Y(E)), RGB(170, 120, 0)
PSet (Trail7X(E), Trail7Y(E)), RGB(170, 100, 0)
PSet (Trail8X(E), Trail8Y(E)), RGB(160, 90, 0)
PSet (Trail9X(E), Trail9Y(E)), RGB(160, 80, 0)
PSet (Trail10X(E), Trail10Y(E)), RGB(0, 0, 0)
PSet (Trail10X(E), Trail10Y(E)), RGB(0, 0, 0)
Case 2
PSet (ParticleX(E), ParticleY(E)), RGB(100, 100, 250)
PSet (Trail1X(E), Trail1Y(E)), RGB(90, 90, 230)
PSet (Trail2X(E), Trail2Y(E)), RGB(80, 80, 210)
PSet (Trail3X(E), Trail3Y(E)), RGB(70, 70, 180)
PSet (Trail4X(E), Trail4Y(E)), RGB(50, 50, 150)
PSet (Trail5X(E), Trail5Y(E)), RGB(30, 30, 130)
PSet (Trail6X(E), Trail6Y(E)), RGB(20, 20, 90)
PSet (Trail7X(E), Trail7Y(E)), RGB(10, 10, 70)
PSet (Trail8X(E), Trail8Y(E)), RGB(0, 0, 40)
PSet (Trail9X(E), Trail9Y(E)), RGB(0, 0, 20)
PSet (Trail10X(E), Trail10Y(E)), RGB(0, 0, 0)
Case 3
CT = (V(E) * V(E))
If CT <= 0 Then Let CT = 0
On Error Resume Next
PSet (ParticleX(E), ParticleY(E)), RGB(CT, CT, CT)
PSet (Trail1X(E), Trail1Y(E)), RGB(CT - 12, CT - 12, CT - 12)
PSet (Trail2X(E), Trail2Y(E)), RGB(CT - 24, CT - 24, CT - 24)
PSet (Trail3X(E), Trail3Y(E)), RGB(CT - 40, CT - 40, CT - 40)
PSet (Trail4X(E), Trail4Y(E)), RGB(CT - 80, CT - 80, CT - 80)
PSet (Trail5X(E), Trail5Y(E)), RGB(CT - 100, CT - 100, CT - 100)
PSet (Trail6X(E), Trail6Y(E)), RGB(CT - 140, CT - 140, CT - 140)
PSet (Trail7X(E), Trail7Y(E)), RGB(CT - 200, CT - 200, CT - 200)
PSet (Trail8X(E), Trail8Y(E)), RGB(CT - 220, CT - 220, CT - 220)
PSet (Trail9X(E), Trail9Y(E)), RGB(CT - 255, CT - 255, CT - 255)
PSet (Trail10X(E), Trail10Y(E)), RGB(0, 0, 0)
Case 4
PSet (ParticleX(E), ParticleY(E)), RGB(250, 0, 0)
PSet (Trail1X(E), Trail1Y(E)), RGB(250, 0, 0)
PSet (Trail2X(E), Trail2Y(E)), RGB(226, 0, 0)
PSet (Trail3X(E), Trail3Y(E)), RGB(200, 0, 0)
PSet (Trail4X(E), Trail4Y(E)), RGB(176, 0, 0)
PSet (Trail5X(E), Trail5Y(E)), RGB(150, 0, 0)
PSet (Trail6X(E), Trail6Y(E)), RGB(124, 0, 0)
PSet (Trail7X(E), Trail7Y(E)), RGB(100, 0, 0)
PSet (Trail8X(E), Trail8Y(E)), RGB(74, 0, 0)
PSet (Trail9X(E), Trail9Y(E)), RGB(50, 0, 0)
PSet (Trail10X(E), Trail10Y(E)), RGB(0, 0, 0)
Case 5
PSet (ParticleX(E), ParticleY(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail2X(E), Trail2Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail3X(E), Trail3Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail4X(E), Trail4Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail5X(E), Trail5Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail6X(E), Trail6Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail7X(E), Trail7Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail8X(E), Trail8Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail9X(E), Trail9Y(E)), RGB(ParticleX(E), ParticleY(E), 800 - ParticleX(E))
PSet (Trail10X(E), Trail10Y(E)), RGB(0, 0, 0)

End Select
ByPass:
Loop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set drawwing coords
If P = 3 Then Exit Sub
Let X1 = X
Let Y1 = Y
End Sub

Private Sub Label2_Click()
T = T + 1
If T = 6 Then Let T = 1
Select Case T
Case 1
Label2.ForeColor = RGB(250, 250, 0)
Let Label2.Caption = "Yellow"
Case 2
Label2.ForeColor = RGB(0, 0, 250)
Let Label2.Caption = "Blue"
Case 3
Label2.ForeColor = RGB(150, 150, 150)
Let Label2.Caption = "Velocity"
Case 4
Label2.ForeColor = RGB(250, 0, 0)
Let Label2.Caption = "Red"
Case 5
Label2.ForeColor = RGB(250, 0, 250)
Let Label2.Caption = "Vaired"
End Select
Let C(Lgrp) = T
End Sub

Private Sub Option1_Click()
Let Part(Lgrp) = 1
End Sub

Private Sub Option2_Click()
Let Part(Lgrp) = 2

End Sub

Private Sub Option3_Click()
Let Part(Lgrp) = 3

End Sub

Private Sub Option4_Click()
Let Part(Lgrp) = 4

End Sub

Private Sub Option5_Click()
Let Part(Lgrp) = 5
End Sub

Private Sub Option6_Click()
Let Part(Lgrp) = 6

End Sub

Private Sub Option7_Click()
Let Part(Lgrp) = 7

End Sub

