VERSION 5.00
Begin VB.Form frmGM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphics Method Demo"
   ClientHeight    =   5880
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuGraphics 
      Caption         =   "&Graphics"
      Begin VB.Menu mnuDP 
         Caption         =   "&Draw Points"
      End
      Begin VB.Menu mnuDL 
         Caption         =   "&Draw Lines"
      End
      Begin VB.Menu mnuDML 
         Caption         =   "&Draw Messy Lines"
      End
      Begin VB.Menu mnus1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
   End
   Begin VB.Menu mnuDB 
      Caption         =   "D&raw Box"
      Begin VB.Menu mnuR 
         Caption         =   "R&ed"
      End
      Begin VB.Menu mnuG 
         Caption         =   "Gree&n"
      End
      Begin VB.Menu mnuB 
         Caption         =   "&Blue"
      End
      Begin VB.Menu mnus2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSS 
         Caption         =   "&Set Style"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DP
Dim DL
Dim DML
Private Sub Form_Load()
DP = 0
DL = 0
DML = 0
End Sub

Private Sub mnuAbout_Click()
MsgBox "Graphics Method Demo" & vbCrLf & "by Lam Ri Hui" & vbCrLf & "(rihui@email.com)", , "Graphic Method Demo"

End Sub

Private Sub mnuB_Click()
Me.FillColor = RGB(0, 0, 255)
Me.Line (100, 100)-Step(8265, 5625), RGB(0, 0, 255), B
End Sub

Private Sub mnuClear_Click()
DP = 0
DL = 0
DML = 0
Me.Cls
End Sub

Private Sub mnuDL_Click()
DL = 1
End Sub

Private Sub mnuDML_Click()
DML = 1
End Sub

Private Sub mnuDP_Click()
DP = 1
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuG_Click()
Me.FillColor = RGB(0, 255, 0)
Me.Line (100, 100)-Step(8265, 5625), RGB(0, 255, 0), B
End Sub

Private Sub mnuR_Click()
Me.FillColor = RGB(255, 0, 0)
Me.Line (100, 100)-Step(8265, 5625), RGB(255, 0, 0), B
End Sub

Private Sub mnuSS_Click()
Dim FU
Dim I
I = "Enter a number between 0 and 7 " + "for the FillStyle"
FU = InputBox(I, "Setting the FillStyle")
Me.Cls
If Val(FU) >= 0 And Val(FU) <= 7 Then
Me.FillStyle = Val(FU)
Else
Beep
MsgBox ("Invalid FillStyle")
End If
Me.Line (100, 100)-Step(8265, 5625), RGB(0, 0, 0), B
End Sub

Private Sub Timer1_Timer()
Dim R, G, B
Dim X, Y
Dim Counter

If DP = 1 Then
For Counter = 1 To 500 Step 1
R = Rnd * 255
G = Rnd * 255
B = Rnd * 255
X = Rnd * Me.ScaleWidth
Y = Rnd * Me.ScaleHeight
Me.PSet (X, Y), RGB(R, G, B)
Next
End If

If DL = 1 Then
For Counter = 1 To 500 Step 1
R = Rnd * 255
G = Rnd * 255
B = Rnd * 255
Me.PSet Step(1, 1), RGB(R, G, B)
If CurrentX >= Me.ScaleWidth Then
CurrentX = Rnd * Me.ScaleWidth
End If

If CurrentY >= Me.ScaleHeight Then
CurrentY = Rnd * Me.ScaleHeight
End If
Next
End If
If DML = 1 Then
For Counter = 1 To 1 Step 1
R = Rnd * 255
G = Rnd * 255
B = Rnd * 255
Line -(Rnd * Me.ScaleWidth, Rnd * Me.ScaleHeight), RGB(R, G, B)
Next
End If

End Sub
