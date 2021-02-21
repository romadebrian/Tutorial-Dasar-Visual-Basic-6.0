VERSION 5.00
Begin VB.Form DemoEventKeyboard 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Demo Event Keyboard"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form13"
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRoket 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1680
      Picture         =   "DemoEventKeyboard.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "DemoEventKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xpos As Single, ypos As Single
Private Sub Form_Load()
xpos = (Me.ScaleWidth - picRoket.Width) / 2
ypos = (Me.ScaleHeight - picRoket.Height) / 2
picRoket.Move xpos, ypos
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyLeft
Call RoketKeKiri
Case vbKeyRight
Call RoketKeKanan
End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
Call RoketKeAtas
Case vbKeyDown
Call RoketKeBawah
End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
If MsgBox("Tutup program ?", vbQuestion + vbYesNo, _
Me.Caption) = vbYes Then Unload Me
End If
End Sub
Private Sub RoketKeKiri()
xpos = xpos - 10
If xpos < 0 Then
xpos = 0
End If
picRoket.Move xpos
End Sub
Private Sub RoketKeKanan()
xpos = xpos + 10
If xpos > Me.ScaleWidth - picRoket.Width Then
xpos = Me.ScaleWidth - picRoket.Width
End If
picRoket.Move xpos
End Sub
Private Sub RoketKeAtas()
ypos = ypos - 10
If ypos < 0 Then
ypos = 0
End If
picRoket.Move xpos, ypos
End Sub
Private Sub RoketKeBawah()
ypos = ypos + 10
If ypos > Me.ScaleHeight - picRoket.Height Then
ypos = Me.ScaleHeight - picRoket.Height
End If
picRoket.Move xpos, ypos
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub
