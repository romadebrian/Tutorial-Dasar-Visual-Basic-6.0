VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Struktur IF…THEN"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3015
   LinkTopic       =   "Form5"
   ScaleHeight     =   4440
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "Latihan 5.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "nusantara" Then
Image1.Visible = True
Else
MsgBox "password salah"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub
