VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Variabel Test"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "TEST 3"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TEST 2"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST 1"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim test2 As Integer

Private Sub Command1_Click(Index As Integer)
Dim test1 As String
test1 = "nusantara"
Label1.Caption = test1
Label2.Caption = test2
Label3.Caption = test3
End Sub

Private Sub Command2_Click(Index As Integer)
test2 = 10
Label1.Caption = test1
Label2.Caption = test2
Label3.Caption = test3
End Sub

Private Sub Command3_Click(Index As Integer)
Const test3 As Single = 90.55
'test3 = 50.22
Label1.Caption = test1
Label2.Caption = test2
Label3.Caption = test3
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub
