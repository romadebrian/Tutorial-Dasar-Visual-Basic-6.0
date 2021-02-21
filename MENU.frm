VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "MENU"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form12"
   ScaleHeight     =   7710
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "Demo Event Keyboard"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   6600
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Form 2 - ListBox"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Form 3 - Variable String, Int, Constanta"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   4935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Form 4 - Penggunaan Operator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   4935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Form 5 - IF ELSE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   4935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Form 6 - SELECT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   4935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Form 7 - Looping"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   4935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Form 8 - Array"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   4200
      Width           =   4935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Form 9 - Kalkulator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Width           =   4935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Form 10 - Error Handling"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   4935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Form 11 - Pemanggilan Procedure"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   6000
      Width           =   4935
   End
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   7200
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Form 1 - Text Color"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Form1.Show
Form12.Hide
End Sub

Private Sub Command10_Click(Index As Integer)
Form10.Show
Form12.Hide
End Sub

Private Sub Command11_Click(Index As Integer)
Form11.Show
Form12.Hide
End Sub

Private Sub Command12_Click()
DemoEventKeyboard.Show
Form12.Hide
End Sub

Private Sub Command2_Click(Index As Integer)
Form2.Show
Form12.Hide
End Sub

Private Sub Command3_Click(Index As Integer)
Form3.Show
Form12.Hide
End Sub

Private Sub Command4_Click(Index As Integer)
Form4.Show
Form12.Hide
End Sub

Private Sub Command5_Click(Index As Integer)
Form5.Show
Form12.Hide
End Sub

Private Sub Command6_Click(Index As Integer)
Form6.Show
Form12.Hide
End Sub

Private Sub Command7_Click(Index As Integer)
Form7.Show
Form12.Hide
End Sub

Private Sub Command8_Click(Index As Integer)
Form8.Show
Form12.Hide
End Sub

Private Sub Command9_Click(Index As Integer)
Form9.Show
Form12.Hide
End Sub

Private Sub exit_Click(Index As Integer)
End
End Sub
