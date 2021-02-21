VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Property Test"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Pilihan"
      Height          =   1455
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
      Begin VB.CheckBox Check3 
         Caption         =   "Garis Bawah"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Miring"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tebal"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Hijau"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Merah"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Biru"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SELESAI"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tuliskan Nama Anda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Label2.FontBold = Check1.Value
End Sub

Private Sub Check2_Click()
Label2.FontItalic = Check2.Value
End Sub

Private Sub Check3_Click()
Label2.FontUnderline = Check3.Value
End Sub

Private Sub Command1_Click()
    Label2.Caption = Text1.Text
End Sub

Private Sub Command2_Click()
    Load Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub

Private Sub Option1_Click()
    Label2.ForeColor = vbBlue
End Sub

Private Sub Option2_Click()
    Label2.ForeColor = vbRed
End Sub

Private Sub Option3_Click()
    Label2.ForeColor = vbGreen
End Sub
