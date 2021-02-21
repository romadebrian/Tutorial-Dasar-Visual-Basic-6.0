VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Operator Test"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3765
   LinkTopic       =   "Form4"
   ScaleHeight     =   4665
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Op. Logika"
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   3735
      Begin VB.OptionButton Option14 
         Caption         =   "And"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Or"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Not"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Op. Perbandingan"
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
      Begin VB.OptionButton Option11 
         Caption         =   "<="
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option10 
         Caption         =   "<>"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "<"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   ">="
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
         Caption         =   "="
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   ">"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Op. Aritmatika"
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   3735
      Begin VB.OptionButton Option5 
         Caption         =   "/"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&&"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "*"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "+"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Var 2 :"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Var 1 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var1 As Single, var2 As Single
Dim hasil As Single

Private Sub Option1_Click(Index As Integer)
var1 = Text1.Text
var2 = Text2.Text
hasil = var1 + var2
Label3.Caption = hasil
End Sub
Private Sub Option2_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = var1 - var2
Label3.Caption = hasil
End Sub
Private Sub Option3_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = var1 * var2
Label3.Caption = hasil
End Sub
Private Sub Option4_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = var1 / var2
Label3.Caption = hasil
End Sub
Private Sub Option5_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = var1 & var2
Label3.Caption = hasil
End Sub
Private Sub Option6_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = (var1 > var2)
'Label3.Caption = hasil
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option7_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = (var1 < var2)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option8_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = (var1 = var2)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option9_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = (var1 <> var2)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option10_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = (var1 >= var2)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option11_Click()
var1 = Text1.Text
var2 = Text2.Text
hasil = (var1 <= var2)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option12_Click()
var1 = IIf(Text1.Text = "True", -1, 0)
hasil = Not (var1)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option13_Click()
var1 = IIf(Text1.Text = "True", -1, 0)
var2 = IIf(Text2.Text = "True", -1, 0)
hasil = (var1 Or var2)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Option14_Click()
var1 = IIf(Text1.Text = "True", -1, 0)
var2 = IIf(Text2.Text = "True", -1, 0)
hasil = (var1 And var2)
Label3.Caption = Format(hasil, "True/False")
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub
