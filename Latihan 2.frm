VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Method Test"
   ClientHeight    =   3090
   ClientLeft      =   1875
   ClientTop       =   2280
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.AddItem Combo1.Text
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
List1.Clear
End Sub

Private Sub Form_Load()
Combo1.AddItem "ROMA"
Combo1.AddItem "ANGGA"
Combo1.AddItem "ZAKI"
Combo1.AddItem "IRA"
Combo1.AddItem "AGUS"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub
