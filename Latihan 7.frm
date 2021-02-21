VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Struktur Looping"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Do While"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Do Until"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "For Next 2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "For Next 1"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Command1_Click()
    List1.Clear
    For i = 1 To 100
        List1.AddItem "Angka" & i
    Next i
End Sub

Private Sub Command2_Click()
    List1.Clear
    For i = 100 To 1 Step -2
        List1.AddItem "Angka" & i
    Next i
End Sub

Private Sub Command3_Click()
    List1.Clear
    i = Asc("A")
    Do Until i > Asc("Z")
        List1.AddItem "Huruf" & Chr(i)
        i = i + 1
    Loop
End Sub

Private Sub Command4_Click()
    List1.Clear
    i = Asc("Z")
    Do While i >= Asc("A")
        List1.AddItem "Huruf" & Chr(i)
        i = i - 1
    Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub
