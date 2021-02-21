VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Array Test"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form8"
   ScaleHeight     =   3090
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Redim "
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NamaSiswa() As String

Private Sub Form_Load()
    Dim i As Integer
    ReDim NamaSiswa(1 To 5)
    
    For i = 1 To 5
        Combo1.AddItem i
    Next i
    Combo1.ListIndex = 0
End Sub

Private Sub Command1_Click()
    Dim no As Integer, i As Integer
    
    no = CInt(Combo1.Text)
    NamaSiswa(no) = InputBox("tuliskan Nama Siswa No :" & no, "Input Nama Siswa")
    If NamaSiswa(no) <> "" Then
        List1.Clear
        For i = 1 To UBound(NamaSiswa)
            List1.AddItem "NamaSiswa(" & i & ")=" & NamaSiswa(i)
        Next i
    End If
End Sub

Private Sub Command2_Click()
    Dim num As Integer, i As Integer
    If Not IsNumeric(Text1.Text) Then Exit Sub
    num = CInt(Text1.Text)
    ReDim NamaSiswa(1 To num)
    Combo1.Clear
    List1.Clear
    For i = 1 To UBound(NamaSiswa)
        Combo1.AddItem i
        List1.AddItem "NamaSiswa(" & i & ")=" & NamaSiswa(i)
    Next i
    Combo1.ListIndex = 0
End Sub




    

Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub
