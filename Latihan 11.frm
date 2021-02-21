VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Pemanggilan Procedure"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form11"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Function Test"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Dobel-Klik di Sini"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sub Test"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Procedure 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TulisTeks(teks As String, warna As ColorConstants)
    With Label1
    .Caption = teks
    .ForeColor = warna
    End With
End Sub

Private Function JumlahAngka() As String
    Dim angka1 As String, angka2 As String
    Dim hasil As Single
    angka1 = InputBox("Tulis angka 1 :", "Jumlah Angka")
    angka2 = InputBox("Tulis angka 2 :", "Jumlah Angka")
    If angka1 <> "" And angka2 <> "" Then
    hasil = CSng(angka1) + CSng(angka2)
    JumlahAngka = CStr(hasil)
    End If
End Function

Private Sub Label1_DblClick()
    Call TulisTeks("Hai", vbBlue)
End Sub

Private Sub Command1_Click()
    Call TulisTeks("Hallo", vbRed)
End Sub

Private Sub Command2_Click()
    Label1.Caption = "Jumlah = " & JumlahAngka()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form12.Show
End Sub
