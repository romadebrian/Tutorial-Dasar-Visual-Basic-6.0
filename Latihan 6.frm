VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Struktur SELECT…CASE"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form6"
   ScaleHeight     =   3345
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblDiskon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Diskon"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblJumlah 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Jumlah"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblHarga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Harga"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblBarang 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Barang"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Jumlah"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Pilih Barang :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "Disket"
List1.AddItem "Buku"
List1.AddItem "Kertas"
List1.AddItem "Pulpen"
End Sub
Private Sub Command1_Click()
Dim harga As Currency, total As Currency
Dim jumlah As Integer
Dim diskon As Single
Dim satuan As String
If List1.Text = "" Then
MsgBox "Anda belum memilih barang !!"
List1.ListIndex = 0
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "Anda belum mengisi jumlah barang !!"
Text1.SetFocus
Exit Sub
End If
Select Case List1.Text
Case "Disket"
harga = 35000
satuan = "Box"
Case "Buku"
harga = 20000
satuan = "Lusin"
Case "Kertas"
harga = 25000
satuan = "Rim"
Case "Pulpen"
harga = 10000
satuan = "Pak"
End Select
lblBarang.Caption = "Barang : " & List1.Text
lblHarga.Caption = "Harga : " & Format(harga, "Currency") & "/" & satuan
lblJumlah.Caption = "Jumlah : " & Text1.Text & " " & satuan
jumlah = Text1.Text
Select Case jumlah
Case Is < 10
diskon = 0
Case 10 To 20
diskon = 0.15
Case Else
diskon = 0.2
End Select
total = jumlah * (harga * (1 - diskon))
lblDiskon.Caption = "Diskon : " & Format(diskon, "0 %")
lblTotal.Caption = "Total Bayar : " & Format(total, "Currency")
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form12.Show
End Sub

Private Sub lblJumlah2_Click(Index As Integer)
End Sub

Private Sub lblTotal2_Click(Index As Integer)

End Sub

