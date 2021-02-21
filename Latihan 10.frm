VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Error Handle"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form10"
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   7560
      Picture         =   "Latihan 10.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Open Picture F"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   11055
      Left            =   0
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim FileName As String
    Dim ErrMsg As String
    Dim Ask As VbMsgBoxResult
    On Error GoTo AdaError
Awal:
    Image1.Picture = Nothing
    FileName = InputBox("Ketikkan path dan nama file gambar :", "Open Picture File", FileName)
    If FileName <> "" Then
        Image1.Picture = LoadPicture(FileName)
    End If
    Exit Sub
AdaError:
    Select Case Err.Number
        Case 53
            ErrMsg = "File [" & FileName & "] tidak ada !"
        Case 71
            ErrMsg = "Disket belum dimasukkan !"
        Case Else
            ErrMsg = Err.Description
        End Select
    Ask = MsgBox(ErrMsg, vbCritical + vbRetryCancel, Me.Caption)
    Select Case Ask
        Case vbRetry
            If Err.Number = 53 Then Resume Awal Else Resume
        Case vbCancel
        Resume Next
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form12.Show
End Sub
