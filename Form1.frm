VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memilih Semua Item di ListBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1320
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()   'Masukkan data ke listbox
  For i = 0 To 10
     List1.AddItem "List ke-" & i, i
  Next i
End Sub

Private Sub Command1_Click()   'Sorot seluruh item data
  For x = 0 To List1.ListCount - 1
      List1.Selected(x) = True
  Next x
End Sub

Private Sub Command2_Click()   'Hilangkan semua sorot 'sebelumnya
  For x = 0 To List1.ListCount - 1
      List1.Selected(x) = False
  Next x
End Sub

