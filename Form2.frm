VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   8400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Album"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   4575
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
File1.Pattern = "*.jpg ;*.bmp"
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub Dir1_Click()
File1.Path = Dir1.Path
File1.Pattern = "*.jpg;*.bmp"
If File1.ListCount <> 0 Then
File1.ListIndex = 0
End If
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Pattern = "*.jpg ;*.bmp"
If File1.ListCount <> 0 Then
File1.ListIndex = 0
End If
End Sub

Private Sub cmdLoad_Click()
Form1.Show
Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Me.Hide
End Sub

