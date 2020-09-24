VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Flip Album2"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFlip 
      Caption         =   "Flip"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtRight 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox txtLeft 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image Image1 
      Height          =   4410
      Left            =   4380
      Stretch         =   -1  'True
      Top             =   420
      Width           =   3285
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   525
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   4290
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   450
      Width           =   3180
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      Height          =   2280
      Left            =   6540
      Stretch         =   -1  'True
      Top             =   255
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image ImageDefault 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   4275
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image ImgBinder 
      Height          =   5535
      Left            =   4020
      Picture         =   "Form1.frx":0A51
      Stretch         =   -1  'True
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   2355
      Left            =   6060
      Stretch         =   -1  'True
      Top             =   300
      Width           =   285
   End
   Begin VB.Image imgBackGround 
      Height          =   5655
      Left            =   105
      Picture         =   "Form1.frx":730B
      Stretch         =   -1  'True
      Top             =   105
      Width           =   8055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load New Album"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W, H, T, L
Dim H1, W1, A1
Dim H2, W2, A2

Private Sub Form_Activate()
X1 = imgBackGround.Left + imgBackGround.Width
Y1 = imgBackGround.Top + imgBackGround.Height
Dim X, Y

Y = imgBackGround.Top

For X = X1 To X1 + 200 Step 30
Line (X, Y)-(X, Y + imgBackGround.Height)
Y = Y + 30
Next X


X = imgBackGround.Left
For Y = Y1 To Y1 + 200 Step 30
Line (X, Y)-(X + imgBackGround.Width, Y)
X = X + 30
Next Y

End Sub

Private Sub Form_Load()
Form1.Icon = LoadPicture(App.Path & "\Book02.ico")
imgBackGround.Picture = LoadPicture(App.Path & "\Club Deco.gif")
ImgBinder.Picture = LoadPicture(App.Path & "\Binder.jpg")
ImgBinder.Left = imgBackGround.Left + (imgBackGround.Width - ImgBinder.Width) / 2

ImageDefault.Picture = LoadPicture(App.Path & "\Blank.jpg")

T = ImageDefault.Top
L = ImageDefault.Left
W = ImageDefault.Width
H = ImageDefault.Height
End Sub


Private Sub mnuLoad_Click()
Form2.Show
Me.Hide
End Sub
Private Sub cmdFlip_Click()
MMControl1.filename = App.Path & "\Mode.wav"
MMControl1.Command = "Open"
MMControl1.Command = "Play"


If Form2.File1.ListIndex < Form2.File1.ListCount Then
Image1.Visible = False
Image1.Picture = LoadPicture(Form2.File1.Path & "\" & Form2.File1.filename)
Image1.Left = ImageDefault.Left
Image1.Top = ImageDefault.Top
Image1.Width = W


Image3.Picture = LoadPicture(App.Path & "\Blank.jpg")
Image3.Left = ImageDefault.Left
Image3.Top = ImageDefault.Top
Image3.Width = W

Dim SRatio
H1 = Image1.Picture.Height
W1 = Image1.Picture.Width

'*******************
If W1 > W Then
SRatio = W / W1 'Resise Ratio
W1 = W1 * SRatio
H1 = H1 * SRatio
End If
If H1 > H Then
SRatio = H / H1 'Resise Ratio
W1 = W1 * SRatio
H1 = H1 * SRatio
End If
Image1.Visible = True

Image1.Visible = True
txtLeft.Text = Form2.File1.filename

For V = W + ImgBinder.Width To 2 * W + ImgBinder.Width Step 500
PaintPicture Image1, L - ImgBinder.Width, T, W - V, H1
If H > H1 + 100 Then
PaintPicture Image3, L - ImgBinder.Width, T + H1, W - V, H - H1
End If
Next V
Image1.Visible = False


If Form2.File1.ListIndex < Form2.File1.ListCount - 1 Then
Form2.File1.ListIndex = Form2.File1.ListIndex + 1

Image2.Picture = LoadPicture(Form2.File1.Path & "\" & Form2.File1.filename)

Image2.Left = ImageDefault.Left
Image2.Top = ImageDefault.Top
Image2.Width = W

H2 = Image2.Picture.Height
W2 = Image2.Picture.Width
    If W2 > W Then
    SRatio = W / W2 'Resise Ratio
    W2 = W2 * SRatio
    H2 = H2 * SRatio
    End If
    If H2 > H Then
    SRatio = H / H2 'Resise Ratio
    W2 = W2 * SRatio
    H2 = H2 * SRatio
    End If


Image2.Visible = True
txtRight.Text = Form2.File1.filename

Else
Exit Sub
End If

If Form2.File1.ListIndex < Form2.File1.ListCount - 1 Then
    Form2.File1.ListIndex = Form2.File1.ListIndex + 1
Else
    Exit Sub
End If

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

