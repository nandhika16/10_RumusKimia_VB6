VERSION 5.00
Begin VB.Form Form_Massa 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3840
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tampilkan Hasil"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      MaskColor       =   &H0000FFFF&
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      MaskColor       =   &H0000FFFF&
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Judul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOKIOMETRI Massa"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3270
      TabIndex        =   9
      Top             =   600
      Width           =   3345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Masukkan Jumlah Mol (n)"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Masukkan Ar atau Mr "
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Massa (m)"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   480
      X2              =   9480
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Image Image1 
      Height          =   5700
      Left            =   0
      Picture         =   "Massa.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "Form_Massa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n = Text1.Text
r = Text2.Text

Text3.Text = n * r
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

