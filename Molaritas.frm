VERSION 5.00
Begin VB.Form Form_Molaritas 
   Caption         =   "Molaritas"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
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
      Left            =   8040
      MaskColor       =   &H0000FFFF&
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
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
      Left            =   7680
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   3840
      Width           =   2415
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
      Left            =   7800
      MaskColor       =   &H0000FFFF&
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   4200
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   480
      X2              =   9480
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Molaritas (M)"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2220
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Volume gas (V)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Jumlah Mol (n)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Judul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOKIOMETRI MOLARITAS"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   5220
      Left            =   0
      Picture         =   "Molaritas.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "Form_Molaritas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
mol = Text1.Text
volume = Text2.Text

Text3.Text = mol / volume

End Sub

Private Sub List1_Click()

End Sub

Private Sub Label4_Click()
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

