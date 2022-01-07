VERSION 5.00
Begin VB.Form Form_M_Campuran 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
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
      Left            =   7320
      TabIndex        =   15
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   3960
      TabIndex        =   11
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   3960
      TabIndex        =   10
      Top             =   5760
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3960
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3960
      TabIndex        =   2
      Top             =   3480
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
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9840
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Zat B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Molaritas (M)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Volume zat (V)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Zat A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label judul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOLARITAS CAMPURAN"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3300
      TabIndex        =   7
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Molaritas (M)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Volume zat (V)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Molaritas (M)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   8220
      Left            =   0
      Picture         =   "M_Campuran.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "Form_M_Campuran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
MolarA = Text1.Text
VolumeA = Text2.Text
VolumeB = Text4.Text
MolarB = Text5.Text
Hasil = (MolarA * VolumeA) + (MolarB * VolumeB) / (VolumeA + VolumeB)
Text3.Text = Hasil

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text5.Text = ""
Text2.Text = ""
Text4.Text = ""
Text3.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
