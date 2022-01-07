VERSION 5.00
Begin VB.Form Form_PEN 
   Caption         =   "PEN"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9705
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
      Height          =   735
      Left            =   8040
      MaskColor       =   &H0000FFFF&
      TabIndex        =   15
      Top             =   4680
      Width           =   1005
   End
   Begin VB.TextBox Neutron 
      Height          =   405
      Left            =   7560
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Proton 
      Height          =   405
      Left            =   7560
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Elektron 
      Height          =   405
      Left            =   7560
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox MA 
      Height          =   405
      Left            =   3720
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox NamaA 
      Height          =   405
      Left            =   3720
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox NA 
      Height          =   405
      Left            =   3720
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hasil"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      MaskColor       =   &H0000FFFF&
      TabIndex        =   1
      Top             =   4560
      Width           =   1005
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
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   1560
      Y2              =   4320
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Neutron ="
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Proton = "
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Elektron ="
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Massa Atom = "
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Judul 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Proton Elektron dan Neutron"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Nama Atom = "
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Masukkan Nomor Atom = "
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   5940
      Left            =   0
      Picture         =   "PEN.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Form_PEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Z As Integer
Dim A As Integer
Dim P As Integer
Dim E As Integer
Dim n As Integer

Private Sub Command1_Click()
Proton.Text = Z
Elektron.Text = Z
Neutron.Text = Val(A - Z)
End Sub

Private Sub Command2_Click()
NA.Text = ""
MA.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub MA_Change()
A = MA.Text
End Sub

Private Sub NA_Change()
Z = NA.Text
End Sub
