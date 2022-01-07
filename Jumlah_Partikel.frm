VERSION 5.00
Begin VB.Form Form_Jumlah_Partikel 
   Caption         =   "Jmlh_Prtkl"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   Picture         =   "Jumlah_Partikel.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Hasil 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   7200
      Width           =   3975
   End
   Begin VB.CommandButton Keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   8
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   7
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton Tampilkan 
      Caption         =   "Tampilkan"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox n 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Jumlah_partikel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasil Jumlah Partikel"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "* 10^23"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "* 6,02"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Keterangan : "
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Masukkan_n 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Masukkan n "
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOKIOMETRI JUMLAH PARTIKEL"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   8520
      Left            =   0
      Picture         =   "Jumlah_Partikel.frx":8055
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "Form_Jumlah_Partikel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub n_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub
Private Sub Tampilkan_Click()
Hasil.Text = (n.Text * 6.02 * 10 ^ 23)
End Sub

Private Sub Hapus_Click()
Hasil.Text = ""
n.Text = ""
End Sub

Private Sub Keluar_Click()
Unload Me
End Sub
