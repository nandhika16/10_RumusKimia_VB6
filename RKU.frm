VERSION 5.00
Begin VB.Form Form_RKU 
   Caption         =   "rku"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   6
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      TabIndex        =   5
      Top             =   6600
      Width           =   3735
   End
   Begin VB.CommandButton Hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Tampilkan 
      Caption         =   "Tampilkan"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Deskripsi : Memasukkan rumus unsur kimia harus menggunakan huruf kapital di awal penulisan!"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   8295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Hasil  Rumus Kimia Unsur"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Masukkan Rumus Kimia Unsur :"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RUMUS KIMIA UNSUR"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   8820
      Left            =   0
      Picture         =   "RKU.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "Form_RKU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Tampilkan_Click()
If Text1.Text = "Aluminium" Then
Text2.Text = "Al"

ElseIf Text1.Text = "Besi" Then
Text2.Text = "Fe"

ElseIf Text1.Text = "Emas" Then
Text2.Text = "Au"

ElseIf Text1.Text = "Helium" Then
Text2.Text = "He"

ElseIf Text1.Text = "Karbon" Then
Text2.Text = "C"

ElseIf Text1.Text = "Magnesium" Then
Text2.Text = "Mg"

ElseIf Text1.Text = "Natrium" Then
Text2.Text = "Na"

ElseIf Text1.Text = "Platina" Then
Text2.Text = "Pt"

ElseIf Text1.Text = "Raksa" Then
Text2.Text = "Hg"

ElseIf Text1.Text = "Xenon" Then
Text2.Text = "Xe"

End If
End Sub

Private Sub Hapus_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Keluar_Click()
Unload Me
End Sub


