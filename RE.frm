VERSION 5.00
Begin VB.Form Form_RE 
   Caption         =   "re"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton Tampilkan 
      Caption         =   "Tampilkan"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   600
      TabIndex        =   10
      Top             =   1440
      Width           =   8775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Hasil Rumus Empiris"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Hasil  Rumus Molekul"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RUMUS EMPIRIS"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Nama 
      Caption         =   "Masukkan Nama Senyawa :"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   8460
      Left            =   0
      Picture         =   "RE.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "Form_RE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Tampilkan_Click()

If Text1.Text = "Air" Then
Text2.Text = "H2O"
Text3.Text = "H2O"

ElseIf Text1.Text = "Amonia" Then
Text2.Text = "NH3"
Text3.Text = "NH3"

ElseIf Text1.Text = "Asam sulfat" Then
Text2.Text = "H2SO4"
Text3.Text = "H2SO4"

ElseIf Text1.Text = "Garap dapur" Then
Text2.Text = "NaCl"
Text3.Text = "NaCl"

ElseIf Text1.Text = "Garam Inggris" Then
Text2.Text = "MgSO4"
Text3.Text = "MgSO4"

ElseIf Text1.Text = "Asam cuka" Then
Text2.Text = "C2H4O2"
Text3.Text = "CH2O"

ElseIf Text1.Text = "Glukosa" Then
Text2.Text = "C6H12O6"
Text3.Text = "CH2O"

ElseIf Text1.Text = "Etana" Then
Text2.Text = "C2H6"
Text3.Text = "CH3"

ElseIf Text1.Text = "Etuna" Then
Text2.Text = "C2H2"
Text3.Text = "CH"

End If
End Sub

Private Sub Hapus_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Keluar_Click()
Unload Me
End Sub
