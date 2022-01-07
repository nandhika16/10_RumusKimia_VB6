VERSION 5.00
Begin VB.Form Form_RMS 
   Caption         =   "rms"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   6
      Top             =   5520
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Tampilkan 
      Caption         =   "Tampilkan"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Left            =   4320
      TabIndex        =   2
      Top             =   2280
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
      Left            =   840
      TabIndex        =   10
      Top             =   1200
      Width           =   9255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Jumlah Atom"
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
      Left            =   6120
      TabIndex        =   9
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Rumus Molekul Senyawa"
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
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Masukkan 
      Caption         =   "Masukkan Senyawa Kimia :"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RUMUS MOLEKUL SENYAWA"
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
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   7740
      Left            =   0
      Picture         =   "RMS.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "Form_RMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Tampilkan_Click()
If Text1.Text = "Karbon dioksida" Then
Text2.Text = "CO2"
Text3.Text = "3 atom (1 atom C dan 2 atom O)"

ElseIf Text1.Text = "Air" Then
Text2.Text = "H2O"
Text3.Text = "3 atom (2 atom H dan 1 atom O)"

ElseIf Text1.Text = "Amonia" Then
Text2.Text = "NH3"
Text3.Text = "4 atom (1 atom N dan 3 atom H)"

ElseIf Text1.Text = "Asam klorida" Then
Text2.Text = "HCl"
Text3.Text = "2 atom (1 atom H dan 1 atom Cl)"

ElseIf Text1.Text = "Asam sulfat" Then
Text2.Text = "H2So4"
Text3.Text = "7 atom (2 atom H, 1 atom S dan 4 atom O)"

ElseIf Text1.Text = "Glukosa" Then
Text2.Text = "C6H12O6"
Text3.Text = "24 atom (6 atom C, 12 atom H dan 6 atom O)"

ElseIf Text1.Text = "Urea" Then
Text2.Text = "CO(NH)2"
Text3.Text = "6 atom (1 atom C, 1 atom O, 2 atom N dan 2 atom H)"

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
