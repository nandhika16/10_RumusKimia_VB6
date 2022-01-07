VERSION 5.00
Begin VB.Form Program_Rumus_Kimia 
   Caption         =   "Program Kimia"
   ClientHeight    =   6165
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PEMROGRAMAN API"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KELOMPOK 4"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4290
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6060
      Left            =   0
      Picture         =   "Program_Rumus_Kimia.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
   Begin VB.Menu Menu_Awal 
      Caption         =   "Menu Awal"
      Begin VB.Menu Konversi_Suhu 
         Caption         =   "Konversi Suhu"
      End
      Begin VB.Menu RKU 
         Caption         =   "Rumus Kimia Unsur"
      End
      Begin VB.Menu RMS 
         Caption         =   "Rumus Molekul Senyawa"
      End
      Begin VB.Menu RE 
         Caption         =   "Rumus Empiris"
      End
      Begin VB.Menu Jumlah_Partikel 
         Caption         =   "Jumlah Partikel"
      End
      Begin VB.Menu Massa 
         Caption         =   "Massa(m)"
      End
      Begin VB.Menu Molaritas 
         Caption         =   "Molaritas(M)"
      End
      Begin VB.Menu Volume_gas 
         Caption         =   "Volume gas(v)"
      End
      Begin VB.Menu Molaritas_Campuran 
         Caption         =   "Molaritas Campuran"
      End
      Begin VB.Menu PEN 
         Caption         =   "PEN"
      End
   End
   Begin VB.Menu kelaur 
      Caption         =   "keluar"
   End
End
Attribute VB_Name = "Program_Rumus_Kimia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Jumlah_Partikel_Click()
Form_Jumlah_Partikel.Show
End Sub

Private Sub Konversi_Suhu_Click()
Form_Konversi_Suhu.Show
End Sub

Private Sub Massa_Click()
Form_Massa.Show
End Sub

Private Sub Molaritas_Campuran_Click()
Form_M_Campuran.Show
End Sub

Private Sub Molaritas_Click()
Form_Molaritas.Show
End Sub

Private Sub PEN_Click()
Form_PEN.Show
End Sub

Private Sub RE_Click()
Form_RE.Show
End Sub

Private Sub RKU_Click()
Form_RKU.Show
End Sub

Private Sub RMS_Click()
Form_RMS.Show
End Sub

Private Sub Volume_gas_Click()
Form_Stokiometri_Volume.Show
End Sub

