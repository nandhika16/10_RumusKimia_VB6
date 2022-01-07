VERSION 5.00
Begin VB.Form Form_Konversi_Suhu 
   Caption         =   "Konversi Suhu"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Celcius 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   9
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Konversi 
      Caption         =   "Konversi "
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fahrenheit"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reamur"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kelvin"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Fahrenheit 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Reamur 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Kelvin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Masukkan Celcius :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Suhu 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Konversi Suhu"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   7740
      Left            =   0
      Picture         =   "Konversi_Suhu.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "Form_Konversi_Suhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Celcius_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub

Private Sub Fahrenheit_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub


Private Sub Reamur_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub

Private Sub Kelvin_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub

Private Sub Celcius_Change()

If Celcius = "" Then
    Celcius.Enabled = True
    Fahrenheit.Enabled = True
    Reamur.Enabled = True
    Kelvin.Enabled = True
Else
    Celcius.Enabled = True
    Fahrenheit.Enabled = False
    Reamur.Enabled = False
    Kelvin.Enabled = False
End If

End Sub

Private Sub Fahrenheit_Change()

If Fahrenheit = "" Then
    Celcius.Enabled = True
    Fahrenheit.Enabled = True
    Reamur.Enabled = True
    Kelvin.Enabled = True
Else
    Celcius.Enabled = False
    Fahrenheit.Enabled = True
    Reamur.Enabled = False
    Kelvin.Enabled = False
End If

End Sub

Private Sub Reamur_Change()

If Celcius = "" Then
    Celcius.Enabled = True
    Fahrenheit.Enabled = True
    Reamur.Enabled = True
    Kelvin.Enabled = True
Else
    Celcius.Enabled = False
    Fahrenheit.Enabled = False
    Reamur.Enabled = True
    Kelvin.Enabled = False
End If

End Sub


Private Sub Kelvin_Change()

If Celcius = "" Then
    Celcius.Enabled = True
    Fahrenheit.Enabled = True
    Reamur.Enabled = True
    Kelvin.Enabled = True
Else
    Celcius.Enabled = False
    Fahrenheit.Enabled = False
    Reamur.Enabled = False
    Kelvin.Enabled = True
    
End If

End Sub

Private Sub Konversi_Click()

If (Fahrenheit = "" And Reamur = "" And Kelvin = "") Then
    Fahrenheit = (9 / 5) * Val(Celcius) + 32
    Reamur = (4 / 5) * Val(Celcius)
    Kelvin = Val(Celcius) + 273
ElseIf (Celius = "" And Reamur = "" And Kelvin = "") Then
    Celcius = (5 / 9) * (Val(Fahrenheit) - 32)
    Reamur = (4 / 9) * (Val(Fahrenheit) - 32)
    Kelvin = (5 / 9) * (Val(Fahrenheit) - 32) + 273
ElseIf (Celius = "" And Fahrenheit = "" And Kelvin = "") Then
    Celcius = (5 / 4) * Val(Reamur)
    Fahrenheit = (9 / 4) * Val(Reamur) + 32
    Kelvin = (5 / 4) * Val(Reamur) + 273
ElseIf (Celius = "" And Fahrenheit = "" And Reamur = "") Then
    Celcius = Val(Kelvin) - 273
    Fahrenheit = (9 / 5) * (Val(Kelvin) - 273) + 32
    Reamur = (4 / 5) * (Val(Kelvin) - 273)
Else
End If

Celcius.Enabled = False
Fahrenheit.Enabled = False
Reamur.Enabled = False
Kelvin.Enabled = False
End Sub

Private Sub Hapus_Click()
Celcius = ""
Fahrenheit = ""
Reamur = ""
Kelvin = ""

Celcius.Enabled = True
Fahrenheit.Enabled = True
Reamur.Enabled = True
Kelvin.Enabled = True
End Sub

Private Sub Keluar_Click()
Unload Me
End Sub

Private Sub Form_Load()

Celcius.Enabled = True
Fahrenheit.Enabled = True
Reamur.Enabled = True
Kelvin.Enabled = True

End Sub
