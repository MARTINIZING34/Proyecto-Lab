VERSION 5.00
Begin VB.Form frminicio 
   Caption         =   "Inicio"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7395
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   7680
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox txtnumerousuario 
      Height          =   615
      Left            =   7080
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtnumeroclave 
      Height          =   495
      Left            =   7320
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtnombre 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtcontraseña 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtclave 
      DataField       =   "Contraseña"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdingresar 
      Caption         =   "Ingresar"
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtusuario 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1095
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Laboratorios ""El Puente"" "
      BeginProperty Font 
         Name            =   "Segoe MDL2 Assets"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1815
      Left            =   4080
      TabIndex        =   0
      Top             =   -120
      Width           =   3375
   End
End
Attribute VB_Name = "frminicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdingresar_Click()

'Verificación del usuario

If txtusuario.Text = "" Then
    MsgBox "Ingrese su usuario", vbInformation, "Laboratorios el Puente"
    
Else
    If txtusuario.Text = txtnombre.Text Then
        txtnumerousuario.Text = 1
    Else
        MsgBox "Usuario incorrecto", vbInformation, "Laboratorios el Puente"
        
    End If
End If

'Verificación de la contraseña

If txtcontraseña.Text = "" Then
    MsgBox "Ingrese su contraseña", vbInformation, "Laboratorios el Puente"
    
Else
    If txtcontraseña.Text = txtclave.Text Then
        txtnumeroclave.Text = 1
    Else
        MsgBox "Contraseña incorrecta", vbInformation, "Laboratorios el Puente"
        
    End If
End If

'Ingreso al segundo formulario

If txtnumeroclave.Text = txtnumerousuario.Text Then
    Form2.Show
    Unload Me
End If
    
End Sub

Private Sub cmdsalir_Click()
    If MsgBox("¿Desea salir?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

'Datos para que la condición funcione

    txtnumeroclave.Text = 23
    txtnumerousuario.Text = 11
End Sub
