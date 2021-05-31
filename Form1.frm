VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminicio 
   Caption         =   "Inicio"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtnumerousuario 
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtnumeroclave 
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Bravo\Desktop\Git\G1\Laboratorio.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Bravo\Desktop\Git\G1\Laboratorio.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Administrador"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtcontraseña 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
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
      Left            =   2280
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtusuario 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Laboratorios ""El Puente"" "
      BeginProperty Font 
         Name            =   "Segoe MDL2 Assets"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1215
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   2535
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
