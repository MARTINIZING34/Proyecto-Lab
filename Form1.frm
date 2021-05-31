VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminicio 
   Caption         =   "Inicio"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtnombre 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtclave 
      DataField       =   "Contraseña"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from Administrador"
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
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   6
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox txtcontraseña 
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdingresar 
      Caption         =   "Ingresar"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox txtusuario 
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1095
      Left            =   3000
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Laboratorios ""El Puente"" "
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   2175
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frminicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdingresar_Click()

'Verificación de usuario y contraseña

Adodc1.RecordSource = "select * from Administrador where Nombre = '" + txtusuario.Text + "' and Contraseña = '" + txtcontraseña.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    MsgBox "Revise sus datos y vuelva a intentarlo", vbInformation, "Laboratorios el Puente "
Else
    Form2.Show
End If
'Verificación del usuario

'If txtusuario.Text = "" Then
    'MsgBox "Ingrese su usuario", vbInformation, "Laboratorios el Puente"
'Else
    'If txtusuario.Text = txtnombre.Text Then
        'If txtcontraseña.Text = "" Then
            'MsgBox "Ingrese su contraseña", vbInformation, "Laboratorios el Puente"
        'Else
            'If txtcontraseña.Text = txtclave.Text Then
                'Form2.Show
                'Unload Me
            'Else
                'MsgBox "Contraseña incorrecta", vbInformation, "Laboratorios el Puente"
            'End If
        'End If
    'Else
        'MsgBox "Usuario incorrecto", vbInformation, "Laboratorios el Puente"
        
    'End If
'End If

'Verificación de la contraseña



'Ingreso al segundo formulario


    
End Sub

Private Sub cmdsalir_Click()
    If MsgBox("¿Desea salir?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Unload Me
    End If
End Sub


