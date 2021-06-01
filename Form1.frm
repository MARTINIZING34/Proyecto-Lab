VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminicio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Inicio"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9375
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   2175
   End
   Begin VB.TextBox txtcontraseña 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   795
      Left            =   15240
      TabIndex        =   2
      Top             =   6840
      Width           =   3135
   End
   Begin VB.CommandButton cmdingresar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingresar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11760
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox txtusuario 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   795
      Left            =   15240
      TabIndex        =   0
      Top             =   4200
      Width           =   3135
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


Private Sub Form_Load()
    txtusuario.ForeColor = RGB(69, 110, 174)
    txtcontraseña.ForeColor = RGB(69, 110, 174)
    
End Sub

