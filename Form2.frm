VERSION 5.00
Begin VB.Form frmbuscar 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form2"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14865
   LinkTopic       =   "Form2"
   ScaleHeight     =   5895
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Salir"
      Height          =   375
      Left            =   10680
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   5280
      List            =   "Form2.frx":0007
      TabIndex        =   9
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox txtbuscartexto 
      Height          =   405
      Left            =   5280
      TabIndex        =   8
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtmarca 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtfecha 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtcantidad 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Utilizar"
      Height          =   735
      Left            =   10080
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el tipo de búsqueda:"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba el nombre del reactivo:"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   4455
      Left            =   0
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reactivos y marcadores tumorales"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "frmbuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub Command1_Click()
    If Len(Trim(Combo1.Text)) = 0 Then
        MsgBox "Seleccione el tipo de busqueda", vbInformation, "Laboratorios el Puente"
        Combo1.SetFocus
    ElseIf Len(Trim(txtbuscartexto.Text)) = 0 Then
        MsgBox "Escriba el nombre del reactivo", vbInformation, "Laboratorios el Puente"
        txtbuscartexto.SetFocus
    Else
        If LCase(Combo1.Text) = LCase("NombreReactivos") Then
            rs.Find "NombreReactivos = '" & txtbuscartexto.Text & "'", , , 1
        End If
        If rs.BOF = False And rs.EOF = False Then
        'cargar datos a las cajas de texto
            txtnombre.Text = rs.Fields("NombreReactivos")
            txtcantidad.Text = rs.Fields("NúmeroReactivos")
            txtfecha.Text = rs.Fields("FechaExpiración")
            txtmarca.Text = rs.Fields("Marca")
            If MsgBox("¿Desea utilizar el reactivo?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
                Nombre = txtnombre.Text
                Cantidad = txtcantidad.Text
                Fecha = txtfecha.Text
                Marca = txtmarca.Text
                
                frmdatos.Show
            End If
        Else
            MsgBox "Reactivo incorrecto", vbInformation, "Laboratorios el Puente"
        End If
    End If
End Sub


Private Sub Command2_Click()
    If MsgBox("¿Desea salir?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Bravo\Desktop\Git\G1\Laboratorio.mdb;Persist Security Info=False"
    rs.Source = "contactos"
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.Open "select * from Reactivos", cn 'Abrir recordset
    rs.MoveFirst 'Mover al principio
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub
