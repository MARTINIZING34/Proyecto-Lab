VERSION 5.00
Begin VB.Form frmbuscar 
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
      Caption         =   "Salir"
      Height          =   375
      Left            =   11400
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   5280
      List            =   "Form2.frx":0007
      TabIndex        =   9
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtbuscartexto 
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   2640
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
      Caption         =   "Utilizar"
      Height          =   735
      Left            =   10920
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Seleccione el tipo de búsqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Escriba el nombre del reactivo"
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Reactivos y marcadores tumorales"
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   4695
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

