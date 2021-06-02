VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmbuscar 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Selección de reactivos y marcadores tumorales"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14865
   LinkTopic       =   "Form2"
   ScaleHeight     =   7800
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11880
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      RecordSource    =   "Reactivos"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   3615
      Left            =   5160
      TabIndex        =   10
      Top             =   2160
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Salir"
      Height          =   375
      Left            =   12720
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtbuscartexto 
      Height          =   405
      Left            =   5880
      TabIndex        =   8
      Top             =   1560
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
      Left            =   13080
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione un reactivo de la lista"
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
      Width           =   5415
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
    If Len(Trim(txtbuscartexto.Text)) = 0 Then
        MsgBox "Seleccione un reactivo", vbInformation, "Laboratorios el Puente"
        txtbuscartexto.SetFocus
    Else
       
            rs.Find "NombreReactivos = '" & txtbuscartexto.Text & "'", , , 1
        
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
                
                frmuso.Show
                Unload Me
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

Private Sub DataGrid1_Click()
    txtbuscartexto.Text = DataGrid1.Columns(1).Text
End Sub

Private Sub Form_Load()
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Bravo\Desktop\Git\G1\Laboratorio.mdb;Persist Security Info=False"
    rs.Source = "contactos"
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.Open "select * from Reactivos", cn 'Abrir recordset
    rs.MoveFirst 'Mover al principio
    formato
    
End Sub
Sub formato()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(5).Width = 0
End Sub
