VERSION 5.00
Begin VB.Form frmuso 
   BackColor       =   &H8000000E&
   Caption         =   "Form3"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12645
   LinkTopic       =   "Form3"
   ScaleHeight     =   4665
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "salir"
      Height          =   1455
      Left            =   9600
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Usar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6960
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtcantidadutilizar 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtcantidad 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   3495
      Left            =   2400
      Top             =   1200
      Width           =   10215
   End
   Begin VB.Label lblnombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del reactivo"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad a utilizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmuso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        Resultado = Val(txtcantidad.Text) - Val(txtcantidadutilizar.Text)
        Text1.Text = Resultado
        
        frmdatos.Show
End Sub
Private Sub Command2_Click()
    If MsgBox("¿Desea volver a la seleccion de reactivos?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtcantidad.Text = Cantidad
    lblnombre.Caption = Nombre
    If txtcantidad.Text = 10 Then
        MsgBox "La cantidad restante de reactivos es 10 por favor regitre más reactivos"
    End If
End Sub
