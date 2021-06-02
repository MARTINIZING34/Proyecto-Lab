VERSION 5.00
Begin VB.Form frmdatos 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "utilizar"
      Height          =   855
      Left            =   7080
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtmarca 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtfecha 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3600
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3600
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtnombre 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de esxpiracion"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Left            =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del reactivo seleccionado"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmdatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If MsgBox("¿Desea utilizar el reactivo seleccionado?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        frmuso.Show
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Desea regresar a la selección de reactivos?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        frmbuscar.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtnombre.Text = Nombre
    txtcantidad.Text = Cantidad
    txtfecha.Text = Fecha
    txtmarca.Text = Marca
End Sub
