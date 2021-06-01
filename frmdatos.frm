VERSION 5.00
Begin VB.Form frmdatos 
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "utilizar"
      Height          =   855
      Left            =   5880
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "marca"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha de esxpiracion"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "cantidad"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del reactivo seleccionado"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmdatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
