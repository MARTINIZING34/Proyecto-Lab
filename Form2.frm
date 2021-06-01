VERSION 5.00
Begin VB.Form frmbuscar 
   Caption         =   "Form2"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15075
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Utilizar"
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   3000
      TabIndex        =   2
      Text            =   "Ingrese el nombre del reactivvo a utilizar"
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
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
