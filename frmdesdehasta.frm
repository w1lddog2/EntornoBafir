VERSION 5.00
Begin VB.Form frmdesdehasta 
   Caption         =   "Seleccion Desde - Hasta"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmdesdehasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public deSde As Date
Public haSta As Date

Private Sub Command1_Click()
deSde = Combo1.Text
haSta = Combo2.Text
Unload Me
End Sub

