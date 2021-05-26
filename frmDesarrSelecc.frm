VERSION 5.00
Begin VB.Form frmDesarrSelecc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione Formula a Ver"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmDesarrSelecc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Seleccionado

Private Sub Command1_Click()
Seleccionado = Combo1.Text
Me.Hide
End Sub

Private Sub Command2_Click()
Seleccionado = "Salir"
Me.Hide
End Sub
