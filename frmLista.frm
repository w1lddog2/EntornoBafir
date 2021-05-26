VERSION 5.00
Begin VB.Form frmLista 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public seleccionado
Private Sub Command1_Click()
If List1.Text = "" Then
    sdfsdf = MsgBox("Debe seleccionar un elemento", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
frmLista.seleccionado = List1.Text
Me.Hide
End Sub

Private Sub Command2_Click()
frmLista.seleccionado = "Exit"
Me.Hide
End Sub
