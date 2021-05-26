VERSION 5.00
Begin VB.Form frmBackSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione Listado"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Seleccionar"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Nuevo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Listado"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmBackSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tabla As String

Private Sub Command2_Click()
If List1.Text = "" Then
dfsadfsdf = MsgBox("Debe seleccionar un listado", vbCritical + vbOKOnly, "Error")
List1.SetFocus
Exit Sub
End If
tabla = List1.Text
frmBackSel.Hide
End Sub
