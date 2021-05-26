VERSION 5.00
Begin VB.Form frmLista1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmLista1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
frmConsumos.dato = Combo1.Text
Unload Me
End Sub

