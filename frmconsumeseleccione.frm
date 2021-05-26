VERSION 5.00
Begin VB.Form frmconsumeseleccione 
   Caption         =   "Seleccione"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   1575
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmconsumeseleccione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public respuesta As String

Private Sub Command1_Click()
    respuesta = Combo1.Text
    Unload Me
End Sub

