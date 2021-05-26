VERSION 5.00
Begin VB.Form frmAEDvisual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inspección Visual"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Introduzca datos de la inspección visual, tales como si existen grietas, cantidad, grietas pasantes, longitud de las grietas, etc."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmAEDvisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AEDvisual
Private Sub Command1_Click()
AEDvisual = Text1.Text
Me.Hide
End Sub

Private Sub Command2_Click()
AEDvisual = 0
Me.Hide
End Sub

