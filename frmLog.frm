VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Clave"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.flags = 1
Form1.logUser = Text1.Text
Form1.logPass = Text2.Text
Unload Me
End Sub

Private Sub Command2_Click()
Form1.flags = 0
Unload Me
End Sub
