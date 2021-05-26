VERSION 5.00
Begin VB.Form frmAEDPassword 
   Caption         =   "Ingrese Password"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   1005
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAEDPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
okas
End Sub

Sub okas()
Form1.AEDPassword = Text1.Text
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    okas
End If
End Sub
