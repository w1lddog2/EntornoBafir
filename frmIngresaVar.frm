VERSION 5.00
Begin VB.Form frmIngresaVar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Variaciones"
   ClientHeight    =   1575
   ClientLeft      =   2430
   ClientTop       =   1410
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Var. Dureza"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Var. Elong."
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Var. Tracc."
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmIngresaVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmBuscaTraccion.Height = 5565
frmBuscaTraccion.Command2.Visible = False
frmBuscaTraccion.Command5.Visible = True
frmBuscaTraccion.Show (1)

End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
frmfluido.vartracc = 0
Else
frmfluido.vartracc = Text1.Text
End If
If Text2.Text = "" Then
frmfluido.varelong = 0
Else
frmfluido.varelong = Text2.Text
End If
If Text3.Text = "" Then
frmfluido.vardureza = 0
Else
frmfluido.vardureza = Text3.Text
End If
frmIngresaVar.Hide
End Sub

Private Sub Command3_Click()
frmfluido.vartracc = "0"
frmfluido.varelong = "0"
frmfluido.vardureza = "0"
frmIngresaVar.Hide
End Sub
