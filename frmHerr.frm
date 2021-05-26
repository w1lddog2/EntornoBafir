VERSION 5.00
Begin VB.Form frmHerr 
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmHerr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ReDim destinatarios(1)
indicedestinatarios = 1
asunto = "Prueba"
mail = "Prueba"
destinatarios(1) = "pablopirri@bafir.com.ar"
'frmSendinfo.Show
'frmSendinfo.Hide
frmSendinfo.Show
frmSendinfo.Visible = False
Call Moduloenvio
frmSendinfo.Hide
End Sub
