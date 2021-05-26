VERSION 5.00
Begin VB.Form FrmDateIny 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione Rango de Fecha"
   ClientHeight    =   2370
   ClientLeft      =   5355
   ClientTop       =   2415
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Final"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicio"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "FrmDateIny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "" Or Combo2.Text = "" Then
    hgfh = MsgBox("Debe seleccionar una fecha", vbCritical + vbOKOnly, "Error")
    Combo1.SetFocus
    Exit Sub
End If
If CDate(Combo2.Text) < CDate(Combo1.Text) Then
    sdfet = MsgBox("La fecha 2 debe ser menor que la fecha 1", vbCritical + vbOKOnly, "Error")
    Combo2.SetFocus
    Exit Sub
End If
frmIny.Inicio = Combo1.Text
frmIny.Final = Combo2.Text
frmIny.Enabled = True
FrmDateIny.Hide

End Sub
