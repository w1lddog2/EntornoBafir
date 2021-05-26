VERSION 5.00
Begin VB.Form frmAltaFondos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta de fondos"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmAltaFondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
rcset.Open "SELECT codigo, fondo FROM flujo_fondos_fondos where fondo = '" & Text1.Text & "'", CONN, adOpenStatic, adLockReadOnly
If rcset.RecordCount <> 0 Then
    asdasd = MsgBox("Ya existe ese fondo", vbCritical + vbOKOnly, "Error")
    Exit Sub
Else
    rcset.Close
    rcset.Open "SELECT codigo, fondo FROM flujo_fondos_fondos order by codigo desc", CONN, adOpenStatic, adLockOptimistic
    If rcset.RecordCount = 0 Then
        Codigo = 0
    Else
        Codigo = rcset.Fields("codigo")
    End If
    rcset.AddNew
    rcset.Fields("codigo") = Codigo + 1
    rcset.Fields("fondo") = Text1.Text
    rcset.Update
    rcset.Close
End If
asdasd = MsgBox("Fondo agregado", vbInformation + vbOKOnly, "Alta de Fondo")
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Unload Me
End Sub
