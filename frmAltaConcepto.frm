VERSION 5.00
Begin VB.Form frmAltaConcepto 
   Caption         =   "Alta conceptos"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   1815
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmAltaConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")

rcset.Open "SELECT codigo, concepto FROM flujo_fondos_concepto where concepto = '" & Text1.Text & "'", CONN, adOpenStatic, adLockReadOnly
If rcset.RecordCount <> 0 Then
    asdasd = MsgBox("Ya existe ese concepto", vbCritical + vbOKOnly, "Error")
    Exit Sub
Else
    rcset.Close
    rcset.Open "SELECT codigo, concepto FROM flujo_fondos_concepto order by codigo desc", CONN, adOpenStatic, adLockOptimistic
    If rcset.RecordCount = 0 Then
        Codigo = 0
    Else
        Codigo = rcset.Fields("codigo")
    End If
    rcset.AddNew
    rcset.Fields("codigo") = Codigo + 1
    rcset.Fields("concepto") = Text1.Text
    rcset.Update
    CONN.Close
End If
asdasd = MsgBox("Concepto agregado", vbInformation + vbOKOnly, "Alta de concepto")
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Unload Me
End Sub
