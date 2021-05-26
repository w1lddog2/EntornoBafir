VERSION 5.00
Begin VB.Form frmAsientoMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Asiento"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Monto"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Fondo"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmAsientoMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'grabar modif
a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
Dim rcset1 As ADODB.Recordset
Set rcset1 = New ADODB.Recordset
rcset1.CursorLocation = adUseClient
Dim rcset2 As ADODB.Recordset
Set rcset2 = New ADODB.Recordset
rcset2.CursorLocation = adUseClient
Dim rcset3 As ADODB.Recordset
Set rcset3 = New ADODB.Recordset
rcset3.CursorLocation = adUseClient

rcset.Open "SELECT * FROM flujo_fondos_asiento where codigo = " & frmInformefondos.MSFlexGrid1.TextMatrix(frmInformefondos.MSFlexGrid1.Row, 4), CONN, adOpenStatic, adLockOptimistic

'rcset.EditMode
rcset.Fields("fecha") = frmAsientoMod.Text1.Text

rcset.Fields("monto") = CDbl(frmAsientoMod.Text2.Text)
rcset1.Open "SELECT codigo FROM flujo_fondos_concepto where concepto = '" & Combo1.Text & "'", CONN, adOpenStatic, adLockReadOnly
rcset.Fields("concepto") = rcset1.Fields("codigo")
rcset1.Close
rcset1.Open "SELECT codigo FROM flujo_fondos_fondos where fondo = '" & Combo2.Text & "'", CONN, adOpenStatic, adLockReadOnly
rcset.Fields("fondo") = rcset1.Fields("codigo")
rcset1.Close
rcset.Update
frmInformefondos.MSFlexGrid1.TextMatrix(frmInformefondos.MSFlexGrid1.Row, 0) = Text1.Text
frmInformefondos.MSFlexGrid1.TextMatrix(frmInformefondos.MSFlexGrid1.Row, 1) = Combo2.Text
frmInformefondos.MSFlexGrid1.TextMatrix(frmInformefondos.MSFlexGrid1.Row, 2) = Combo1.Text
frmInformefondos.MSFlexGrid1.TextMatrix(frmInformefondos.MSFlexGrid1.Row, 3) = Text2.Text
Me.Enabled = False
asdasdasd = MsgBox("Registro modificado", vbInformation + vbOKOnly, "Modificación de registro")
frmAsientoMod.Show
CONN.Close
frmInformefondos.Enabled = True
frmInformefondos.Command7.Visible = False
Me.Hide
End Sub

Private Sub Command2_Click()
frmInformefondos.Enabled = True
frmInformefondos.Command7.Visible = False
frmAsientoMod.Hide


End Sub
