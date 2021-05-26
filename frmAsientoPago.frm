VERSION 5.00
Begin VB.Form frmAsientoPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asiento de Pagos"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Monto en $"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Fondo"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Concepto de pago"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de pago"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmAsientoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text1.Text = "" Then
    edsdfsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text1.SetFocus
    Exit Sub
Else
    If Not IsDate(Text1.Text) Then
        edsdfsdf = MsgBox("Debe introducir un formato válido de fecha", vbCritical + vbOKOnly, "Error")
        Text1.SetFocus
        Exit Sub
    End If
End If
If Combo1.Text = "" Then
    edsdfsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Combo1.SetFocus
    Exit Sub
End If
If Combo2.Text = "" Then
    edsdfsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Combo2.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    edsdfsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text2.SetFocus
    Exit Sub
Else
    If Not IsNumeric(Text2.Text) Then
        edsdfsdf = MsgBox("Debe introducir un formato válido de monto", vbCritical + vbOKOnly, "Error")
        Text2.SetFocus
        Exit Sub
    End If
End If
a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")

fecha = CDate(Text1.Text)
concepto = Combo1.Text
fondo = Combo2.Text
monto = Text2.Text

rcset.Open "SELECT codigo FROM flujo_fondos_concepto where concepto = '" & concepto & "'", CONN, adOpenStatic, adLockReadOnly
conceptocod = rcset.Fields("codigo")
rcset.Close
rcset.Open "SELECT codigo FROM flujo_fondos_fondos where fondo = '" & fondo & "'", CONN, adOpenStatic, adLockReadOnly
fondocod = rcset.Fields("codigo")
rcset.Close

' inicial = Format(inicial, "YYYY/MM/DD")

rcset.Open "SELECT * FROM flujo_fondos_asiento", CONN, adOpenStatic, adLockOptimistic
rcset.AddNew
rcset.Fields("fecha") = fecha
rcset.Fields("concepto") = conceptocod
rcset.Fields("fondo") = fondocod
a = CDbl(monto)
rcset.Fields("monto") = CDbl(monto)
rcset.Update
sdfsdfsdf = MsgBox("Pago asentado correctamente", vbInformation + vbOKOnly, "Asiento de pago")

Text1.Text = ""
'Combo1.Text = ""
'Combo2.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Combo1.SetFocus
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Command1.SetFocus
End If
End Sub
