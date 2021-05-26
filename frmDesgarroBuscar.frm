VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDesgarroBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Desgarro"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Compuesto"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Partida"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDesgarroBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset

MSFlexGrid1.Clear
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
If Option1.Value = True Then
    buscar = "Compuesto"
End If
If Option2.Value = True Then
    buscar = "Partida"
End If
If Option3.Value = True Then
    buscar = "Código"
End If

Select Case buscar

Case "Compuesto"
Set rs = db.OpenRecordset("SELECT * FROM desgarros where compuesto = '" & Text1.Text & "'")






Case "Partida"
Set rs = db.OpenRecordset("SELECT * FROM desgarros where partida = '" & Text1.Text & "'")





Case "Código"
Set rs = db.OpenRecordset("SELECT * FROM desgarros where ensayo = " & Text1.Text)

End Select

If rs.RecordCount = 0 Then
    sdfsdf = MsgBox("No se encuentra el item buscado", vbCritical + vbOKOnly, "No se encuentra")
    db.Close
    Exit Sub
End If
rs.MoveLast
filr = rs.RecordCount
rs.MoveFirst

MSFlexGrid1.Cols = 7
MSFlexGrid1.Rows = filr + 1

MSFlexGrid1.TextMatrix(0, 0) = "Ensayo"
MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
MSFlexGrid1.TextMatrix(0, 2) = "Partida"
MSFlexGrid1.TextMatrix(0, 3) = "Probeta"
MSFlexGrid1.TextMatrix(0, 4) = "Valor"
MSFlexGrid1.TextMatrix(0, 5) = "Promedio"
MSFlexGrid1.TextMatrix(0, 6) = "Fecha"


For i = 1 To filr
MSFlexGrid1.TextMatrix(i, 0) = rs.Fields("ensayo")
MSFlexGrid1.TextMatrix(i, 1) = rs.Fields("Compuesto")
MSFlexGrid1.TextMatrix(i, 2) = rs.Fields("Partida")
MSFlexGrid1.TextMatrix(i, 3) = rs.Fields("Probeta")
MSFlexGrid1.TextMatrix(i, 4) = rs.Fields("Valor")
serie = Explode("@", rs.Fields("valor"))

cantidad = UBound(serie)

If cantidad > 0 Then
    incremento = 0
    For ll = 0 To cantidad
        incremento = incremento + CDbl(serie(ll))
    Next
    promedio = incremento / (cantidad + 1)
Else
    promedio = rs.Fields("valor")
End If


MSFlexGrid1.TextMatrix(i, 5) = promedio & " kg/mm"

MSFlexGrid1.TextMatrix(i, 6) = rs.Fields("Fecha")
rs.MoveNext
Next

AutoGrid MSFlexGrid1
db.Close
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
Me.Hide
End Sub

Private Sub Text1_Change()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 1
End Sub
