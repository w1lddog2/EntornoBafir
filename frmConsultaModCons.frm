VERSION 5.00
Begin VB.Form frmConsultaModCons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar datos de la consulta"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmConsultaModCons.frx":0000
      Top             =   4080
      Width           =   7215
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   1485
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmConsultaModCons.frx":0008
      Top             =   2040
      Width           =   7215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Agregar a consulta"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Consulta"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Responsable"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmConsultaModCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM Consultas where codigo = " & Text1.Text)

If rs.RecordCount = 0 Then
asddsdf = MsgBox("El codigo buscado no se encuentra", vbCritical + vbOKOnly, "Error")
Text1.Text = ""
Exit Sub
End If
Text2.Text = rs.Fields("fecha_consulta")
Text3.Text = rs.Fields("cliente")
Text4.Text = rs.Fields("responsable_cons")
Text5.Text = rs.Fields("consulta")
Command2.Enabled = True
Text6.SetFocus
db.Close
End Sub

Private Sub Command2_Click()

If Text6.Text = "" Then
    kjjjj = MsgBox("Debe agregar algo", vbCritical + vbOKOnly, "Error")
    Text6.SetFocus
    Exit Sub
End If

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT consulta, estado FROM Consultas where codigo = " & Text1.Text)

rs.Edit
rs.Fields("consulta") = rs.Fields("consulta") & " //" & Date & "// " & Text6.Text
rs.Fields("estado") = False
rs.Update
sdfsfsfd = MsgBox("Se ha agregado satisfatoriamente sus datos a la consulta")
Text5.Text = Text5.Text & " //" & Date & "// " & Text6.Text
Text6.Text = ""
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Form1.Enabled = True
Form1.Visible = True
Me.Hide
End Sub

Private Sub Text1_Change()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Command2.Enabled = fa
End Sub
