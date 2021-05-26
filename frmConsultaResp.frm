VERSION 5.00
Begin VB.Form frmConsultaResp 
   Caption         =   "Responder Consulta"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   8235
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3360
      TabIndex        =   28
      Text            =   "Combo3"
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   5640
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmConsultaResp.frx":0000
      Top             =   1560
      Width           =   8775
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3840
      TabIndex        =   12
      Text            =   "Combo2"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   2880
      Width           =   5415
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmConsultaResp.frx":0006
      Top             =   4320
      Width           =   8775
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text7"
      Top             =   6480
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   7320
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Text            =   "Text9"
      Top             =   7200
      Width           =   2775
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   "Text10"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Text            =   "Text11"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Receptor de consulta"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente"
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de consulta"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo de consulta"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Medio"
      Height          =   255
      Left            =   1680
      TabIndex        =   21
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Temperatura"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Respuesta"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Responsable"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Resp."
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Uso"
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Presion"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "frmConsultaResp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text6.Text = "" And Text7.Text = "" Then
    dasdasda = MsgBox("Al menos respuesta o compuesto deben ser respondidos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Combo3.Text = "" Then
    asdasdasd = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Combo3.SetFocus
    Exit Sub
End If


Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM Consultas where codigo = " & Text1.Text)

rs.Edit
rs.Fields("respuesta") = Text6.Text
rs.Fields("compuesto_elastomero") = Text7.Text
rs.Fields("responsable_resp") = Combo3.Text
rs.Fields("fecha_resp") = Text9.Text
rs.Fields("estado") = True
rs.Update

adsfadsf = MsgBox("Se ha contestado satisfatoriamente la consulta", vbInformation + vbOKOnly, "Respuesta de consulta")

Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
Me.Hide
End Sub

Private Sub Command3_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM Consultas where codigo = " & Text1.Text)

If rs.RecordCount = 0 Then
dfasdf = MsgBox("El codigo buscado no existe", vbCritical + vbOKOnly, "Error")
Exit Sub
End If

If rs.Fields("estado") = True Then
    dsfasdfaf = MsgBox("Esta consulta se encuentra cerrada", vbCritical + vbOKOnly, "Consulta cerrada")
    Exit Sub
End If
Text2.Text = rs.Fields("fecha_consulta")
Combo1.Text = rs.Fields("cliente")
Combo2.Text = rs.Fields("responsable_cons")
Text3.Text = rs.Fields("consulta") & "" ' que en realidad son las observaciones
Text4.Text = rs.Fields("medio") & ""
Text5.Text = rs.Fields("temperatura") & ""
Command1.Enabled = True
End Sub

Private Sub Text1_Change()
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text10.Text = ""
Text11.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo3.Text = ""
Text9.Text = ""

Command1.Enabled = False
End Sub

Private Sub Text6_Change()
Text9.Text = Date
End Sub

Private Sub Text7_Change()
Text9.Text = Date
End Sub
