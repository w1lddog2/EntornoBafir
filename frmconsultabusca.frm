VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmconsultabusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar consulta"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3720
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5953
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   810
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por"
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.OptionButton Option3 
         Caption         =   "Palabra suelta"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      Caption         =   "* Haga doble click en el registro para ver los detalles del mismo"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmconsultabusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
If Option1.Value = True Then 'codigo
    If Not IsNumeric(Text1.Text) Then
    sdfsdfsdf = MsgBox("Debe ingresar un valor numérico", vbCritical + vbOKOnly, "Error")
    Exit Sub
    End If
    Set rs = db.OpenRecordset("Select * from Consultas where codigo = " & Text1.Text)
End If
If Option2.Value = True Then 'cliente
    Set rs = db.OpenRecordset("Select * from Consultas where cliente = '" & Combo1.Text & "'")
End If
If Option3.Value = True Then 'palabra
    Set rs = db.OpenRecordset("Select * from Consultas where consulta like '*" & Text1.Text & "*' or respuesta like '*" & Text1.Text & "*'")
End If
    
    MSFlexGrid1.Clear
    If rs.RecordCount = 0 Then
        sdfdfsdf = MsgBox("No se han encontrado resultados para la búsqueda realizada", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
    rs.MoveLast
    filas = rs.RecordCount
    rs.MoveFirst
    MSFlexGrid1.Cols = 8
    MSFlexGrid1.Rows = filas + 1
    
    MSFlexGrid1.TextMatrix(0, 0) = "Código"
    MSFlexGrid1.TextMatrix(0, 1) = "Fecha cons."
    MSFlexGrid1.TextMatrix(0, 2) = "Cliente"
    MSFlexGrid1.TextMatrix(0, 3) = "resp.cons."
    MSFlexGrid1.TextMatrix(0, 4) = "Medio"
    MSFlexGrid1.TextMatrix(0, 5) = "Temperatura"
    MSFlexGrid1.TextMatrix(0, 6) = "Resp."
    MSFlexGrid1.TextMatrix(0, 7) = "Fecha Res"
        h = 1
    Do Until rs.EOF = True
        MSFlexGrid1.TextMatrix(h, 0) = rs.Fields("codigo")
        MSFlexGrid1.TextMatrix(h, 1) = rs.Fields("fecha_consulta")
        MSFlexGrid1.TextMatrix(h, 2) = rs.Fields("cliente")
        MSFlexGrid1.TextMatrix(h, 3) = rs.Fields("responsable_cons")
        'MSFlexGrid1.TextMatrix(h, 4) = rs.Fields("consulta")
        MSFlexGrid1.TextMatrix(h, 4) = rs.Fields("Medio") & ""
        MSFlexGrid1.TextMatrix(h, 5) = rs.Fields("temperatura") & ""
        If rs.Fields("estado") = True Then
            MSFlexGrid1.TextMatrix(h, 6) = "Si"
        Else
            MSFlexGrid1.TextMatrix(h, 6) = "No"
        End If
        MSFlexGrid1.TextMatrix(h, 7) = rs.Fields("Fecha_Resp") & ""

        h = h + 1
        rs.MoveNext
    Loop
    db.Close
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
frmconsultabusca.Hide
End Sub


Private Sub MSFlexGrid1_DblClick()

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from Consultas where codigo = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))

frmconsultaver.Text1.Text = rs.Fields("codigo")
frmconsultaver.Text2.Text = rs.Fields("fecha_consulta")
frmconsultaver.Combo1.Text = rs.Fields("cliente")
frmconsultaver.Combo2.Text = rs.Fields("responsable_cons")
frmconsultaver.Text3.Text = rs.Fields("consulta") & ""
frmconsultaver.Text4.Text = rs.Fields("medio") & ""
frmconsultaver.Text5.Text = rs.Fields("temperatura") & ""
frmconsultaver.Text6.Text = rs.Fields("respuesta") & ""
frmconsultaver.Text7.Text = rs.Fields("compuesto_elastomero") & ""
frmconsultaver.Text8.Text = rs.Fields("responsable_resp") & ""
frmconsultaver.Text9.Text = rs.Fields("fecha_resp") & ""
frmconsultaver.Text11.Text = rs.Fields("presion") & ""
frmconsultaver.Text10.Text = rs.Fields("uso") & ""

frmconsultabusca.Enabled = False
frmconsultaver.Show

db.Close
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Text1.Visible = True
    Combo1.Visible = False
Else
    Text1.Visible = False
    Combo1.Visible = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    Text1.Visible = False
    Combo1.Visible = True
Else
    Text1.Visible = True
    Combo1.Visible = False
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    Text1.Visible = True
    Combo1.Visible = False
Else
    Text1.Visible = False
    Combo1.Visible = True
End If
End Sub
