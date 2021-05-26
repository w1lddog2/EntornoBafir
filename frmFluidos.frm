VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFluidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fluidos de compuestos"
   ClientHeight    =   4695
   ClientLeft      =   1440
   ClientTop       =   1260
   ClientWidth     =   13680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Checkear Vencimientos"
      Height          =   615
      Left            =   9120
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cargar Nuevo"
      Height          =   615
      Left            =   9120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   615
      Left            =   7560
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ocultar tabla"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   13
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Buscar por lote"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Buscar por tiempo, temperatura y fluido"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Buscar por compuesto"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   10080
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmFluidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT probeta, Codigo, N_FORMULA, PARTIDA, TIEMP_TEMP_OIL, var_vol,var_vol1,var_vol2, var_tracc, var_elong, var_dur, var_masa,var_masa1,var_masa2, fecha_realizacion, tiempo_repeticion, aprovado FROM fluidos WHERE N_FORMULA = '" & Combo1.Text & "' ORDER BY CODIGO ;")

'rs.Index = "primarykey"
'rs.Seek "=", Combo1.Text
'If rs.NoMatch = True Then
If rs.RecordCount = 0 Then
sfd = MsgBox("El fluido para dicho compuesto no existe", vbCritical + vbOKOnly, "Error")
Exit Sub
End If
rs.MoveLast
total = rs.RecordCount
MSFlexGrid1.Cols = 19
MSFlexGrid1.Rows = total + 1
rs.MoveFirst
For contador = 1 To total
MSFlexGrid1.TextMatrix(contador, 0) = rs.Fields("codigo")
MSFlexGrid1.TextMatrix(contador, 1) = rs.Fields("N_FORMULA")
MSFlexGrid1.TextMatrix(contador, 2) = rs.Fields("partida")
MSFlexGrid1.TextMatrix(contador, 3) = rs.Fields("TIEMP_TEMP_OIL")
MSFlexGrid1.TextMatrix(contador, 4) = Format(rs.Fields("VAR_VOL"), "0.00")
MSFlexGrid1.TextMatrix(contador, 5) = Format(rs.Fields("VAR_VOL1"), "0.00") & ""
MSFlexGrid1.TextMatrix(contador, 6) = Format(rs.Fields("VAR_VOL2"), "0.00") & ""

'''''''''''''080124
If IsNull(rs.Fields("var_vol1")) Then
    MSFlexGrid1.TextMatrix(contador, 7) = Format(rs.Fields("var_vol"), "0.00")
Else
    MSFlexGrid1.TextMatrix(contador, 7) = Format((CDbl(rs.Fields("var_vol")) + CDbl(rs.Fields("var_vol1")) + CDbl(rs.Fields("var_vol2"))) / 3, "0.00")
End If
MSFlexGrid1.TextMatrix(contador, 8) = rs.Fields("VAR_TRACC")
MSFlexGrid1.TextMatrix(contador, 9) = rs.Fields("VAR_ELONG")
MSFlexGrid1.TextMatrix(contador, 10) = rs.Fields("VAR_DUR")
MSFlexGrid1.TextMatrix(contador, 11) = rs.Fields("FECHA_REALIZACION")
MSFlexGrid1.TextMatrix(contador, 12) = rs.Fields("Tiempo_repeticion")
MSFlexGrid1.TextMatrix(contador, 13) = rs.Fields("aprovado")
MSFlexGrid1.TextMatrix(contador, 14) = rs.Fields("probeta") & ""
MSFlexGrid1.TextMatrix(contador, 15) = Format(rs.Fields("var_masa"), "0.00") & ""
MSFlexGrid1.TextMatrix(contador, 16) = Format(rs.Fields("var_masa1"), "0.00") & ""
MSFlexGrid1.TextMatrix(contador, 17) = Format(rs.Fields("var_masa2"), "0.00") & ""
If IsNull(rs.Fields("var_masa1")) Then
    MSFlexGrid1.TextMatrix(contador, 18) = Format(rs.Fields("var_masa"), "0.00") & ""
Else
    MSFlexGrid1.TextMatrix(contador, 18) = Format((CDbl(rs.Fields("var_masa")) + CDbl(rs.Fields("var_masa1")) + CDbl(rs.Fields("var_masa2"))) / 3, "0.00")
End If


rs.MoveNext
Next
db.Close
frmFluidos.Height = 5070
MSFlexGrid1.TextMatrix(0, 0) = "codigo"
MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
MSFlexGrid1.TextMatrix(0, 2) = "Partida"
MSFlexGrid1.TextMatrix(0, 3) = "Fluido"
MSFlexGrid1.TextMatrix(0, 4) = "Var.Vol1"
MSFlexGrid1.TextMatrix(0, 5) = "Var.Vol2"
MSFlexGrid1.TextMatrix(0, 6) = "Var.Vol3"
MSFlexGrid1.TextMatrix(0, 7) = "Var.Vol.prom"
MSFlexGrid1.TextMatrix(0, 8) = "Var.Tracc"
MSFlexGrid1.TextMatrix(0, 9) = "Var.Elong"
MSFlexGrid1.TextMatrix(0, 10) = "Var.dur"
MSFlexGrid1.TextMatrix(0, 11) = "Fecha"
MSFlexGrid1.TextMatrix(0, 12) = "Tiempo.rep"
MSFlexGrid1.TextMatrix(0, 13) = "aprobado"
MSFlexGrid1.TextMatrix(0, 14) = "Probeta"
MSFlexGrid1.TextMatrix(0, 15) = "Var.Masa1"
MSFlexGrid1.TextMatrix(0, 16) = "Var.Masa2"
MSFlexGrid1.TextMatrix(0, 17) = "Var.Masa3"
MSFlexGrid1.TextMatrix(0, 18) = "Var.Masa.prom"
'For numero = 0 To 9
'Next
AutoGrid MSFlexGrid1
End Sub

Private Sub Combo2_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT var_masa,var_masa1,var_masa2, probeta, codigo, N_FORMULA, PARTIDA, TIEMP_TEMP_OIL, var_vol,var_vol1,var_vol2, var_tracc, var_elong, var_dur, fecha_realizacion, tiempo_repeticion, aprovado FROM fluidos WHERE TIEMP_TEMP_OIL = '" & Combo2.Text & "' ;")
If rs.RecordCount > 0 Then
rs.MoveLast
total = rs.RecordCount
MSFlexGrid1.Cols = 12
MSFlexGrid1.Rows = total + 1
rs.MoveFirst
For contador = 1 To total
'MSFlexGrid1.TextMatrix(contador, 0) = rs.Fields("codigo")
'MSFlexGrid1.TextMatrix(contador, 1) = rs.Fields("N_FORMULA")
'MSFlexGrid1.TextMatrix(contador, 2) = rs.Fields("partida")
'MSFlexGrid1.TextMatrix(contador, 3) = rs.Fields("TIEMP_TEMP_OIL")
'MSFlexGrid1.TextMatrix(contador, 4) = rs.Fields("VAR_VOL")
'MSFlexGrid1.TextMatrix(contador, 5) = rs.Fields("VAR_TRACC")
'MSFlexGrid1.TextMatrix(contador, 6) = rs.Fields("VAR_ELONG")
'MSFlexGrid1.TextMatrix(contador, 7) = rs.Fields("VAR_DUR")
'MSFlexGrid1.TextMatrix(contador, 8) = rs.Fields("FECHA_REALIZACION")
'MSFlexGrid1.TextMatrix(contador, 9) = rs.Fields("Tiempo_repeticion")
'MSFlexGrid1.TextMatrix(contador, 10) = rs.Fields("aprovado")
'MSFlexGrid1.TextMatrix(contador, 11) = rs.Fields("probeta") & ""
'MSFlexGrid1.TextMatrix(contador, 12) = rs.Fields("var_masa") & ""
    MSFlexGrid1.TextMatrix(contador, 0) = rs.Fields("codigo")
    MSFlexGrid1.TextMatrix(contador, 1) = rs.Fields("N_FORMULA")
    MSFlexGrid1.TextMatrix(contador, 2) = rs.Fields("partida")
    MSFlexGrid1.TextMatrix(contador, 3) = rs.Fields("TIEMP_TEMP_OIL")
    MSFlexGrid1.TextMatrix(contador, 4) = rs.Fields("VAR_VOL")
    MSFlexGrid1.TextMatrix(contador, 5) = rs.Fields("VAR_VOL1") & ""
    MSFlexGrid1.TextMatrix(contador, 6) = rs.Fields("VAR_VOL2") & ""
    MSFlexGrid1.TextMatrix(contador, 7) = (rs.Fields("var_vol") + rs.Fields("var_vol1") + rs.Fields("var_vol2")) / 3
    MSFlexGrid1.TextMatrix(contador, 8) = rs.Fields("VAR_TRACC")
    MSFlexGrid1.TextMatrix(contador, 9) = rs.Fields("VAR_ELONG")
    MSFlexGrid1.TextMatrix(contador, 10) = rs.Fields("VAR_DUR")
    MSFlexGrid1.TextMatrix(contador, 11) = rs.Fields("FECHA_REALIZACION")
    MSFlexGrid1.TextMatrix(contador, 12) = rs.Fields("Tiempo_repeticion")
    MSFlexGrid1.TextMatrix(contador, 13) = rs.Fields("aprovado")
    MSFlexGrid1.TextMatrix(contador, 14) = rs.Fields("probeta") & ""
    MSFlexGrid1.TextMatrix(contador, 15) = rs.Fields("var_masa") & ""
    MSFlexGrid1.TextMatrix(contador, 16) = rs.Fields("var_masa1") & ""
    MSFlexGrid1.TextMatrix(contador, 17) = rs.Fields("var_masa2") & ""
    MSFlexGrid1.TextMatrix(contador, 18) = (rs.Fields("var_masa") + rs.Fields("var_masa1") + rs.Fields("var_masa2")) / 3
rs.MoveNext
Next
db.Close
frmFluidos.Height = 5070
'MSFlexGrid1.TextMatrix(0, 0) = "Codigo"
'MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
'MSFlexGrid1.TextMatrix(0, 2) = "Partida"
'MSFlexGrid1.TextMatrix(0, 3) = "Fluido"
'MSFlexGrid1.TextMatrix(0, 4) = "Var.Vol"
'MSFlexGrid1.TextMatrix(0, 5) = "Var.Tracc"
'MSFlexGrid1.TextMatrix(0, 6) = "Var.Elong"
'MSFlexGrid1.TextMatrix(0, 7) = "Var.Dur"
'MSFlexGrid1.TextMatrix(0, 8) = "Fecha"
'MSFlexGrid1.TextMatrix(0, 9) = "Dias Rep"
'MSFlexGrid1.TextMatrix(0, 10) = "Aprobado"
'MSFlexGrid1.TextMatrix(0, 11) = "Probeta"
'MSFlexGrid1.TextMatrix(0, 12) = "V.Masa"
    MSFlexGrid1.TextMatrix(0, 0) = "codigo"
    MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
    MSFlexGrid1.TextMatrix(0, 2) = "Partida"
    MSFlexGrid1.TextMatrix(0, 3) = "Fluido"
    MSFlexGrid1.TextMatrix(0, 4) = "Var.Vol1"
    MSFlexGrid1.TextMatrix(0, 5) = "Var.Vol2"
    MSFlexGrid1.TextMatrix(0, 6) = "Var.Vol3"
    MSFlexGrid1.TextMatrix(0, 7) = "Var.Vol.prom"
    MSFlexGrid1.TextMatrix(0, 8) = "Var.Tracc"
    MSFlexGrid1.TextMatrix(0, 9) = "Var.Elong"
    MSFlexGrid1.TextMatrix(0, 10) = "Var.dur"
    MSFlexGrid1.TextMatrix(0, 11) = "Fecha"
    MSFlexGrid1.TextMatrix(0, 12) = "Tiempo.rep"
    MSFlexGrid1.TextMatrix(0, 13) = "aprobado"
    MSFlexGrid1.TextMatrix(0, 14) = "Probeta"
    MSFlexGrid1.TextMatrix(0, 15) = "Var.Masa1"
    MSFlexGrid1.TextMatrix(0, 16) = "Var.Masa2"
    MSFlexGrid1.TextMatrix(0, 17) = "Var.Masa3"
    MSFlexGrid1.TextMatrix(0, 18) = "Var.Masa.prom"
Else
Exit Sub
End If
AutoGrid MSFlexGrid1
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
sfgh = MsgBox("Debe ingresar un lote para buscar.", vbCritical + vbOKOnly, "Error")
Text1.SetFocus
Exit Sub
End If

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT var_masa,var_masa1,var_masa2,var_vol1,var_vol2, probeta, codigo, N_FORMULA, PARTIDA, TIEMP_TEMP_OIL, var_vol, var_tracc, var_elong, var_dur, fecha_realizacion, tiempo_repeticion, aprovado FROM fluidos WHERE PARTIDA = '" & Text1.Text & "' ;")
If rs.RecordCount = 0 Then
    fgs = MsgBox("No existe el lote", vbCritical + vbOKOnly, "Error")
    db.Close
    Exit Sub
End If
rs.MoveLast
total = rs.RecordCount
MSFlexGrid1.Cols = 12
MSFlexGrid1.Rows = total + 1
rs.MoveFirst
For contador = 1 To total
'MSFlexGrid1.TextMatrix(contador, 0) = rs.Fields("codigo")
'MSFlexGrid1.TextMatrix(contador, 1) = rs.Fields("N_FORMULA")
'MSFlexGrid1.TextMatrix(contador, 2) = rs.Fields("partida")
'MSFlexGrid1.TextMatrix(contador, 3) = rs.Fields("TIEMP_TEMP_OIL")
'MSFlexGrid1.TextMatrix(contador, 4) = rs.Fields("VAR_VOL")
'MSFlexGrid1.TextMatrix(contador, 5) = rs.Fields("VAR_TRACC")
'MSFlexGrid1.TextMatrix(contador, 6) = rs.Fields("VAR_ELONG")
'MSFlexGrid1.TextMatrix(contador, 7) = rs.Fields("VAR_DUR")
'MSFlexGrid1.TextMatrix(contador, 8) = rs.Fields("FECHA_REALIZACION")
'MSFlexGrid1.TextMatrix(contador, 9) = rs.Fields("Tiempo_repeticion")
'MSFlexGrid1.TextMatrix(contador, 10) = rs.Fields("aprovado")
'MSFlexGrid1.TextMatrix(contador, 11) = rs.Fields("probeta") & ""
'MSFlexGrid1.TextMatrix(contador, 12) = rs.Fields("var_masa") & ""
    MSFlexGrid1.TextMatrix(contador, 0) = rs.Fields("codigo")
    MSFlexGrid1.TextMatrix(contador, 1) = rs.Fields("N_FORMULA")
    MSFlexGrid1.TextMatrix(contador, 2) = rs.Fields("partida")
    MSFlexGrid1.TextMatrix(contador, 3) = rs.Fields("TIEMP_TEMP_OIL")
    MSFlexGrid1.TextMatrix(contador, 4) = rs.Fields("VAR_VOL")
    MSFlexGrid1.TextMatrix(contador, 5) = rs.Fields("VAR_VOL1") & ""
    MSFlexGrid1.TextMatrix(contador, 6) = rs.Fields("VAR_VOL2") & ""
    MSFlexGrid1.TextMatrix(contador, 7) = (rs.Fields("var_vol") + rs.Fields("var_vol1") + rs.Fields("var_vol2")) / 3
    MSFlexGrid1.TextMatrix(contador, 8) = rs.Fields("VAR_TRACC")
    MSFlexGrid1.TextMatrix(contador, 9) = rs.Fields("VAR_ELONG")
    MSFlexGrid1.TextMatrix(contador, 10) = rs.Fields("VAR_DUR")
    MSFlexGrid1.TextMatrix(contador, 11) = rs.Fields("FECHA_REALIZACION")
    MSFlexGrid1.TextMatrix(contador, 12) = rs.Fields("Tiempo_repeticion")
    MSFlexGrid1.TextMatrix(contador, 13) = rs.Fields("aprovado")
    MSFlexGrid1.TextMatrix(contador, 14) = rs.Fields("probeta") & ""
    MSFlexGrid1.TextMatrix(contador, 15) = rs.Fields("var_masa") & ""
    MSFlexGrid1.TextMatrix(contador, 16) = rs.Fields("var_masa1") & ""
    MSFlexGrid1.TextMatrix(contador, 17) = rs.Fields("var_masa2") & ""
    MSFlexGrid1.TextMatrix(contador, 18) = (rs.Fields("var_masa") + rs.Fields("var_masa1") + rs.Fields("var_masa2")) / 3

rs.MoveNext
Next
frmFluidos.Height = 5070
'MSFlexGrid1.TextMatrix(0, 0) = "codigo"
'MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
'MSFlexGrid1.TextMatrix(0, 2) = "Partida"
'MSFlexGrid1.TextMatrix(0, 3) = "Fluido"
'MSFlexGrid1.TextMatrix(0, 4) = "Var.Vol"
'MSFlexGrid1.TextMatrix(0, 5) = "Var.Tracc"
'MSFlexGrid1.TextMatrix(0, 6) = "Var.Elong"
'MSFlexGrid1.TextMatrix(0, 7) = "Var.Dur"
'MSFlexGrid1.TextMatrix(0, 8) = "Fecha"
'MSFlexGrid1.TextMatrix(0, 9) = "Dias Rep"
'MSFlexGrid1.TextMatrix(0, 10) = "Aprobado"
'MSFlexGrid1.TextMatrix(0, 11) = "Probeta"
'MSFlexGrid1.TextMatrix(0, 12) = "V.Masa"
    MSFlexGrid1.TextMatrix(0, 0) = "codigo"
    MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
    MSFlexGrid1.TextMatrix(0, 2) = "Partida"
    MSFlexGrid1.TextMatrix(0, 3) = "Fluido"
    MSFlexGrid1.TextMatrix(0, 4) = "Var.Vol1"
    MSFlexGrid1.TextMatrix(0, 5) = "Var.Vol2"
    MSFlexGrid1.TextMatrix(0, 6) = "Var.Vol3"
    MSFlexGrid1.TextMatrix(0, 7) = "Var.Vol.prom"
    MSFlexGrid1.TextMatrix(0, 8) = "Var.Tracc"
    MSFlexGrid1.TextMatrix(0, 9) = "Var.Elong"
    MSFlexGrid1.TextMatrix(0, 10) = "Var.dur"
    MSFlexGrid1.TextMatrix(0, 11) = "Fecha"
    MSFlexGrid1.TextMatrix(0, 12) = "Tiempo.rep"
    MSFlexGrid1.TextMatrix(0, 13) = "aprobado"
    MSFlexGrid1.TextMatrix(0, 14) = "Probeta"
    MSFlexGrid1.TextMatrix(0, 15) = "Var.Masa1"
    MSFlexGrid1.TextMatrix(0, 16) = "Var.Masa2"
    MSFlexGrid1.TextMatrix(0, 17) = "Var.Masa3"
    MSFlexGrid1.TextMatrix(0, 18) = "Var.Masa.prom"
AutoGrid MSFlexGrid1
End Sub

Private Sub Command2_Click()
MSFlexGrid1.Clear
Combo1.Text = ""
frmFluidos.Height = 2085
End Sub

Private Sub Command3_Click()
Form1.Enabled = True
Form1.Visible = True
MSFlexGrid1.Clear
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
frmFluidos.Hide
End Sub

Private Sub Command4_Click()
frmfluido.Show
frmFluidos.Enabled = False
frmFluidos.Visible = False


frmfluido.Command4.Enabled = False
frmfluido.Command5.Enabled = True
frmfluido.Command6.Enabled = True
frmfluido.Command7.Enabled = True
frmfluido.Command8.Enabled = False
frmfluido.Command9.Enabled = False
frmfluido.Command10.Enabled = False
frmfluido.Command11.Enabled = False
frmfluido.Check1.Caption = "Sin seleccionar"
frmfluido.Check1.Value = 0
frmfluido.Option1.Value = True
frmfluido.Option2.Value = False
frmfluido.Combo1.Clear
frmfluido.Text10.Text = ""
frmfluido.List1.Clear
frmfluido.Text1.Text = ""
frmfluido.Label1.Caption = ""

frmfluido.Text12.Text = ""
frmfluido.Text13.Text = ""
frmfluido.Text14.Text = ""
frmfluido.Text15.Text = ""

frmfluido.Text2.Text = ""
frmfluido.Text3.Text = ""
frmfluido.Text4.Text = ""
frmfluido.Text5.Text = ""
frmfluido.Text6.Text = ""
frmfluido.Text7.Text = ""
frmfluido.Text8.Text = ""
frmfluido.Text9.Text = ""
frmfluido.Text11.Text = ""

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT N_FORMULA FROM Fluidos group by N_formula")
Set rs1 = db.OpenRecordset("SELECT referencia FROM ensayos where tipo = 'Envejecimiento'")
Set rs2 = db.OpenRecordset("SELECT codigo FROM fluidos order by codigo asc")
rs2.MoveLast
Codigo = CInt(rs2.Fields("codigo"))
Codigo = Codigo + 1


rs.MoveFirst
Do While rs.EOF = False
frmfluido.Combo1.AddItem (rs.Fields("n_formula"))
rs.MoveNext
Loop
rs1.MoveFirst
Do While rs1.EOF = False
frmfluido.List1.AddItem (rs1.Fields("referencia"))
rs1.MoveNext
Loop
frmfluido.Label1.Caption = Codigo
frmfluido.List1.Text = ""
db.Close
'FrmFluidoNuevo.Combo1.Clear
'FrmFluidoNuevo.List1.Clear
'FrmFluidoNuevo.Combo1.Text = ""
'FrmFluidoNuevo.List1.Text = ""
'FrmFluidoNuevo.Text1.Text = ""
'FrmFluidoNuevo.Text2.Text = ""
'FrmFluidoNuevo.Text3.Text = ""
'FrmFluidoNuevo.Text4.Text = ""
'FrmFluidoNuevo.Text5.Text = ""
'FrmFluidoNuevo.text6.Text = ""
'FrmFluidoNuevo.Label1.Caption = ""
'FrmFluidoNuevo.Show
'frmFluidos.Enabled = False
'frmFluidos.Visible = False
'Dim db As Database
'Dim rs As Recordset
'Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
'Set rs = db.OpenRecordset("SELECT N_FORMULA FROM Formbase")
'rs.MoveFirst
'a = rs.Fields("N_formula")
'Do While rs.EOF <> True
'rs.MoveNext
'Loop
'rs.MovePrevious
'fiLa = rs.RecordCount
'rs.MoveFirst
'For contador = 1 To fiLa
'b = rs.Fields("N_FORMULA").Value
'On Error Resume Next
'FrmFluidoNuevo.Combo1.AddItem (rs.Fields("N_FORMULA"))
'rs.MoveNext
'Next
'db.Close

'Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
'Set rs = db.OpenRecordset("Select TIEMP_TEMP_OIL From Fluidos GROUP BY TIEMP_TEMP_OIL")

'rs.MoveFirst
'Do While rs.EOF = False
'FrmFluidoNuevo.List1.AddItem (rs.Fields("tiemp_temp_oil"))
'rs.MoveNext
'Loop
'db.Close
End Sub

Private Sub Command5_Click()
Dim db As Database
Dim rs As Recordset
Dim mensaje As String
Dim fecHai As Date
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select N_Formula, tiemp_temp_oil, tiempo_repeticion, fecha_realizacion From fluidos")
mensaje1 = "Los ensayos de fluidos que se han vencido a la fecha son : "
mensaje = "Los ensayos de fluidos que se han vencido a la fecha son : "
rs.MoveFirst

Do While rs.EOF = False
    fecHai = rs.Fields("Fecha_realizacion")
    suma = (fecHai) + rs.Fields("tiempo_repeticion")
    If rs.Fields("tiempo_repeticion") = "0" Then
        rs.MoveNext
    Else
        If suma <= Date Then
            mensaje = mensaje & rs.Fields("N_formula") & " en " & rs.Fields("tiemp_temp_oil") & ", "
        End If
        rs.MoveNext
    End If
Loop
If mensaje = mensaje1 Then
mensaje = mensaje & " Sin ensayos Vencidos"
End If
fghytws = MsgBox(mensaje, vbInformation + vbOKOnly, "Informe de ensayos")
End Sub

Private Sub Form_Load()
Combo2.Enabled = False
Text1.Enabled = False
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
Command1.Enabled = False
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select N_FORMULA From Fluidos GROUP BY N_FORMULA")

rs.MoveFirst
Do While rs.EOF = False
Combo1.AddItem (rs.Fields("N_FORMULA"))
rs.MoveNext
Loop
db.Close





End Sub

Private Sub Option1_Click()
Combo2.Text = ""
Text1.Text = ""
Combo1.Enabled = True
Combo2.Enabled = False
Text1.Enabled = False
Command1.Enabled = False
Combo1.Clear
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select N_FORMULA From Fluidos GROUP BY N_FORMULA")

rs.MoveFirst
Do While rs.EOF = False
Combo1.AddItem (rs.Fields("N_FORMULA"))
rs.MoveNext
Loop
db.Close
End Sub

Private Sub Option2_Click()
Combo2.Clear

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select TIEMP_TEMP_OIL From Fluidos GROUP BY TIEMP_TEMP_OIL")

rs.MoveFirst
Do While rs.EOF = False
Combo2.AddItem (rs.Fields("tiemp_temp_oil"))
rs.MoveNext
Loop
db.Close
Combo1.Text = ""
Text1.Text = ""
Combo1.Enabled = False
Combo2.Enabled = True
Text1.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Option3_Click()
Combo1.Text = ""
Combo2.Text = ""
Combo1.Enabled = False
Combo2.Enabled = False
Text1.Enabled = True
Command1.Enabled = True
End Sub
