VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBuscaTraccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Valores de Tracción"
   ClientHeight    =   6495
   ClientLeft      =   3540
   ClientTop       =   1260
   ClientWidth     =   12720
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Compresion"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compara"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   4560
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Compuesto y ensayo"
      Height          =   195
      Left            =   2640
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   16
      FixedCols       =   3
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Codigo de ensayo"
      Height          =   195
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Por:"
      Height          =   1935
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option4 
         Caption         =   "Compuesto"
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Compuesto y Partida"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Ensayo"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Partida"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "frmBuscaTraccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset

MSFlexGrid1.TextMatrix(0, 0) = "Codigo"
MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
MSFlexGrid1.TextMatrix(0, 2) = "Partida"
MSFlexGrid1.TextMatrix(0, 3) = "Estado"
MSFlexGrid1.TextMatrix(0, 4) = "Referencia"
MSFlexGrid1.TextMatrix(0, 5) = "Tracción1"
MSFlexGrid1.TextMatrix(0, 6) = "Tracción2"
MSFlexGrid1.TextMatrix(0, 7) = "Tracción3"
MSFlexGrid1.TextMatrix(0, 8) = "Tracc.Prom"
MSFlexGrid1.TextMatrix(0, 9) = "Elongación1" ' 6
MSFlexGrid1.TextMatrix(0, 10) = "Elongación2"
MSFlexGrid1.TextMatrix(0, 11) = "Elongación3"
MSFlexGrid1.TextMatrix(0, 12) = "Elong.Prom"
MSFlexGrid1.TextMatrix(0, 13) = "Dureza" '7
MSFlexGrid1.TextMatrix(0, 14) = "Probeta" '8
MSFlexGrid1.TextMatrix(0, 15) = "Observacion"



Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")

If Option1.Value = True Then
    Set rs = db.OpenRecordset("Select observacion,probeta, CODIGO_ENSAYO, COMPUESTO, PARTIDA, ESTADO_ENSAYO, REFERENCIA, DUREZA, TRACCION, traccion1, traccion2, ELONGACION, elongacion1, elongacion2 from traccion where compuesto = '" & Text1.Text & "' AND partida = '" & Text2.Text & "' order by CODIGO_ENSAYO desc;")
    If rs.RecordCount > 0 Then
    rs.MoveLast
    End If
    If rs.RecordCount > 0 Then
        MSFlexGrid1.Rows = rs.RecordCount + 1
        frmBuscaTraccion.Height = 6900
    Else
    asdfadf = MsgBox("No se han encontrado registros", vbCritical + vbOKOnly, "Error")
    Exit Sub
    End If
    rs.MoveFirst
    For bucleflex = 1 To rs.RecordCount
        MSFlexGrid1.TextMatrix(bucleflex, 0) = rs.Fields("codigo_ensayo")
        MSFlexGrid1.TextMatrix(bucleflex, 1) = rs.Fields("Compuesto")
        MSFlexGrid1.TextMatrix(bucleflex, 2) = rs.Fields("partida")
        MSFlexGrid1.TextMatrix(bucleflex, 3) = rs.Fields("estado_ensayo")
        MSFlexGrid1.TextMatrix(bucleflex, 4) = rs.Fields("referencia")
        MSFlexGrid1.TextMatrix(bucleflex, 5) = rs.Fields("traccion")
        MSFlexGrid1.TextMatrix(bucleflex, 6) = rs.Fields("traccion1") & ""
        MSFlexGrid1.TextMatrix(bucleflex, 7) = rs.Fields("traccion2") & ""
        If IsNull(rs.Fields("traccion1")) Then
            MSFlexGrid1.TextMatrix(bucleflex, 8) = rs.Fields("traccion")
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 8) = Format((CDbl(rs.Fields("traccion")) + CDbl(rs.Fields("traccion1")) + CDbl(rs.Fields("traccion2"))) / 3, "0.00")
        End If
        
        MSFlexGrid1.TextMatrix(bucleflex, 9) = rs.Fields("elongacion")
        MSFlexGrid1.TextMatrix(bucleflex, 10) = rs.Fields("elongacion1") & ""
        MSFlexGrid1.TextMatrix(bucleflex, 11) = rs.Fields("elongacion2") & ""
        If IsNull(rs.Fields("elongacion1")) Then
            MSFlexGrid1.TextMatrix(bucleflex, 12) = rs.Fields("elongacion")
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 12) = (CDbl(rs.Fields("elongacion")) + CDbl(rs.Fields("elongacion1")) + CDbl(rs.Fields("elongacion2"))) / 3
        End If
        MSFlexGrid1.TextMatrix(bucleflex, 13) = rs.Fields("dureza")
        MSFlexGrid1.TextMatrix(bucleflex, 14) = rs.Fields("probeta") & ""
        If rs.Fields("Observacion") = "" Then
            MSFlexGrid1.TextMatrix(bucleflex, 15) = "No"
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 15) = "Si"
        End If
        rs.MoveNext
    Next 'bucleflex
End If 'option1.value
    
If Option2.Value = True Then
'MsgBox ("Select CODIGO_ENSAYO, COMPUESTO, PARTIDA, ESTADO_ENSAYO, REFERENCIA, DUREZA, TRACCION, ELONGACION from traccion where codigo_ensayo = '" & Text1.Text & "' ;")
    Set rs = db.OpenRecordset("Select probeta, CODIGO_ENSAYO, COMPUESTO, PARTIDA, ESTADO_ENSAYO, REFERENCIA, DUREZA, TRACCION, traccion1, traccion2, ELONGACION, elongacion1, elongacion2, observacion from traccion where codigo_ensayo = " & Text1.Text & " order by codigo_ensayo desc")
    
    If rs.RecordCount > 0 Then
        rs.MoveLast
    End If
    If rs.RecordCount > 0 Then
        MSFlexGrid1.Rows = rs.RecordCount + 1
        frmBuscaTraccion.Height = 6900
    Else
    asdfadf = MsgBox("No se han encontrado registros", vbCritical + vbOKOnly, "Error")
    Exit Sub
    End If
    rs.MoveFirst
        MSFlexGrid1.TextMatrix(1, 0) = rs.Fields("codigo_ensayo")
        MSFlexGrid1.TextMatrix(1, 1) = rs.Fields("Compuesto")
        MSFlexGrid1.TextMatrix(1, 2) = rs.Fields("partida")
        MSFlexGrid1.TextMatrix(1, 3) = rs.Fields("estado_ensayo")
        MSFlexGrid1.TextMatrix(1, 4) = rs.Fields("referencia")
        MSFlexGrid1.TextMatrix(1, 5) = rs.Fields("traccion")
        MSFlexGrid1.TextMatrix(1, 6) = rs.Fields("traccion1") & ""
        MSFlexGrid1.TextMatrix(1, 7) = rs.Fields("traccion2") & ""
        If rs.Fields("traccion1") = 0 Then
            MSFlexGrid1.TextMatrix(1, 8) = rs.Fields("traccion")
        Else
            MSFlexGrid1.TextMatrix(1, 8) = Format((CDbl(rs.Fields("traccion")) + CDbl(rs.Fields("traccion1")) + CDbl(rs.Fields("traccion2"))) / 3, "0.00")
        End If
        
        MSFlexGrid1.TextMatrix(1, 9) = rs.Fields("elongacion")
        MSFlexGrid1.TextMatrix(1, 10) = rs.Fields("elongacion1") & ""
        MSFlexGrid1.TextMatrix(1, 11) = rs.Fields("elongacion2") & ""
        If rs.Fields("elongacion1") = 0 Then
            MSFlexGrid1.TextMatrix(1, 12) = rs.Fields("elongacion")
        Else
            MSFlexGrid1.TextMatrix(1, 12) = (CDbl(rs.Fields("elongacion")) + CDbl(rs.Fields("elongacion1")) + CDbl(rs.Fields("elongacion2"))) / 3 & ""
        End If
        MSFlexGrid1.TextMatrix(1, 13) = rs.Fields("dureza")
        MSFlexGrid1.TextMatrix(1, 14) = rs.Fields("probeta") & ""
        If rs.Fields("Observacion") = "" Then
            MSFlexGrid1.TextMatrix(1, 15) = "No"
        Else
            MSFlexGrid1.TextMatrix(1, 15) = "Si"
        End If
        
End If 'option1.value
If Option3.Value = True Then
    Set rs = db.OpenRecordset("Select probeta, CODIGO_ENSAYO, COMPUESTO, PARTIDA, ESTADO_ENSAYO, REFERENCIA, DUREZA, TRACCION, traccion1, traccion2, ELONGACION, elongacion1, elongacion2, observacion from traccion where compuesto = '" & Text1.Text & "' AND referencia = '" & Combo1 & "' order by codigo_ensayo desc;")
    If rs.RecordCount > 0 Then
        rs.MoveLast
    End If
    
    
    If rs.RecordCount > 0 Then
        MSFlexGrid1.Rows = rs.RecordCount + 1
        frmBuscaTraccion.Height = 6900
    Else
    asdfadf = MsgBox("No se han encontrado registros", vbCritical + vbOKOnly, "Error")
    Exit Sub
    End If
    rs.MoveFirst
    For bucleflex = 1 To rs.RecordCount
            MSFlexGrid1.TextMatrix(bucleflex, 0) = rs.Fields("codigo_ensayo")
        MSFlexGrid1.TextMatrix(bucleflex, 1) = rs.Fields("Compuesto")
        MSFlexGrid1.TextMatrix(bucleflex, 2) = rs.Fields("partida")
        MSFlexGrid1.TextMatrix(bucleflex, 3) = rs.Fields("estado_ensayo")
        MSFlexGrid1.TextMatrix(bucleflex, 4) = rs.Fields("referencia")
        MSFlexGrid1.TextMatrix(bucleflex, 5) = rs.Fields("traccion")
        MSFlexGrid1.TextMatrix(bucleflex, 6) = rs.Fields("traccion1") & ""
        MSFlexGrid1.TextMatrix(bucleflex, 7) = rs.Fields("traccion2") & ""
        If rs.Fields("traccion1") = 0 Then
            MSFlexGrid1.TextMatrix(bucleflex, 8) = rs.Fields("traccion")
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 8) = Format((rs.Fields("traccion") + rs.Fields("traccion1") + rs.Fields("traccion2")) / 3, "0.00")
        End If
        
        MSFlexGrid1.TextMatrix(bucleflex, 9) = rs.Fields("elongacion")
        MSFlexGrid1.TextMatrix(bucleflex, 10) = rs.Fields("elongacion1") & ""
        MSFlexGrid1.TextMatrix(bucleflex, 11) = rs.Fields("elongacion2") & ""
        If rs.Fields("elongacion1") = 0 Then
            MSFlexGrid1.TextMatrix(bucleflex, 12) = rs.Fields("elongacion")
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 12) = (rs.Fields("elongacion") + rs.Fields("elongacion1") + rs.Fields("elongacion2")) / 3 & ""
        End If
        MSFlexGrid1.TextMatrix(bucleflex, 13) = rs.Fields("dureza")
        MSFlexGrid1.TextMatrix(bucleflex, 14) = rs.Fields("probeta") & ""
        If rs.Fields("Observacion") = "" Then
            MSFlexGrid1.TextMatrix(bucleflex, 15) = "No"
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 15) = "Si"
        End If
        rs.MoveNext
    Next 'bucleflex
End If
If Option4.Value = True Then
    Set rs = db.OpenRecordset("Select probeta, CODIGO_ENSAYO, COMPUESTO, PARTIDA, ESTADO_ENSAYO, REFERENCIA, DUREZA, TRACCION, traccion1, traccion2, ELONGACION, elongacion1, elongacion2, observacion from traccion where compuesto = '" & Text1.Text & "' order by codigo_ensayo desc")
    If rs.RecordCount > 0 Then
    rs.MoveLast
    End If
    If rs.RecordCount > 0 Then
        MSFlexGrid1.Rows = rs.RecordCount + 1
        frmBuscaTraccion.Height = 6900
    Else
    asdfadf = MsgBox("No se han encontrado registros", vbCritical + vbOKOnly, "Error")
    Exit Sub
    End If
    rs.MoveFirst
    For bucleflex = 1 To rs.RecordCount
            MSFlexGrid1.TextMatrix(bucleflex, 0) = rs.Fields("codigo_ensayo")
        MSFlexGrid1.TextMatrix(bucleflex, 1) = rs.Fields("Compuesto")
        MSFlexGrid1.TextMatrix(bucleflex, 2) = rs.Fields("partida")
        MSFlexGrid1.TextMatrix(bucleflex, 3) = rs.Fields("estado_ensayo")
        MSFlexGrid1.TextMatrix(bucleflex, 4) = rs.Fields("referencia")
        MSFlexGrid1.TextMatrix(bucleflex, 5) = rs.Fields("traccion")
        MSFlexGrid1.TextMatrix(bucleflex, 6) = rs.Fields("traccion1") & ""
        MSFlexGrid1.TextMatrix(bucleflex, 7) = rs.Fields("traccion2") & ""
        If rs.Fields("traccion1") = 0 Then
            MSFlexGrid1.TextMatrix(bucleflex, 8) = rs.Fields("traccion")
        ElseIf IsNull(rs.Fields("traccion1")) Then
            MSFlexGrid1.TextMatrix(bucleflex, 8) = rs.Fields("traccion")
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 8) = Format((CDbl(rs.Fields("traccion")) + CDbl(rs.Fields("traccion1")) + CDbl(rs.Fields("traccion2"))) / 3, "0.00")
        End If
        
        MSFlexGrid1.TextMatrix(bucleflex, 9) = rs.Fields("elongacion")
        MSFlexGrid1.TextMatrix(bucleflex, 10) = rs.Fields("elongacion1") & ""
        MSFlexGrid1.TextMatrix(bucleflex, 11) = rs.Fields("elongacion2") & ""
        If rs.Fields("elongacion1") = 0 Then
            MSFlexGrid1.TextMatrix(bucleflex, 12) = rs.Fields("elongacion")
        ElseIf IsNull(rs.Fields("elongacion1")) Then
            MSFlexGrid1.TextMatrix(bucleflex, 12) = rs.Fields("elongacion")
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 12) = Format((CDbl(rs.Fields("elongacion")) + CDbl(rs.Fields("elongacion1")) + CDbl(rs.Fields("elongacion2"))) / 3, "0.00")
        End If
        MSFlexGrid1.TextMatrix(bucleflex, 13) = rs.Fields("dureza")
        MSFlexGrid1.TextMatrix(bucleflex, 14) = rs.Fields("probeta") & ""
        If rs.Fields("Observacion") = "" Then
            MSFlexGrid1.TextMatrix(bucleflex, 15) = "No"
        ElseIf IsNull(rs.Fields("observacion")) Then
            MSFlexGrid1.TextMatrix(bucleflex, 15) = "No"
        Else
            MSFlexGrid1.TextMatrix(bucleflex, 15) = "Si"
        End If
        
        rs.MoveNext
    Next 'bucleflex
End If 'option4.value

AutoGrid MSFlexGrid1
db.Close
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
frmBuscaTraccion.Hide
End Sub

Private Sub Command3_Click()
codigooriginal = InputBox("Ingrese el código de ensayo original a comparar", "Original")
If codigooriginal = "" Then
    Exit Sub
End If
codigoenvejecido = InputBox("ingrese el código de ensayo envejecido a comparar", "Envejecido")
If codigoenvejecido = "" Then
    Exit Sub
End If

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
'Set rs = db.OpenRecordset("Select codigo_ensayo, compuesto, partida, estado_ensayo, referencia, dureza, traccion, elongacion From traccion where codigo_ensayo = '" & codigooriginal & " or codigo_ensayo = " & codigoenvejecido & "'")
Set rs = db.OpenRecordset("Select * From traccion where codigo_ensayo = " & codigooriginal)
Set rs1 = db.OpenRecordset("Select * From traccion where codigo_ensayo = " & codigoenvejecido)

If rs.RecordCount = 0 Then
    sdf = MsgBox("No se han encontrado uno o más registros", vbCritical + vbOKOnly, "error")
    Exit Sub
End If
If rs1.RecordCount = 0 Then
    sdf = MsgBox("No se han encontrado uno o más registros", vbCritical + vbOKOnly, "error")
    Exit Sub
End If

compuesto1 = rs.Fields("Compuesto")
partida1 = rs.Fields("partida")
estado1 = rs.Fields("estado_ensayo")
referencia1 = rs.Fields("referencia")
dureza1 = rs.Fields("dureza")
traccion1a = rs.Fields("traccion")
traccion1b = rs.Fields("traccion1")
traccion1c = rs.Fields("traccion2")
elongacion1a = rs.Fields("elongacion")
elongacion1b = rs.Fields("elongacion1")
elongacion1c = rs.Fields("elongacion2")

'rs.MoveNext

compuesto2 = rs1.Fields("Compuesto")
partida2 = rs1.Fields("partida")
estado2 = rs1.Fields("estado_ensayo")
referencia2 = rs1.Fields("referencia")
dureza2 = rs1.Fields("dureza")
traccion2a = rs1.Fields("traccion")
traccion2b = rs1.Fields("traccion1")
traccion2c = rs1.Fields("traccion2")
elongacion2a = rs1.Fields("elongacion")
elongacion2b = rs1.Fields("elongacion1")
elongacion2c = rs1.Fields("elongacion2")
    
vardureza = dureza2 - dureza1
If dureza2 > dureza1 Then
    vardureza = "+" & vardureza
End If
If IsNull(rs.Fields("traccion1")) Or IsNull(rs1.Fields("traccion1")) Then
    vartraccion = (traccion2a * 100 / traccion1a) - 100
Else
    vartraccion = (((CDbl(traccion2a) + CDbl(traccion2b) + CDbl(traccion2c)) / 3) * 100 / ((CDbl(traccion1a) + CDbl(traccion1b) + CDbl(traccion1c)) / 3)) - 100
End If

If vartraccion > 0 Then
vartraccion = "+" & vartraccion
End If



If IsNull(rs1.Fields("elongacion1")) Or IsNull(rs.Fields("elongacion1")) Then
    varelongacion = (elongacion2a * 100 / elongacion1a) - 100
Else
    varelongacion = (((CDbl(elongacion2a) + CDbl(elongacion2b) + CDbl(elongacion2c)) / 3) * 100 / ((CDbl(elongacion1a) + CDbl(elongacion1b) + CDbl(elongacion1c)) / 3)) - 100
End If
If varelongacion > 0 Then

varelongacion = "+" & varelongacion
End If



rta = MsgBox("Comparación entre " & compuesto1 & " - " & partida1 & " - " & estado1 & " - " & referencia1 & " y " & compuesto2 & " - " & partida2 & " - " & estado2 & " - " & referencia2 & " Var. Dureza = " & vardureza & " Var. Tracción = " & vartraccion & " Var. Elongación = " & varelongacion, vbInformation + vbOKOnly, "Informe")



End Sub

Private Sub Command4_Click()
frmBuscaTraccion.Enabled = False
frmBuscaTraccion.Visible = False
If Text1.Text <> "" And Text2.Text <> "" Then
Dim db As Database
Dim rs As Recordset
'esto es para que busque directamente la compresion
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select probeta, compresion, codigo_ensayo, compuesto, partida, tiempo_temperatura, compresion_porc from compresion where compuesto ='" & Text1.Text & "' and partida ='" & Text2.Text & "'")
    If rs.RecordCount = 0 Then
        frmBuscaComp.Show
        frmBuscaTraccion.Enabled = False
        frmBuscaComp.Command2.Visible = False
        frmBuscaComp.Command3.Visible = True
        frmBuscaComp.Height = 2745
        frmBuscaComp.Text1.Text = ""
        frmBuscaComp.Text2.Text = ""
        frmBuscaComp.Text3.Text = ""
    Else
        frmBuscaComp.Height = 4485
        frmBuscaComp.MSFlexGrid1.TextMatrix(0, 0) = "Compuesto"
        frmBuscaComp.MSFlexGrid1.TextMatrix(0, 1) = "Partida"
        frmBuscaComp.MSFlexGrid1.TextMatrix(0, 2) = "Codigo"
        frmBuscaComp.MSFlexGrid1.TextMatrix(0, 3) = "Ensayo"
        frmBuscaComp.MSFlexGrid1.TextMatrix(0, 4) = "% de deformación"
        frmBuscaComp.MSFlexGrid1.TextMatrix(0, 5) = "Probeta"
        frmBuscaComp.MSFlexGrid1.TextMatrix(0, 6) = "Compresion"
        
        rs.MoveLast
        ultimo = rs.RecordCount
        rs.MoveFirst
        frmBuscaComp.MSFlexGrid1.Rows = ultimo + 1
        For indice = 1 To ultimo
            frmBuscaComp.MSFlexGrid1.TextMatrix(indice, 0) = rs.Fields("compuesto")
            frmBuscaComp.MSFlexGrid1.TextMatrix(indice, 1) = rs.Fields("Partida")
            frmBuscaComp.MSFlexGrid1.TextMatrix(indice, 2) = rs.Fields("Codigo_ensayo")
            frmBuscaComp.MSFlexGrid1.TextMatrix(indice, 3) = rs.Fields("tiempo_temperatura")
            frmBuscaComp.MSFlexGrid1.TextMatrix(indice, 4) = rs.Fields("compresion_porc")
            frmBuscaComp.MSFlexGrid1.TextMatrix(indice, 5) = rs.Fields("probeta") & ""
            frmBuscaComp.MSFlexGrid1.TextMatrix(indice, 6) = rs.Fields("Compresion") & ""
            rs.MoveNext
        Next
        rs.MovePrevious
        frmBuscaComp.Command3.Visible = True
        frmBuscaComp.Command2.Visible = True
        frmBuscaComp.Text4.Visible = True
        frmBuscaComp.Text5.Visible = True
        frmBuscaComp.Text6.Visible = True
        frmBuscaComp.Text4.Text = rs.Fields("compuesto")
        frmBuscaComp.Text5.Text = rs.Fields("Partida")
        frmBuscaComp.Text6 = rs.Fields("codigo_ensayo")
        frmBuscaComp.Command1.Enabled = False
        frmBuscaComp.Show
        db.Close
    End If
Else
        frmBuscaComp.Show
        frmBuscaTraccion.Enabled = False
        frmBuscaComp.Command2.Visible = False
        frmBuscaComp.Command3.Visible = True
        frmBuscaComp.Height = 2745
        frmBuscaComp.Text1.Text = ""
        frmBuscaComp.Text2.Text = ""
        frmBuscaComp.Text3.Text = ""
End If
AutoGrid frmBuscaComp.MSFlexGrid1
End Sub

Private Sub Command5_Click()
frmBuscaTraccion.Hide
End Sub

Private Sub Command6_Click()
frmBuscaTraccion.Command5.Visible = False
frmBuscaTraccion.Command6.Visible = False
frmBuscaTraccion.Command2.Visible = False
frmModPartidasNuevas.Enabled = True
frmBuscaTraccion.Hide

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Command4.Enabled = True
    frmBuscaTraccion.Height = 5565
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Text2.Visible = True
    Label2.Visible = True
    Label1.Caption = "Compuesto"
    Label2.Caption = "Partida"
    Label3.Visible = False
    Combo1.Visible = False
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Clear
    Combo1.Text = ""
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    Command4.Enabled = False
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Clear
    Combo1.Text = ""
    frmBuscaTraccion.Height = 5565
    Option1.Value = False
    Option4.Value = False
    Label3.Visible = False
    Combo1.Visible = False
    Option3.Value = False
    Text2.Visible = False
    Label2.Visible = False
    Label1.Caption = "Codigo"
    Text1.Text = ""
    
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    Command4.Enabled = False
    Text1.Text = ""
    Text2.Text = ""
    frmBuscaTraccion.Height = 5565
    Combo1.Clear
    Option1.Value = False
    Option4.Value = False
    Label1.Caption = "Compuesto"
    Option2.Value = False
    Label2.Visible = False
    Text2.Visible = False
    Label3.Visible = True
    Combo1.Visible = True
    Dim db As Database
    Dim rs As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
    Set rs = db.OpenRecordset("Select referencia from traccion group by referencia")
    rs.MoveFirst
    Do Until rs.EOF = True
        Combo1.AddItem (rs.Fields("referencia"))
        rs.MoveNext
    Loop
    db.Close
End If

End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Command4.Enabled = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Text2.Visible = False
Combo1.Visible = False
Label2.Caption = ""
Label3.Caption = ""
End If



End Sub

Private Sub Text1_Change()
Text2.Text = ""
End Sub
