VERSION 5.00
Begin VB.Form frmBusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   7680
   ClientLeft      =   750
   ClientTop       =   570
   ClientWidth     =   13695
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command14 
      Caption         =   "Dar de alta "
      Enabled         =   0   'False
      Height          =   255
      Left            =   9240
      TabIndex        =   48
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Marcar como Leido"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10680
      TabIndex        =   47
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Marcar como no leido"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10680
      TabIndex        =   46
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Laboratorio"
      Height          =   1575
      Left            =   10320
      TabIndex        =   45
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Marcar como no leido"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   43
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Marcar como Leido"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   42
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   5880
      TabIndex        =   41
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox text12 
      Height          =   315
      Left            =   2400
      TabIndex        =   40
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ComboBox Text9 
      Height          =   315
      Left            =   2400
      TabIndex        =   39
      Text            =   "Combo1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9840
      TabIndex        =   38
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9480
      TabIndex        =   37
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9840
      TabIndex        =   36
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9840
      TabIndex        =   35
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      Height          =   495
      Left            =   4920
      TabIndex        =   34
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   495
      Left            =   4200
      TabIndex        =   33
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   255
      Left            =   7560
      TabIndex        =   32
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Text            =   "13"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Text            =   "11"
      Top             =   1560
      Width           =   5055
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "10"
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Text            =   "8"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmBusca.frx":0000
      Top             =   3360
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   4920
      Width           =   6975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   975
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmBusca.frx":0006
      Top             =   6240
      Width           =   7335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar otro"
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   975
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmBusca.frx":0008
      Top             =   2280
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingeniería"
      Height          =   1575
      Left            =   7800
      TabIndex        =   44
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Caracteres Restantes"
      Height          =   255
      Left            =   9840
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label21"
      Height          =   255
      Left            =   12000
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Esp. Plano - Norma"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Referencia de pieza"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Responsable"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de solicitud"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo de solicitud"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Solic. de documentacion"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Fecha de solicitud"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Compuesto Recomendado"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label19 
      Caption         =   "Fecha de Recom."
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "Responsable"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Caracteres Restantes"
      Height          =   255
      Left            =   9840
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   255
      Left            =   12000
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Caracteres Restantes"
      Height          =   255
      Left            =   9840
      TabIndex        =   14
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label23 
      Caption         =   "Label21"
      Height          =   255
      Left            =   12000
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dbbusc As Database
Public rsbusc As Recordset
Public query
Public posicion
Public total
Public Flag As Boolean
Public Campo As String
Public caracteres As Integer

Private Sub Command1_Click()
Flag = False
Command14.Enabled = False
Form1.Enabled = True
Form1.Visible = True
Text9.Clear
On Error Resume Next
dbbusc.Close
frmBusca.Hide
End Sub



Private Sub Command10_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Laboratorio", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'labkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If



frmBusca.Enabled = False
Set dbbusc = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rsbusc = dbbusc.OpenRecordset("Select revisado_lab,revisado_ing,solicitud_documentacion from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)
'caracteres = CInt(Label21.Caption) - 1
Campo = "solicitud_documentacion"
frmModificaCotizaciones.Show
frmModificaCotizaciones.Label1.Caption = Label21.Caption
Command9.Enabled = False
Command10.Enabled = False
Command12.Enabled = False
Command15.Enabled = False
frmModificaCotizaciones.Text1.Text = ""
Label2.Caption = ""
frmModificaCotizaciones.Label2 = Text1.Text
frmModificaCotizaciones.Text1.SetFocus

End Sub

Private Sub Command11_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Laboratorio", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'labkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select revisado_lab from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)

rs.Edit
rs.Fields("revisado_lab") = False
rs.Update
db.Close

sdfsdfsf = MsgBox("Se ha marcado la recomendación como no leido", vbInformation + vbOKOnly, "Marcado como NO leido")
End Sub

Private Sub Command12_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Laboratorio", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'labkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If





frmBusca.Enabled = False
Set dbbusc = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rsbusc = dbbusc.OpenRecordset("Select revisado_lab,revisado_ing,comp_recomendado from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)
'caracteres = 255 - Len(Text3.Text) - 1
Campo = "comp_recomendado"
frmModificaCotizaciones.Show
'frmModificaCotizaciones.Label1.Caption = caracteres
Command9.Enabled = False
Command10.Enabled = False
Command12.Enabled = False
Command15.Enabled = False
frmModificaCotizaciones.Text1.Text = ""
Label2.Caption = ""
frmModificaCotizaciones.Label2 = Text3.Text
frmModificaCotizaciones.Text1.SetFocus

End Sub

Private Sub Command13_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Laboratorio", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'labkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select revisado_lab from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)

rs.Edit
rs.Fields("revisado_lab") = True
rs.Update
db.Close

sdfsdfsf = MsgBox("Se ha marcado la recomendación como leido", vbInformation + vbOKOnly, "Marcado como leido")
End Sub

Private Sub Command14_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Ingeniería", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'ingkey'")

If UCase(rs.Fields("dato")) = UCase(contra) Then
Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If
frmCotizAlta.Text1.Text = Text10.Text
frmCotizAlta.Text2.Text = Text12.Text
frmCotizAlta.Text3.Text = Text11.Text
frmCotizAlta.Text4.Text = Text3.Text
Me.Enabled = False
db.Close
frmCotizAlta.Show
End Sub

Private Sub Command15_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Laboratorio", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'labkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If




frmBusca.Enabled = False
Set dbbusc = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rsbusc = dbbusc.OpenRecordset("Select revisado_lab,revisado_ing,observ_lab from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)
'caracteres = CInt(Label23.Caption) - 1
Campo = "observ_lab"
frmModificaCotizaciones.Show
'frmModificaCotizaciones.Label1.Caption = caracteres
Command9.Enabled = False
Command10.Enabled = False
Command12.Enabled = False
Command15.Enabled = False
frmModificaCotizaciones.Text1.Text = ""
Label2.Caption = ""
frmModificaCotizaciones.Label2 = Text6.Text
frmModificaCotizaciones.Text1.SetFocus
End Sub

Private Sub Command2_Click()
rsbusc.MovePrevious
Text5.Text = rsbusc.Fields("codigo_recomendacion")
Text8.Text = rsbusc.Fields("fecha_solicitud")
Text9.Text = rsbusc.Fields("respons_solici")
Text10.Text = rsbusc.Fields("referencia")
Text11.Text = rsbusc.Fields("esp_plano_norma")
Text12.Text = rsbusc.Fields("cliente")
Text7.Text = rsbusc.Fields("observaciones")
Text1.Text = rsbusc.Fields("solicitud_documentacion")
Text2.Text = rsbusc.Fields("fecha_sol_docu")
Text3.Text = rsbusc.Fields("comp_recomendado")
Text13.Text = rsbusc.Fields("fecha_recomend")
Text4.Text = rsbusc.Fields("responsable_recomend")
Text6.Text = rsbusc.Fields("observ_lab")
posicion = rsbusc.AbsolutePosition
If posicion = 0 Then
    Command2.Enabled = False
End If
If posición <> total Then
    Command5.Enabled = True
End If
End Sub

Private Sub Command3_Click()

'If Label4.Caption < 0 Then
'msdfsdf = MsgBox("La cantidad de caracteres es mayor a la posible", vbCritical + vbOKOnly, "Error")
'Text7.SetFocus
'Exit Sub
'End If
'If Label21.Caption < 0 Then
'msdfsdf = MsgBox("La cantidad de caracteres es mayor a la posible", vbCritical + vbOKOnly, "Error")
'Text1.SetFocus
'Exit Sub
'End If
'If Label23.Caption < 0 Then
'msdfsdf = MsgBox("La cantidad de caracteres es mayor a la posible", vbCritical + vbOKOnly, "Error")
'Text6.SetFocus
'Exit Sub
'End If
strQ = "Select Codigo_recomendacion, fecha_solicitud, respons_solici, referencia, esp_plano_norma, cliente, observaciones, solicitud_documentacion, fecha_sol_docu, comp_recomendado, fecha_recomend, responsable_recomend, observ_lab  From compuestos_para_cotizacion where"
If Text5.Text <> "" Then
    strQry = "codigo_recomendacion = " & Text5.Text    '"'" & Text5.Text & "'"
Else
    strQry = "codigo_recomendacion IS NOT NULL"
    strQry1 = "codigo_recomendacion IS NOT NULL"
End If
If Text8.Text <> "" Then
strQry = strQry & " AND " & "fecha_solicitud = " & "'" & Text8.Text & "'"
End If
If Text9.Text <> "" Then
strQry = strQry & " AND " & "respons_solici = " & "'" & Text9.Text & "'"
End If
If Text10.Text <> "" Then
strQry = strQry & " AND " & "referencia = " & "'" & Text10.Text & "'"
End If
If Text11.Text <> "" Then
strQry = strQry & " AND " & "esp_plano_norma = " & "'" & Text11.Text & "'"
End If
If Text12.Text <> "" Then
strQry = strQry & " AND " & "cliente = " & "'" & Text12.Text & "'"
End If
If Text7.Text <> "" Then
strQry = strQry & " AND " & "observaciones = " & "'" & Text7.Text & "'"
End If
If Text1.Text <> "" Then
strQry = strQry & " AND " & "solicitud_documentacion = " & "'" & Text1.Text & "'"
End If
If Text2.Text <> "" Then
strQry = strQry & " AND " & "fecha_sol_docu = " & "'" & Text2.Text & "'"
End If
If Text3.Text <> "" Then
strQry = strQry & " AND " & "comp_recomendado = " & "'" & Text3.Text & "'"
End If
If Text13.Text <> "" Then
strQry = strQry & " AND " & "fecha_recomend = " & "'" & Text13.Text & "'"
End If
If Text4.Text <> "" Then
strQry = strQry & " AND " & "responsable_recomend = " & "'" & Text4.Text & "'"
End If
If Text6.Text <> "" Then
strQry = strQry & " AND " & "observ_lab = " & "'" & Text6.Text & "'"
End If
If strQry = strQ Then
    gsdgsdfg = MsgBox("No ha ingresado ningun valor para realizar la busqueda", vbCritical + vbOKOnly, "Error")
    Text5.SetFocus
    Exit Sub
End If
If strQry = strQry1 Then
    gsdgsdfg = MsgBox("No ha ingresado ningun valor para realizar la busqueda", vbCritical + vbOKOnly, "Error")
    Text5.SetFocus
    Exit Sub
End If

'Dim dbbusc As Database
'Dim rsbusc As Recordset

Set dbbusc = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
query = strQ & " " & strQry

'asasd = InputBox("bla", , query)


Set rsbusc = dbbusc.OpenRecordset(query)
'Set rs = db.OpenRecordset("Select Codigo_recomendacion, fecha_solicitud, respons_solici, referencia, esp_plano_norma, cliente, observaciones, solicitud_documentacion, fecha_sol_docu, comp_recomendado, fecha_recomend, responsable_recomend, observ_lab  From compuestos_para_cotizacion where codigo_recomendacion IS NOT NULL AND cliente = Bafir")
If rsbusc.RecordCount = 0 Then


    fsgsdh = MsgBox("No se ha encontrado ningun registro coincidente", vbInformation + vbOKOnly, "Busqueda")
    Text5.SetFocus
    dbbusc.Close
    Exit Sub
End If

'rsbusc.MoveFirst
Text5.Text = rsbusc.Fields("codigo_recomendacion")
Text8.Text = rsbusc.Fields("fecha_solicitud")
Text9.Text = rsbusc.Fields("respons_solici")
Text10.Text = rsbusc.Fields("referencia")
Text11.Text = rsbusc.Fields("esp_plano_norma")
Text12.Text = rsbusc.Fields("cliente")
Text7.Text = rsbusc.Fields("observaciones")
Text1.Text = rsbusc.Fields("solicitud_documentacion")
Text2.Text = rsbusc.Fields("fecha_sol_docu")
Text3.Text = rsbusc.Fields("comp_recomendado")
Text13.Text = rsbusc.Fields("fecha_recomend")
Text4.Text = rsbusc.Fields("responsable_recomend")
Text6.Text = rsbusc.Fields("observ_lab")
Command14.Enabled = True
Command13.Enabled = True
Command11.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command12.Enabled = True
Command15.Enabled = True
posicion = rsbusc.AbsolutePosition
rsbusc.MoveLast
total = rsbusc.RecordCount - 1
rsbusc.MoveFirst
Command2.Enabled = False
If total = posicion Then
Command5.Enabled = False
Else
Command5.Enabled = True
End If
Flag = True
'dbbusc.Close
Command3.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Command4_Click()
Command14.Enabled = False
Command13.Enabled = False
Command11.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
'Command11.Enabled = False
Command12.Enabled = False
Command15.Enabled = False
frmBusca.Text5.Text = ""
frmBusca.Text8.Text = ""
frmBusca.Text9.Text = ""
frmBusca.Text10.Text = ""
frmBusca.Text11.Text = ""
frmBusca.Text12.Text = ""
frmBusca.Text7.Text = ""
frmBusca.Text1.Text = ""
frmBusca.Text2.Text = ""
frmBusca.Text3.Text = ""
frmBusca.Text13.Text = ""
frmBusca.Text4.Text = ""
frmBusca.Text6.Text = ""
Command3.Enabled = True
Command4.Enabled = False
Command2.Enabled = False
Command5.Enabled = False
Flag = False
On Error Resume Next
rsbusc.Close
End Sub

Private Sub Command5_Click()
rsbusc.MoveNext
Text5.Text = rsbusc.Fields("codigo_recomendacion")
Text8.Text = rsbusc.Fields("fecha_solicitud")
Text9.Text = rsbusc.Fields("respons_solici")
Text10.Text = rsbusc.Fields("referencia")
Text11.Text = rsbusc.Fields("esp_plano_norma")
Text12.Text = rsbusc.Fields("cliente")
Text7.Text = rsbusc.Fields("observaciones")
Text1.Text = rsbusc.Fields("solicitud_documentacion")
Text2.Text = rsbusc.Fields("fecha_sol_docu")
Text3.Text = rsbusc.Fields("comp_recomendado")
Text13.Text = rsbusc.Fields("fecha_recomend")
Text4.Text = rsbusc.Fields("responsable_recomend")
Text6.Text = rsbusc.Fields("observ_lab")
posicion = rsbusc.AbsolutePosition
If posicion = total Then
Command5.Enabled = False
End If
If posicion <> 0 Then
Command2.Enabled = True
End If
End Sub

Private Sub Command6_Click()




Dim appp As New Excel.Application
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim r As Excel.Range

Set wb = appp.Workbooks.Open("\\Servidor2\e\EntornoBafir\Planillas\informecotiz.xls", , True)
Set ws = wb.Worksheets(1)


ws.Cells(13, 4) = Text5.Text
ws.Cells(15, 4) = Text8.Text
ws.Cells(17, 4) = Text9.Text
ws.Cells(19, 4) = Text10.Text
ws.Cells(23, 4) = Text11.Text
ws.Cells(27, 4) = Text12.Text
ws.Cells(29, 4) = Text7.Text
ws.Cells(33, 4) = Text1.Text
ws.Cells(36, 4) = Text2.Text
ws.Cells(38, 4) = Text3.Text
ws.Cells(41, 4) = Text4.Text
ws.Cells(43, 4) = Text6.Text


ws.PrintOut
DoEvents
wb.Close (False)





End Sub

Private Sub Command7_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Ingeniería", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'ingkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select revisado_ing from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)

rs.Edit
rs.Fields("revisado_ing") = True
rs.Update
db.Close

sdfsdfsf = MsgBox("Se ha marcado la recomendación como leido", vbInformation + vbOKOnly, "Marcado como leido")
End Sub

Private Sub Command8_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Ingeniería", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'ingkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select revisado_ing from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)

rs.Edit
rs.Fields("revisado_ing") = False
rs.Update
db.Close

sdfsdfsf = MsgBox("Se ha marcado la recomendación como no leido", vbInformation + vbOKOnly, "Marcado como NO leido")
End Sub

Private Sub Command9_Click()
frmPassword.Show (1)
contra = frmPassword.Password
'contra = InputBox("Ingrese contraseña de Ingeniería", "Contraseña")
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'ingkey'")



If UCase(rs.Fields("dato")) = UCase(contra) Then

Else
fsdfss = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
Exit Sub
End If


frmBusca.Enabled = False
Set dbbusc = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rsbusc = dbbusc.OpenRecordset("Select revisado_lab,revisado_ing, Observaciones from compuestos_para_cotizacion where codigo_recomendacion = " & Text5.Text)
'caracteres = CInt(Label4.Caption) - 1
Campo = "observaciones"
frmModificaCotizaciones.Show
'frmModificaCotizaciones.Label1 = Label4.Caption
Command9.Enabled = False
Command10.Enabled = False
Command12.Enabled = False
Command15.Enabled = False
frmModificaCotizaciones.Text1.Text = ""
Label2.Caption = ""
frmModificaCotizaciones.Label2 = Text7.Text
frmModificaCotizaciones.Text1.SetFocus





End Sub

Private Sub Text1_Change()
'Label21.Caption = 255 - Len(Text1.Text)
End Sub


Private Sub Text6_Change()
'Label23.Caption = 255 - Len(Text6.Text)
End Sub

Private Sub Text7_Change()
'Label4.Caption = 255 - Len(Text7.Text)
End Sub
