VERSION 5.00
Begin VB.Form frmfluido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fluidos"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text15 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Promedio"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   55
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Promedio"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   54
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   5760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Promedio"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   49
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Promedio"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   48
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   360
      TabIndex        =   47
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Seleccionar Probeta"
      Height          =   375
      Left            =   2760
      TabIndex        =   46
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Promedio"
      Height          =   255
      Left            =   9960
      TabIndex        =   45
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Promedio"
      Height          =   255
      Left            =   8640
      TabIndex        =   44
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Promedio"
      Height          =   255
      Left            =   7320
      TabIndex        =   43
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Promedio"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   42
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Promedio"
      Height          =   255
      Left            =   9960
      TabIndex        =   41
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Promedio"
      Height          =   255
      Left            =   8640
      TabIndex        =   40
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Promedio"
      Height          =   255
      Left            =   7320
      TabIndex        =   39
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Promedio"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   38
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   8520
      TabIndex        =   4
      Text            =   "Text11"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   495
      Left            =   5160
      TabIndex        =   20
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ingresar"
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text10"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9960
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8640
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      Height          =   2295
      Left            =   360
      TabIndex        =   23
      Top             =   1680
      Width           =   1335
      Begin VB.OptionButton Option2 
         Caption         =   "Envejecido"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8640
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "Peso en aire 3"
      Height          =   255
      Left            =   5760
      TabIndex        =   53
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "Peso en aire 2"
      Height          =   255
      Left            =   4560
      TabIndex        =   52
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Peso en aire 3"
      Height          =   255
      Left            =   5760
      TabIndex        =   51
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Peso en aire 2"
      Height          =   255
      Left            =   4560
      TabIndex        =   50
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   7080
      Y1              =   2280
      Y2              =   3960
   End
   Begin VB.Label Label15 
      Caption         =   "Tiempo de repetición en días"
      Height          =   375
      Left            =   8520
      TabIndex        =   37
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Peso en agua 3"
      Height          =   255
      Left            =   9960
      TabIndex        =   36
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Peso en agua 3"
      Height          =   255
      Left            =   9960
      TabIndex        =   35
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Peso en agua 2"
      Height          =   255
      Left            =   8640
      TabIndex        =   34
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Peso en agua 2"
      Height          =   255
      Left            =   8640
      TabIndex        =   33
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Peso en agua 1"
      Height          =   255
      Left            =   7320
      TabIndex        =   32
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Peso en agua 1"
      Height          =   255
      Left            =   7320
      TabIndex        =   31
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Peso en aire"
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Peso en aire 1"
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Ensayo"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Partida"
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmfluido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vartracc
Public varelong
Public vardureza
Public asdasd As String

Private Sub Command1_Click()
'''''''''''''viejo
'frmEnsayo.Show
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'frmEnsayo.Caption = Combo1.Text & " " & Text10.Text
'frmfluido.Enabled = False
''''''''''''''viejo
Me.Enabled = False

Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

    
    rst.Open "SELECT codigo FROM ensayos order by codigo desc", cnn, adOpenStatic, adLockReadOnly
    codigoens = rst.Fields("codigo") + 1
    

frmEnsayoAgregarNuevo.Text1.Text = codigoens
frmEnsayoAgregarNuevo.List1.Clear
frmEnsayoAgregarNuevo.List2.Clear
frmEnsayoAgregarNuevo.List1.AddItem ("Envejecimiento")
frmEnsayoAgregarNuevo.List1.AddItem ("Compresion")
    rst.Close
    rst.Open "SELECT nombre FROM stock where tipo = 'Fluidos' or tipo = 'Reactivos' order by nombre asc", cnn, adOpenStatic, adLockReadOnly

frmEnsayoAgregarNuevo.List2.AddItem "Aire"
Do Until rst.EOF = True
    frmEnsayoAgregarNuevo.List2.AddItem (rst.Fields("Nombre"))
    rst.MoveNext
Loop




cnn.Close
frmEnsayoAgregarNuevo.List1.Text = "Envejecimiento"
frmEnsayoAgregarNuevo.Command1.Visible = False
frmEnsayoAgregarNuevo.Command2.Visible = False
frmEnsayoAgregarNuevo.Command3.Visible = False
frmEnsayoAgregarNuevo.Command4.Visible = False
frmEnsayoAgregarNuevo.Command5.Visible = False
frmEnsayoAgregarNuevo.Command6.Visible = False
frmEnsayoAgregarNuevo.Command7.Visible = True
frmEnsayoAgregarNuevo.Command8.Visible = True

frmEnsayoAgregarNuevo.Show
frmEnsayoAgregarNuevo.Text2.SetFocus
End Sub

Private Sub Command10_Click()
asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text7.Text = este / asdasd
End Sub

Private Sub Command11_Click()
asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text8.Text = este / asdasd
End Sub

Private Sub Command12_Click()
frmfluido.Enabled = False
frmSelProbeta.Show (1)
frmfluido.Enabled = True
probeta = frmSelProbeta.probeta
Check1.Value = 1
Check1.Caption = probeta
End Sub

Private Sub Command2_Click()
Static vartracc
Static varelong
Static vardureza

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset

If Option1.Value = True Then 'original
    If Check1.Value = 0 Then
        sdfsd = MsgBox("Debe seleccionar probeta", vbCritical + vbOKOnly, "No seleccionó probeta")
        Command12.SetFocus
        Exit Sub
    End If
    If Combo1.Text = "" Then
    sdffsdf = MsgBox("Debe seleccionar un compuesto", vbCritical + vbOKOnly, "Error")
    Combo1.SetFocus
    Exit Sub
    End If
    If Text10.Text = "" Then
    sdffsdf = MsgBox("Debe indicar partida", vbCritical + vbOKOnly, "Error")
    Text10.SetFocus
    Exit Sub
    End If
    If List1.Text = "" Then
    sdffsdf = MsgBox("Debe seleccionar un Ensayo", vbCritical + vbOKOnly, "Error")
    List1.SetFocus
    Exit Sub
    End If
    ' para masa
    
    
        If Text1.Text = "" Or Text12 = "" Or Text13 = "" Then
        sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
        Text1.SetFocus
        Exit Sub
        End If
    If Text2.Text = "" And Text3.Text = "" And Text4.Text = "" Then
        asdd = MsgBox("Se realizará solo la variación de masa", vbInformation + vbOKOnly, "Variación de masa")
    Else
    
        If Text2.Text = "" Then
        sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
        Text2.SetFocus
        Exit Sub
        End If
        
        If Text3.Text = "" Then
        sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
        Text3.SetFocus
        Exit Sub
        End If
        
        If Text4.Text = "" Then
        sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
        Text4.SetFocus
        Exit Sub
        End If
    End If
    
    If Text11.Text = "" Then
    sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text11.SetFocus
    Exit Sub
    End If
    
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset


    sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    rst.Open "SELECT codigo FROM ensayos where referencia = '" & List1.Text & "'", cnn, adOpenStatic, adLockReadOnly
    
    codref = rst.Fields("codigo")
    
    cnn.Close
    
    
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
    Set rs = db.OpenRecordset("SELECT codref,probeta, codigo, n_formula, partida, tiemp_temp_oil, var_vol, var_tracc, var_elong, var_dur, fecha_realizacion, tiempo_repeticion, aprovado FROM fluidos")
    Set rs1 = db.OpenRecordset("SELECT codigo, valmasa, valmasa2, valmasa3, valagua1, valagua2, valagua3 FROM fluido_temp")
    Set rs2 = db.OpenRecordset("SELECT * FROM densidades")
    rs.AddNew
    rs.Fields("codigo") = Label1.Caption
    rs.Fields("codref") = codref
    rs.Fields("N_formula") = Combo1.Text
    rs.Fields("Partida") = Text10.Text
    rs.Fields("tiemp_temp_oil") = List1.Text
    rs.Fields("var_vol") = "0"
    rs.Fields("var_TRACC") = "0"
    rs.Fields("var_ELONG") = "0"
    rs.Fields("var_dur") = "0"
    rs.Fields("fecha_realizacion") = Date
    rs.Fields("tiempo_repeticion") = Text11.Text
    rs.Fields("aprovado") = 0
    rs.Fields("probeta") = Check1.Caption
    rs.Update
    rs1.AddNew
    rs1.Fields("codigo") = Label1.Caption
    rs1.Fields("valmasa") = Text1.Text
    rs1.Fields("valmasa2") = Text12.Text
    rs1.Fields("valmasa3") = Text13.Text
    rs1.Fields("valagua1") = Text2.Text
    rs1.Fields("valagua2") = Text3.Text
    rs1.Fields("valagua3") = Text4.Text
    rs1.Update
    
    rs2.AddNew
    rs2.Fields("Fecha") = Date
    rs2.Fields("compuesto") = Combo1.Text
    rs2.Fields("partida") = Text10.Text
    masa = (CDbl(Text1.Text) + CDbl(Text12.Text) + CDbl(Text13.Text)) / 3
    agua = (CDbl(Text2.Text) + CDbl(Text3.Text) + CDbl(Text4.Text)) / 3
    densidad = (0.9971 * masa) / (masa - agua)
    rs2.Fields("densidad") = densidad
    rs2.Update
    rs2.Close
    
    frmfluido.Text10.Text = ""
    frmfluido.Text1.Text = ""
    frmfluido.Text2.Text = ""
    frmfluido.Text3.Text = ""
    frmfluido.Text4.Text = ""
    frmfluido.Text5.Text = ""
    frmfluido.Text6.Text = ""
    frmfluido.Text7.Text = ""
    frmfluido.Text8.Text = ""
    frmfluido.Text9.Text = ""
    frmfluido.Text11.Text = ""
    List1.Text = ""
    
    
    'Combo1.SetFocus
    frmFluidos.Enabled = True
    frmFluidos.Visible = True
    frmfluido.Hide
    
Else 'arriba original, abajo envejecido
    If Text9.Text = "" Then
    sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text9.SetFocus
    Exit Sub
    End If
    
    If Text5.Text = "" Then
    sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text5.SetFocus
    Exit Sub
    End If
    
    If Text14.Text = "" Then
    sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text14.SetFocus
    Exit Sub
    End If
    
    If Text15.Text = "" Then
    sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text15.SetFocus
    Exit Sub
    End If
    
    If Text2.Text = "" Then
    
    Else
        If Text6.Text = "" Then
        sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
        Text6.SetFocus
        Exit Sub
        End If
        
        If Text7.Text = "" Then
        sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
        Text7.SetFocus
        Exit Sub
        End If
        
        If Text8.Text = "" Then
        sdffsdf = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
        Text8.SetFocus
        Exit Sub
        End If
    End If
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
    Set rs = db.OpenRecordset("SELECT var_masa, var_masa1, var_masa2, codigo, n_formula, partida, tiemp_temp_oil, var_vol, var_vol1, var_vol2, var_tracc, var_elong, var_dur, fecha_realizacion, tiempo_repeticion, aprovado FROM fluidos where codigo = " & Text9.Text & "")
    Set rs1 = db.OpenRecordset("SELECT codigo, valmasa, valmasa2, valmasa3, valagua1, valagua2, valagua3 FROM fluido_temp where codigo = '" & Text9.Text & "'")
    ''''''''''''''''''aca empezamos a modificar 080124
    masaoriginal = rs1.Fields("valmasa")
    Dim a As Integer
    a = InStr(1, (masaoriginal), ".")
    If a <> 0 Then
    Mid(masaoriginal, a) = ","
    End If
    
    masa2 = rs1.Fields("valmasa2")
    masa2 = Replace(masa2, ".", ",")
    
    masa3 = rs1.Fields("valmasa3")
    masa3 = Replace(masa3, ".", ",")
        
    aguaoriginal1 = rs1.Fields("valagua1")
    a = InStr(1, (aguaoriginal1), ".")
     If a <> 0 Then
    Mid(aguaoriginal1, a) = ","
    End If
    
    aguaoriginal2 = rs1.Fields("valagua2")
    a = InStr(1, (aguaoriginal2), ".")
     If a <> 0 Then
    Mid(aguaoriginal2, a) = ","
    End If
    
    aguaoriginal3 = rs1.Fields("valagua3")
    a = InStr(1, (aguaoriginal3), ".")
     If a <> 0 Then
    Mid(aguaoriginal3, a) = ","
    End If
    
    masaenvejecida = Text5.Text
    a = InStr(1, (masaenvejecida), ".")
     If a <> 0 Then
    Mid(masaenvejecida, a) = ","
    End If
    
    masaenvejecida2 = Text14.Text
    masaenvejecida2 = Replace(masaenvejecida2, ".", ",")
    
    masaenvejecida3 = Text15.Text
    masaenvejecida3 = Replace(masaenvejecida3, ".", ",")
    
    aguaenvejecida1 = Text6.Text
    a = InStr(1, (aguaenvejecida1), ".")
     If a <> 0 Then
    Mid(aguaenvejecida1, a) = ","
    End If
    
    aguaenvejecida2 = Text7.Text
    a = InStr(1, (aguaenvejecida2), ".")
     If a <> 0 Then
    Mid(aguaenvejecida2, a) = ","
    End If
    
    aguaenvejecida3 = Text8.Text
    a = InStr(1, (aguaenvejecida3), ".")
     If a <> 0 Then
    Mid(aguaenvejecida3, a) = ","
    End If
    
    masaoriginal = CDbl(masaoriginal)
    masa2 = CDbl(masa2)
    masa3 = CDbl(masa3)
    
    If Text2.Text <> "" Then
        aguaoriginal1 = CDbl(aguaoriginal1)
        aguaoriginal2 = CDbl(aguaoriginal2)
        aguaoriginal3 = CDbl(aguaoriginal3)
        'aguaoriginalprom = (aguaoriginal1 + aguaoriginal2 + aguaoriginal3) / 3
    End If
    
    masaenvejecida = CDbl(masaenvejecida)
    
    If Text2.Text <> "" Then
        aguaenvejecida1 = CDbl(aguaenvejecida1)
        aguaenvejecida2 = CDbl(aguaenvejecida2)
        aguaenvejecida3 = CDbl(aguaenvejecida3)
        'aguaenvejecidaprom = (aguaenvejecida1 + aguaenvejecida2 + aguaenvejecida3) / 3
    End If
    
    'vv = (((masaenvejecida - aguaenvejecidaprom) - (masaoriginal - aguaoriginalprom)) / (masaoriginal - aguaoriginalprom)) * 100
    vv = (((masaenvejecida - aguaenvejecida1) - (masaoriginal - aguaoriginal1)) / (masaoriginal - aguaoriginal1)) * 100
    vv1 = (((masaenvejecida2 - aguaenvejecida2) - (masa2 - aguaoriginal2)) / (masa2 - aguaoriginal2)) * 100
    vv2 = (((masaenvejecida3 - aguaenvejecida3) - (masa3 - aguaoriginal3)) / (masa3 - aguaoriginal3)) * 100
    vvp = (vv + vv1 + vv2) / 3
    
    
    'vmasa = (masaenvejecida * 100 / masaoriginal) - 100
    vmasa = (masaenvejecida * 100 / masaoriginal) - 100
    vmasa1 = (masaenvejecida2 * 100 / masa2) - 100
    vmasa2 = (masaenvejecida3 * 100 / masa3) - 100
    vmp = (vmasa + vmasa1 + vmasa2) / 3
    
    
    
    If Text2.Text <> "" Then
    sadfsdf = MsgBox("La variación de volumen del ensayo es de " & vvp & "." & "Desea ingresar las variaciones de tracción, elongación y dureza?", vbInformation + vbYesNo, "Carga de datos")
    Else
    sadfsdf = MsgBox("La variación de masa del ensayo es de " & vmp & "." & "Desea ingresar las variaciones de tracción, elongación y dureza?", vbInformation + vbYesNo, "Carga de datos")
    End If
    
    
    If sadfsdf = vbNo Then
    rs.Edit
    If Text2.Text <> "" Then
        rs.Fields("var_vol") = vv
        rs.Fields("var_vol1") = vv1
        rs.Fields("var_vol2") = vv2
    End If
    rs.Fields("var_masa") = vmasa
    rs.Fields("var_masa1") = vmasa1
    rs.Fields("var_masa2") = vmasa2
    
    rs.Update
    rs1.Edit
    rs1.Delete
    'rs1.Update
    Else
    frmIngresaVar.Text1.Text = ""
    frmIngresaVar.Text2.Text = ""
    frmIngresaVar.Text3.Text = ""
    frmIngresaVar.Caption = Combo1.Text & " " & Text10.Text & " " & List1.Text
    
    frmIngresaVar.Show (1)
    
    rs.Edit
    rs.Fields("var_vol") = vv
    rs.Fields("var_vol1") = vv1
    rs.Fields("var_vol2") = vv2
    rs.Fields("var_masa") = vmasa
    rs.Fields("var_masa1") = vmasa1
    rs.Fields("var_masa2") = vmasa2
    rs.Fields("var_tracc") = frmfluido.vartracc
    rs.Fields("var_elong") = frmfluido.varelong
    rs.Fields("var_dur") = frmfluido.vardureza
    rs.Update
    rs1.Edit
    rs1.Delete
    
    
    End If
End If

sdfdf = MsgBox("Se han guardado los datos para el ensayo nº " & Label1.Caption & ".", vbInformation + vbOKOnly, "Info")
frmfluido.Label1.Caption = ""

Set rs = db.OpenRecordset("Select N_FORMULA From Fluidos GROUP BY N_FORMULA")

rs.MoveFirst
Do While rs.EOF = False
frmFluidos.Combo1.AddItem (rs.Fields("N_FORMULA"))
rs.MoveNext
Loop




db.Close
frmFluidos.Enabled = True
frmFluidos.Visible = True
frmfluido.Hide
End Sub

Private Sub Command3_Click()
frmFluidos.Enabled = True
frmFluidos.Visible = True
frmfluido.Hide
End Sub

Private Sub Command4_Click()

asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
    Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text1.Text = este / asdasd
End Sub

Private Sub Command5_Click()
asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text2.Text = este / asdasd
End Sub

Private Sub Command6_Click()
asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
If Valor = "" Then
Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text3.Text = este / asdasd
End Sub

Private Sub Command7_Click()
asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
If Valor = "" Then
Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text4.Text = este / asdasd
End Sub

Private Sub Command8_Click()
asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
If Valor = "" Then
Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text5.Text = este / asdasd
End Sub

Private Sub Command9_Click()
asdasd = InputBox("Cantidad de valores a promediar?", "Promedio")
If asdasd = "" Then
Exit Sub
End If
If IsInteger(asdasd) = False Then
    dfsdf = MsgBox("El número de valores no es correcto", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
a = InStr(1, (asdasd), ".")
    If a <> 0 Then
    Mid(asdasd, a) = ","
    End If
asdasd = CInt(asdasd)
este = 0
For hacer = 1 To asdasd
Valor = InputBox("Ingrese el valor " & hacer & " de " & asdasd, "Ingreso de datos")
If Valor = "" Then
Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text6.Text = este / asdasd
End Sub

Private Sub Option1_Click()
Combo1.Enabled = True
Text10.Enabled = True
List1.Enabled = True
Combo1.Text = ""
List1.Text = ""
Command1.Enabled = True
Text1.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text11.Enabled = True
Command4.Enabled = True
Command5.Enabled = True

Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False


frmfluido.Text10.Text = ""
frmfluido.Text1.Text = ""
frmfluido.Label1.Caption = ""
frmfluido.Text2.Text = ""
frmfluido.Text3.Text = ""
frmfluido.Text4.Text = ""
frmfluido.Text5.Text = ""
frmfluido.Text6.Text = ""
frmfluido.Text7.Text = ""
frmfluido.Text8.Text = ""
frmfluido.Text9.Text = ""
frmfluido.Text11.Text = ""
List1.Text = ""

Dim db As Database
Dim rs2 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs2 = db.OpenRecordset("SELECT codigo FROM fluidos")
rs2.MoveLast
Codigo = CInt(rs2.Fields("codigo"))
Codigo = Codigo + 1
Label1.Caption = Codigo
db.Close
End Sub

Private Sub Option2_Click()
Combo1.Enabled = False
Text10.Enabled = False
List1.Enabled = False
Label1.Caption = ""
Command1.Enabled = False
Text1.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text11.Enabled = False
Text9.SetFocus
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
frmfluido.Text10.Text = ""
frmfluido.Text1.Text = ""
frmfluido.Label1.Caption = ""
frmfluido.Text2.Text = ""
frmfluido.Text3.Text = ""
frmfluido.Text4.Text = ""
frmfluido.Text5.Text = ""
frmfluido.Text6.Text = ""
frmfluido.Text7.Text = ""
frmfluido.Text8.Text = ""
frmfluido.Text9.Text = ""
frmfluido.Text11.Text = ""
List1.Text = ""
End Sub

Private Sub Text5_GotFocus()
If Text9.Text = "" Then
Text9.SetFocus
End If

End Sub

Private Sub Text6_gotfocus()
If Text9.Text = "" Then
Text9.SetFocus
End If
End Sub

Private Sub Text7_gotfocus()
If Text9.Text = "" Then
Text9.SetFocus
End If
End Sub

Private Sub Text8_gotfocus()
If Text9.Text = "" Then
Text9.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
    If Text9.Text = "" Then
    Exit Sub
    Else
    Dim db As Database
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
    Set rs = db.OpenRecordset("SELECT probeta, codigo, n_formula, partida, tiemp_temp_oil, var_vol, var_tracc, var_elong, var_dur, fecha_realizacion, tiempo_repeticion, aprovado FROM fluidos where codigo =" & Text9.Text & "")
    Set rs1 = db.OpenRecordset("SELECT codigo, valmasa, valmasa2, valmasa3, valagua1, valagua2, valagua3 FROM fluido_temp where codigo ='" & Text9.Text & "'")
    
    If rs.RecordCount = 0 Then
    sdfdfs = MsgBox("No se encuentra el ensayo", vbCritical + vbOKOnly, "Error")
    Text9.Text = ""
    Text9.SetFocus
    Exit Sub
    End If
    If rs.RecordCount <> 0 And rs1.RecordCount = 0 Then
    sdfsdfsdf = MsgBox("Este ensayo se encuentra terminado", vbInformation + vbOKOnly, "Ensayo cerrado")
    Text9.Text = ""
    Text9.SetFocus
    Exit Sub
    End If
    Text5.Enabled = True
    Text14.Enabled = True
    Text15.Enabled = True
    
    Label1.Caption = Text9.Text
    Combo1.Text = rs.Fields("n_formula")
    Text10.Text = rs.Fields("partida")
    List1.Text = rs.Fields("tiemp_temp_oil")
    Text11.Text = rs.Fields("tiempo_repeticion")
    Check1.Caption = rs.Fields("probeta") & ""
    If Check1.Caption <> "" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    Text1.Text = rs1.Fields("valmasa")
    Text12.Text = rs1.Fields("valmasa2")
    Text13.Text = rs1.Fields("valmasa3")
    Text2.Text = rs1.Fields("valagua1")
    Text3.Text = rs1.Fields("valagua2")
    Text4.Text = rs1.Fields("valagua3")
    
    If Text2.Text = "" Then
        Text6.Enabled = False
        Text7.Enabled = False
        Text8.Enabled = False
    End If
        
    sdfsdf = MsgBox("Se realizará solo la variación de masa", vbInformation + vbOKOnly, "Variación de masa")
    
    Check1.Caption = rs.Fields("probeta") & ""
    
db.Close
    End If
    Text5.SetFocus
End Sub
