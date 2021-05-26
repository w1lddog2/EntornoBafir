VERSION 5.00
Begin VB.Form FrmFluidoNuevo 
   Caption         =   "Cargar Fluido"
   ClientHeight    =   3930
   ClientLeft      =   3315
   ClientTop       =   2445
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   3930
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Tag             =   "/"
   Begin VB.CommandButton Command8 
      Caption         =   "Pendientes"
      Height          =   255
      Left            =   6120
      TabIndex        =   28
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Buscar Tracciones"
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ingresar Envej."
      Height          =   255
      Left            =   7920
      TabIndex        =   26
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ingresar Original"
      Height          =   255
      Left            =   7920
      TabIndex        =   25
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
      Height          =   615
      Left            =   7920
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ensayo Aprobado"
      Height          =   495
      Left            =   3840
      TabIndex        =   21
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar otro"
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   615
      Left            =   7920
      TabIndex        =   18
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   615
      Left            =   7920
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Text            =   "6"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Text            =   "5"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Text            =   "4"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Text            =   "3"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Text            =   "2"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   360
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   7920
      X2              =   9360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   7920
      X2              =   9360
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label12 
      Caption         =   "Modo Automático"
      Height          =   255
      Left            =   7920
      TabIndex        =   24
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Modo Manual"
      Height          =   255
      Left            =   7920
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Duración (Días)"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Var. Dureza"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Var. Elong."
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Var. Tracción"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Var. Vol o masa"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Ensayo"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Partida"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "FrmFluidoNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset
If Combo1.Text = "" Then
    age = MsgBox("Debe seleccionar un compuesto", vbCritical + vbOKOnly, "Error")
    Combo1.SetFocus
    Exit Sub
End If
If Text1.Text = "" Then
    age = MsgBox("Debe ingresar una partida", vbCritical + vbOKOnly, "Error")
    Text1.SetFocus
    Exit Sub
End If
If List1.Text = "" Then
    age = MsgBox("Debe seleccionar un ensayo", vbCritical + vbOKOnly, "Error")
    List1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    age = MsgBox("Ingresar el valor", vbCritical + vbOKOnly, "Error")
    Text2.SetFocus
    Exit Sub
End If
If Text3.Text = "" Then
    age = MsgBox("Ingresar el valor", vbCritical + vbOKOnly, "Error")
    Text3.SetFocus
    Exit Sub
End If
If Text4.Text = "" Then
    age = MsgBox("Ingresar el valor", vbCritical + vbOKOnly, "Error")
    Text4.SetFocus
    Exit Sub
End If
If Text5.Text = "" Then
    age = MsgBox("Ingresar el valor", vbCritical + vbOKOnly, "Error")
    Text5.SetFocus
    Exit Sub
End If
If Text6.Text = "" Then
    age = MsgBox("Ingresar el valor", vbCritical + vbOKOnly, "Error")
    Text6.SetFocus
    Exit Sub
End If
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
'Set rs = db.OpenRecordset("Select tiempo_repeticion, fecha_realizacion, var_dur, N_formula, tiemp_temp_oil, var_vol, var_tracc, var_elong, partida, aprovado FROM fluidos WHERE N_FORMULA = '" & Combo1.Text & "' AND tiemp_temp_oil = '" & List1.Text & "' ;")
Set rs = db.OpenRecordset("Select codigo, tiempo_repeticion, fecha_realizacion, var_dur, N_formula, tiemp_temp_oil, var_vol, var_tracc, var_elong, partida, aprovado FROM fluidos WHERE codigo = '" & Label13.Caption & "' ;")
If rs.RecordCount > 0 Then
rs.Edit
Else
rs.AddNew
End If
rs.Fields("codigo") = Label13.Caption
rs.Fields("N_formula") = Combo1.Text
rs.Fields("partida") = Text1.Text
rs.Fields("Tiemp_temp_oil") = List1.Text
rs.Fields("var_vol") = Text2.Text
rs.Fields("var_tracc") = Text3.Text
rs.Fields("var_elong") = Text4.Text
rs.Fields("var_dur") = Text5.Text
rs.Fields("fecha_realizacion") = Date
rs.Fields("tiempo_repeticion") = Text6.Text
rs.Fields("aprovado") = Check1.Value
rs.Update
db.Close
gfd = MsgBox("Los datos del fluido se han cargado satisfactoriamente", vbInformation + vbOKOnly, "Carga de datos")
frmFluidos.Enabled = True
frmFluidos.Visible = True
frmFluidos.Option3 = True
FrmFluidoNuevo.Hide

End Sub

Private Sub Command2_Click()
If Combo1.Text <> "" Then
    fsdfsdf = MsgBox("No ha grabado el registro. Si sale perderá todos los datos. Desea continuar?", vbCritical + vbYesNo, "Atención")
    If fsdfsdf = vbNo Then
        Exit Sub
    End If
End If
frmFluidos.Enabled = True
frmFluidos.Combo1.Clear
frmFluidos.Combo2.Clear
frmFluidos.Visible = True
FrmFluidoNuevo.Hide
frmFluidos.Option1 = True
    
End Sub


Private Sub Command3_Click()
frmEnsayo.Show
FrmFluidoNuevo.Enabled = False
FrmFluidoNuevo.Visible = False
frmEnsayo.Text1.Text = ""
frmEnsayo.Text2.Text = ""
frmEnsayo.Text3.Text = ""
End Sub

Private Sub Command4_Click()
Do
numeroprobetas = InputBox("Ingrese el número de probetas", "Cantidad de probetas")
If numeroprobetas = "" Then
Exit Sub
End If
If IsNumeric(numeroprobetas) = False Then
 asdasf = MsgBox("Ingrese un número válido")
End If
Loop Until IsNumeric(numeroprobetas) = True
numeroprobetas = CInt(numeroprobetas)
ReDim pesoaireoriginal(1 To numeroprobetas)
ReDim pesoaguaoriginal(1 To numeroprobetas)
Dim traccionoriginal
Dim elongacionoriginal
Dim durezaoriginal
ReDim pesoaireenvejecido(1 To numeroprobetas)
ReDim pesoaguaenvejecido(1 To numeroprobetas)
Dim traccionenvejecido
Dim elongacionenvejecido
Dim durezaenvejecido
For aireo = 1 To numeroprobetas
    pesoaireoriginal(aireo) = InputBox("Ingrese el peso en aire original " & aireo & "/" & numeroprobetas, "Peso en aire original")
    a = InStr(1, (pesoaireoriginal(aireo)), ".")
     If a <> 0 Then
    Mid(pesoaireoriginal(aireo), a) = ","
    End If
    
Next 'aireo
For airee = 1 To numeroprobetas
    pesoaireenvejecido(airee) = InputBox("Ingrese el peso en aire envejecido " & airee & "/" & numeroprobetas, "Peso en aire envejecido")
    a = InStr(1, pesoaireenvejecido(airee), ".")
    If a <> 0 Then
    Mid(pesoaireenvejecido(airee), a) = ","
    End If
    
Next 'airee

For Aguao = 1 To numeroprobetas
    cuantosaguaori = InputBox("Ingrese la cantidad de valores que ha tomado del peso en agua original", Aguao & "/" & numeroprobetas & " Original")
    pesoaguaoriginal(Aguao) = 0
    temporal = 0
    For promedioaguao = 1 To cuantosaguaori
        'pesoaguaoriginal(aguao) = InputBox("Ingrese el peso en agua original " & aguao & "/" & numeroprobetas & "(" & promedioaguao & "/" & cuantosaguaori & ")", "Peso en agua original")
        temporal = InputBox("Ingrese el peso en agua original " & Aguao & "/" & numeroprobetas & "(" & promedioaguao & "/" & cuantosaguaori & ")", "Peso en agua original")
        'Next 'promedioaguao
        'pesoaguaoriginal(aguao) = pesoaguaoriginal(aguao) / cuantosaguaori
        'For loopaguao = 1 To numeroprobetas
        a = InStr(1, temporal, ".")
        If a <> 0 Then
            Mid(temporal, a) = ","
        End If
        pesoaguaoriginal(Aguao) = pesoaguaoriginal(Aguao) + temporal
        'Next 'loopaguao
        'b = pesoaguaoriginal(1)
        'For loopaguao1 = 2 To numeroprobetas
        'b = b + pesoaguaoriginal(loopaguao1)
        'Next
    Next ' promedioaguao
    pesoaguaoriginal(Aguao) = pesoaguaoriginal(Aguao) / cuantosaguaori
Next 'aguao
For aguae = 1 To numeroprobetas
    cuantosaguaenv = InputBox("Ingrese la cantidad de valores que ha tomado del peso en agua envejecido", aguae & "/" & numeroprobetas & " Envejecidos")
    pesoaguaenvejecido(aguae) = 0
    temporal = 0
    For promedioaguae = 1 To cuantosaguaenv
        'pesoaguaenvejecido(promedioaguae) = InputBox("Ingrese el peso en agua envejecido " & aguae & "/" & numeroprobetas & "(" & promedioaguae & "/" & cuantosaguaenv & ")", "Peso en agua envejecido")
        temporal = InputBox("Ingrese el peso en agua envejecido " & aguae & "/" & numeroprobetas & "(" & promedioaguae & "/" & cuantosaguaenv & ")", "Peso en agua envejecido")
        'Next promedioaguae
        'For loopaguae = 1 To numeroprobetas
        a = InStr(1, temporal, ".")
        If a <> 0 Then
        Mid(temporal, a) = ","
        End If
        pesoaguaenvejecido(aguae) = pesoaguaenvejecido(aguae) + temporal
        'Next ' loopaguae
    Next ' promedioaguae
    pesoaguaenvejecido(aguae) = pesoaguaenvejecido(aguae) / cuantosaguaenv
Next ' aguae


   traccionoriginal = InputBox("Ingrese la tracción original ", "Tracción Original")
    a = InStr(1, traccionoriginal, ".")
    If a <> 0 Then
    Mid(traccionoriginal, a) = ","
    End If

   traccionenvejecido = InputBox("Ingrese la tracción envejecida ", "Tracción envejecida")
    a = InStr(1, traccionenvejecido, ".")
    If a <> 0 Then
    Mid(traccionenvejecido, a) = ","
    End If

    elongacionoriginal = InputBox("Ingrese la elongación original ", "Elongación original")
    a = InStr(1, elongacionoriginal, ".")
    If a <> 0 Then
    Mid(elongacionoriginal, a) = ","
    End If


    elongacionenvejecido = InputBox("Ingrese la elongación envejecida ", "Elongación envejecida")
    a = InStr(1, elongacionenvejecido, ".")
    If a <> 0 Then
    Mid(elongacionenvejecido, a) = ","
    End If

durezaoriginal = InputBox("Ingresa la dureza original", "Dureza Original")
    a = InStr(1, durezaoriginal, ".")
    If a <> 0 Then
    Mid(durezaoriginal, a) = ","
    End If
    durezaenvejecido = InputBox("Ingresa la dureza Envejecida", "Dureza Envejecida")
    a = InStr(1, durezaenvejecido, ".")
    If a <> 0 Then
    Mid(durezaenvejecido, a) = ","
    End If
    'separacion = InputBox("Ingrese la separación de las marcas en las probetas, para elongación (mm)", "Marcas de elongación")
    pesoaireo = 0
    pesoairee = 0
    pesoaguao = 0
    pesoaguae = 0
    
    
    On Error Resume Next
For buclepromedio = 1 To numeroprobetas
    pesoaireo = pesoaireo + pesoaireoriginal(buclepromedio)
    pesoairee = pesoairee + pesoaireenvejecido(buclepromedio)
    pesoaguao = pesoaguao + pesoaguaoriginal(buclepromedio)
    pesoaguae = pesoaguae + pesoaguaenvejecido(buclepromedio)
Next 'buclepromedio
pesoaireo = pesoaireo / numeroprobetas
pesoairee = pesoairee / numeroprobetas
pesoaguao = pesoaguao / numeroprobetas
pesoaguae = pesoaguae / numeroprobetas

variacionmasa = 100 - (pesoairee * 100 / pesoaireo)
If variacionmasa >= 0 Then
    variacionmasa = "+" & variacionmasa
End If
variacionvolumen = (((pesoairee - pesoaguae) - (pesoaireo - pesoaguao)) / (pesoaireo - pesoaguao)) * 100
If variacionvolumen >= 0 Then
    variacionvolumen = "+" & variacionvolumen
End If

variaciontraccion = (traccionenvejecido * 100 / traccionoriginal) - 100
If variaciontraccion >= 0 Then
    variaciontraccion = "+" & variaciontraccion
End If

variacionelongacion = (elongacionenvejecido * 100 / elongacionoriginal) - 100
If variacionelongacion >= 0 Then
    variacionelongacion = "+" & variacionelongacion
End If

variaciondureza = durezaenvejecido - durezaoriginal
If variaciondureza >= 0 Then
    variaciondureza = "+" & variaciondureza
End If

Text2.Text = "vm: " & variacionmasa & " vv: " & variacionvolumen
Text3.Text = variaciontraccion & " %"
Text4.Text = variacionelongacion & " %"
Text5.Text = variaciondureza & " pts"

End Sub

Private Sub Command5_Click()
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT codigo, Valores, valmasa, probetas  FROM fluido_temp")
Set rs1 = db.OpenRecordset("Select codigo, N_formula, partida, tiemp_temp_oil, var_vol, var_tracc, var_elong, Var_dur, fecha_realizacion, tiempo_repeticion, aprovado  From Fluidos")
rs1.MoveLast
Codigo = rs1.RecordCount + 1
compues = InputBox("Ingrese el compuesto", "Iniciando ensayo de fluido Nº " & Codigo & " - Valores Originales")
partid = InputBox("Ingrese partida", "Codigo " & Codigo & " Compuesto " & compues)
tiemp_temp_oil = InputBox("Ingrese ensayo con el siguiente formato 24 hs 70ºC OIL3, Respetando los espacios y unidades", "Ingrese Tipo de ensayo")
probetas = InputBox("Ingrese la cantidad de probetas que utilizará", "Cantidad de probetas")
ReDim fluorig(1 To probetas)
ReDim masa(1 To probetas)



tiempo_repeticion = InputBox("Ingrese la periodicidad en días en la que tendrá que repetirse el ensayo")

For meolvide = 1 To probetas
    masa(meolvide) = InputBox("Ingrese el peso en aire " & meolvide & "/" & probetas)

    yu = InStr(1, masa(meolvide), ".")
        If yu <> 0 Then
            Mid(masa(meolvide), yu) = ","
        End If
Next ' meolvide


For compacmasa = 1 To CInt(probetas)
    If compacmasa = 1 Then
        datomasa = CStr(masa(compacmasa)) & "@"
    End If
    If compacmasa = CInt(probetas) Then
        datomasa = datomasa & CStr(masa(compacmasa))
    End If
    If compacmasa <> 1 And compacmasa <> CInt(probetas) Then
        datomasa = datomasa & CStr(masa(compacmasa)) & "@"
    End If
Next 'compacmasa


For numeroprobetas = 1 To probetas
fluorig(numeroprobetas) = 0
valorporpRob = InputBox("Ingrese la cantidad de valores de peso en agua que ha tomado para la probeta " & numeroprobetas)
    For vxp = 1 To CInt(valorporpRob) ' te quedaste aca 050811
        provis = InputBox("Ingrese peso en agua " & vxp & " de la probeta " & numeroprobetas, compues & " " & partid & " " & tiemp_temp_oil)
        a = InStr(1, provis, ".")
        If a <> 0 Then
            Mid(provis, a) = ","
        End If
        fluorig(numeroprobetas) = fluorig(numeroprobetas) + CDbl(provis)
    Next 'vxp loop de cada valor de cada probeta
    fluorig(numeroprobetas) = fluorig(numeroprobetas) / valorporpRob
Next ' Numeroprobetas' loop de cada probeta por separado
For compactar = 1 To CInt(probetas)
    If compactar = 1 Then
        dato = CStr(fluorig(compactar)) & "@"
    End If
    If compactar = CInt(probetas) Then
        dato = dato & CStr(fluorig(compactar))
    End If
    If compactar <> 1 And compactar <> CInt(probetas) Then
        dato = dato & CStr(fluorig(compactar)) & "@"
    End If
Next 'compactar

rs.AddNew
rs1.AddNew
rs.Fields("codigo") = Codigo
rs1.Fields("codigo") = Codigo

rs1.Fields("n_formula") = compues

rs1.Fields("partida") = partid

rs1.Fields("tiemp_temp_oil") = tiemp_temp_oil
rs.Fields("valores") = dato
rs.Fields("valmasa") = datomasa
rs.Fields("probetas") = probetas

rs1.Fields("tiempo_repeticion") = tiempo_repeticion

rs1.Fields("aprovado") = False
rs1.Fields("var_vol") = 0
rs1.Fields("var_tracc") = 0
rs1.Fields("fecha_realizacion") = Date
rs1.Fields("var_elong") = 0
rs1.Fields("var_dur") = 0
rs.Update
rs1.Update
db.Close
MsgBox ("Se han guardado los valores del ensayo Nº " & Codigo & " " & compues & " " & partid & " " & tiemp_temp_oil)
End Sub

Private Sub Command6_Click()
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Codigo = InputBox("Ingrese el codigo del ensayo de fluido que desea ingresar sus valores envejecidos", "Valores Envejecidos - Ensayo de Fluido")


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT codigo, Valores, valmasa, probetas  FROM fluido_temp where codigo = '" & Codigo & "'")
Set rs1 = db.OpenRecordset("Select codigo, N_formula, partida, tiemp_temp_oil, var_vol, var_tracc, var_elong, Var_dur, fecha_realizacion, tiempo_repeticion, aprovado  From Fluidos where codigo = '" & Codigo & "'")

If rs1.RecordCount = 0 Then
    sfgsf = MsgBox("No existe el ensayo indicado", vbCritical + vbOKOnly, "Error")
    db.Close
    Exit Sub
Else
    If rs.RecordCount = 0 Then
    sdfsdf = MsgBox("El ensayo ya ha sido cargado", vbCritical + vbOKOnly, "Ensayo ya caargado")
    db.Close
    Exit Sub
    End If
End If

probetas = CInt(rs.Fields("probetas"))
ReDim fluorig(1 To probetas)
ReDim masa(1 To probetas)
For meolvide = 1 To probetas
    masa(meolvide) = InputBox("Ingrese el peso en aire " & meolvide & "/" & probetas, "Envejecidos " & rs1.Fields("N_formula") & " " & rs1.Fields("partida") & " " & rs1.Fields("tiemp_temp_oil"))

    yu = InStr(1, masa(meolvide), ".")
        If yu <> 0 Then
            Mid(masa(meolvide), yu) = ","
        End If
Next ' meolvide

For numeroprobetas = 1 To probetas
fluorig(numeroprobetas) = 0
valorporpRob = InputBox("Ingrese la cantidad de valores de peso en agua que ha tomado para la probeta " & numeroprobetas)
    For vxp = 1 To CInt(valorporpRob) ' te quedaste aca 050811
        provis = InputBox("Ingrese Valor " & vxp & " de la probeta " & numeroprobetas, rs1.Fields("N_formula") & " " & rs1.Fields("partida") & " " & rs1.Fields("tiemp_temp_oil"))
        a = InStr(1, provis, ".")
        If a <> 0 Then
            Mid(provis, a) = ","
        End If
        fluorig(numeroprobetas) = fluorig(numeroprobetas) + CDbl(provis)
    Next 'vxp loop de cada valor de cada probeta
    fluorig(numeroprobetas) = fluorig(numeroprobetas) / valorporpRob
Next ' Numeroprobetas' loop de cada probeta por separado
For compactar = 1 To CInt(probetas)
    If compactar = 1 Then
        dato = CStr(fluorig(compactar)) & "@"
    End If
    If compactar = CInt(probetas) Then
        dato = dato & CStr(fluorig(compactar))
    End If
    If compactar <> 1 And compactar <> CInt(probetas) Then
        dato = dato & CStr(fluorig(compactar)) & "@"
    End If
Next 'compactar

cantidadarrovas = probetas - 1

ReDim originales(1 To probetas)
ReDim envejecidos(1 To probetas)
ReDim arr(1 To cantidadarrovas) 'este guarda las posiciones de las arrova en el string "dato"



For buscandoarova = 1 To cantidadarrovas
    If buscandoarova = 1 Then
        arr(buscandoarova) = InStr(1, dato, "@")
    Else
        arr(buscandoarova) = InStr((arr(buscandoarova - 1) + 1), dato, "@")
    End If
Next 'buscandoarova


For sacar = 1 To probetas
    If sacar = 1 Then
        envejecidos(sacar) = Mid(dato, 1, arr(sacar) - 1)
    End If
    If sacar <> 1 And sacar <> probetas Then
        envejecidos(sacar) = Mid(dato, arr(sacar - 1) + 1, arr(sacar) - (arr(sacar - 1) + 1))
    End If
    If sacar = probetas Then
        fg = Len(dato)
        envejecidos(sacar) = Mid(dato, arr(sacar - 1) + 1, fg - arr(sacar - 1))
    End If
Next 'sacar

For buscandoarova = 1 To cantidadarrovas
    If buscandoarova = 1 Then
        arr(buscandoarova) = InStr(1, rs.Fields("valores"), "@")
    Else
        arr(buscandoarova) = InStr((arr(buscandoarova - 1) + 1), rs.Fields("valores"), "@")
    End If
Next 'buscandoarova



For sacaro = 1 To probetas
    If sacaro = 1 Then
        originales(sacaro) = Mid(rs.Fields("valores"), 1, arr(sacaro) - 1)
    End If
    If sacaro <> 1 And sacaro <> probetas Then
        originales(sacaro) = Mid(rs.Fields("valores"), arr(sacaro - 1) + 1, arr(sacaro) - (arr(sacaro - 1) + 1))
    End If
    If sacaro = probetas Then
        fg = Len(dato)
        originales(sacaro) = Mid(rs.Fields("valores"), arr(sacaro - 1) + 1, fg - arr(sacaro - 1))
    End If
Next 'sacaro
    
    ReDim masaor(1 To probetas)
    
    For buscandoarova = 1 To cantidadarrovas
    If buscandoarova = 1 Then
        arr(buscandoarova) = InStr(1, rs.Fields("valmasa"), "@")
    Else
        arr(buscandoarova) = InStr((arr(buscandoarova - 1) + 1), rs.Fields("valmasa"), "@")
    End If
Next 'buscandoarova
    
    
    
    For sacarom = 1 To probetas
    If sacarom = 1 Then
        masaor(sacarom) = Mid(rs.Fields("valmasa"), 1, arr(sacarom) - 1)
    End If
    If sacarom <> 1 And sacarom <> probetas Then
        masaor(sacarom) = Mid(rs.Fields("valmasa"), arr(sacarom - 1) + 1, arr(sacarom) - (arr(sacarom - 1) + 1))
    End If
    If sacarom = probetas Then
        fg = Len(rs.Fields("valmasa"))
        masaor(sacarom) = Mid(rs.Fields("valmasa"), arr(sacarom - 1) + 1, fg - arr(sacarom - 1))
    End If
Next 'sacarom
    
    mo = 0
    For vamos = 1 To probetas
    mo = mo + CSng(masaor(vamos))
    Next 'vamos
    mo = mo / probetas
    
    mae = 0
    For vamos = 1 To probetas
    mae = mae + CSng(masa(vamos))
    Next 'vamos
    mae = mae / probetas
    
    vo = 0
    For vamos = 1 To probetas
    vo = vo + CSng(originales(vamos))
    Next 'vamos
    vo = vo / probetas
    
    ve = 0
    For vamos = 1 To probetas
    ve = ve + CSng(envejecidos(vamos))
    Next 'vamos
    ve = ve / probetas
    
    vv = (((mae - ve) - (mo - vo)) / (mo - vo)) * 100
    
    Text2.Text = "vv= " & vv
    Label13.Caption = Codigo
    Combo1.Text = rs1.Fields("N_formula")
    Text1.Text = rs1.Fields("Partida")
    Text6.Text = rs1.Fields("tiempo_repeticion")
db.Close
End Sub

Private Sub Command7_Click()
frmBuscaTraccion.Show
frmBuscaTraccion.Height = 5565
frmBuscaTraccion.Command2.Visible = False
frmBuscaTraccion.Command5.Visible = True
End Sub

Private Sub Command8_Click()
frmPendFlu.Show
FrmFluidoNuevo.Enabled = False
FrmFluidoNuevo.Visible = False
frmPendFlu.MSFlexGrid1.Clear
frmPendFlu.MSFlexGrid1.TextMatrix(0, 0) = "Codigo"
frmPendFlu.MSFlexGrid1.TextMatrix(0, 1) = "P.Agua"
frmPendFlu.MSFlexGrid1.TextMatrix(0, 2) = "P.Aire"
frmPendFlu.MSFlexGrid1.TextMatrix(0, 3) = "Cant.Prob."
frmPendFlu.MSFlexGrid1.TextMatrix(0, 4) = "Compuesto"
frmPendFlu.MSFlexGrid1.TextMatrix(0, 5) = "Partida"
frmPendFlu.MSFlexGrid1.TextMatrix(0, 6) = "Ensayo"

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT codigo, Valores, valmasa, probetas  FROM fluido_temp")

If rs.RecordCount = 0 Then
dsfgsdgsa = MsgBox("No hay pendientes", vbInformation + vbOKOnly, "Sin registros")
frmPendFlu.Command1.SetFocus
Exit Sub
Else
rs.MoveLast
frmPendFlu.MSFlexGrid1.Rows = rs.RecordCount + 1
rs.MoveFirst
For bubu = 1 To rs.RecordCount
frmPendFlu.MSFlexGrid1.TextMatrix(bubu, 0) = rs.Fields("codigo")
Set rs1 = db.OpenRecordset("Select codigo, N_Formula, Partida, Tiemp_temp_oil From Fluidos Where  codigo = '" & rs.Fields("codigo") & "'")
frmPendFlu.MSFlexGrid1.TextMatrix(bubu, 4) = rs1.Fields("N_Formula")
frmPendFlu.MSFlexGrid1.TextMatrix(bubu, 5) = rs1.Fields("Partida")
frmPendFlu.MSFlexGrid1.TextMatrix(bubu, 6) = rs1.Fields("Tiemp_temp_oil")

frmPendFlu.MSFlexGrid1.TextMatrix(bubu, 1) = rs.Fields("Valores")
frmPendFlu.MSFlexGrid1.TextMatrix(bubu, 2) = rs.Fields("valmasa")
frmPendFlu.MSFlexGrid1.TextMatrix(bubu, 3) = rs.Fields("probetas")
rs.MoveNext
Next
End If
End Sub

Private Sub Text2_Change()
Label1.Caption = Date
End Sub
