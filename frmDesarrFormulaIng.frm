VERSION 5.00
Begin VB.Form frmDesarrFormulaIng 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Formula Desarrollo"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Mail a Producción"
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Quitar Item"
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Agregar Item"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar Como"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar Valores"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      ItemData        =   "frmDesarrFormulaIng.frx":0000
      Left            =   3720
      List            =   "frmDesarrFormulaIng.frx":000A
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Partes Totales"
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Densidad con acelerado"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Precio con acelerado"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   7320
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Formula"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Etapa"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Partes"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmDesarrFormulaIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 
 gi = Combo1.Count
 For i = 1 To gi - 1
    Unload Combo1(i)
    Unload Combo2(i)
    Unload Text1(i)
 Next
 Form1.Enabled = True
 Form1.Visible = True
 Me.Hide
End Sub

Private Sub Command2_Click()
    
    
    
    
    Dim db As Database
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, False, ";pwd=flanflus")
    
    
    Set rs = db.OpenRecordset("Select * from partes_desarrollo where n_formula = '" & Label5.Caption & "'")
    
    If rs.RecordCount <> 0 Then
        asdasd = MsgBox("El compuesto que está por guardar ya existe. Desea reemplazarlo?", vbCritical + vbYesNo, "Sobreescribir?")
        If asdasd = vbYes Then
        
            Do Until rs.EOF = True
                rs.Delete
                rs.MoveNext
            Loop
        Else
            Exit Sub
        End If
    End If
    
    
    
    Set rs = db.OpenRecordset("Select * from partes_desarrollo")
    
    componentes = Combo1.Count
    For k = 0 To componentes - 1
        If Combo1(k).Text = "" Then
            sdfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
            Exit Sub
        End If
        If Combo2(k).Text = "" Then
            sdfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
            Exit Sub
        End If
        If Text1(k).Text = "" Then
            sdfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
            Exit Sub
        End If
    Next
    For k = 0 To componentes - 1
        Text1(k).Text = Replace(Text1(k).Text, ".", ",")
    Next
    
    
    
    
    
    For i = 0 To componentes - 1
        rs.AddNew
        rs.Fields("N_FORMULA") = Label5.Caption
        
        Set rs1 = db.OpenRecordset("Select COD_PROD from producto where descrip = '" & Combo1(i).Text & "'")
        rs.Fields("cod_prod") = rs1.Fields("cod_prod")
        rs.Fields("partes") = punto_por_coma(Text1(i).Text)
        rs.Fields("etapa") = Combo2(i).Text
        rs.Update
    Next
    
    db.Close
    sdfsdf = MsgBox("Se ha guardado satisfactoriamente la formula", vbInformation + vbOKOnly, "Grabado de formula de desarrollo")
    Form1.Enabled = True
    Form1.Visible = True
     gi = Combo1.Count
 For i = 1 To gi - 1
    Unload Combo1(i)
    Unload Combo2(i)
    Unload Text1(i)
 Next
    If Check1.Value = 1 Then
        '''''''''Envia mail a producción
        ReDim destinatarios(1 To 2)
        indicedestinatarios = 2
        asunto = "Entorno Bafir: Solicitud de Mezclado"
        mail = ": Automensaje: Se solicita el mezclado del compuesto de desarrollo   " & Label5.Caption & ".  El mismo ya se encuentra cargado en el sistema. Muchas Gracias."
        destinatarios(1) = "vflor@bafir.com.ar"
        destinatarios(2) = "pablopirri@bafir.com.ar"
        frmSendinfo.Show
        frmSendinfo.Visible = True
        Call Moduloenvio
        frmSendinfo.Hide
    End If
    
    Me.Hide
End Sub

Private Sub Command3_Click()

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")

materias = Combo1.Count

For k = 0 To materias - 1
    Text1(k).Text = Replace(Text1(k).Text, ".", ",")
Next



preCio = 0
densIdad = 0
partesTot = 0
For i = 0 To materias - 1
    Set rs = db.OpenRecordset("Select precio, pesoesp from producto where descrip = '" & Combo1(i).Text & "'")
    If rs.RecordCount = 0 Then
        asdasdasd = MsgBox("La materia prima " & Combo1(i).Text & " es incorrecta", vbCritical + vbOKOnly, "Error")
        db.Close
        Exit Sub
    End If
    adsfdfd = rs.Fields("pesoesp")
    If IsNull(rs.Fields("pesoesp")) Then
        asdadas = MsgBox("No están cargados los datos de precio o densidad del compuesto", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
partesTot = partesTot + CDbl(Text1(i).Text)
preCio = preCio + CDbl(rs.Fields("precio")) * CDbl(Text1(i).Text)
densIdad = densIdad + CDbl(Text1(i).Text) / CDbl(rs.Fields("pesoesp"))

Next
Label7.Caption = partesTot
Label12.Caption = Format(preCio / partesTot, "0.000") & " u$s"
Label13.Caption = Format(partesTot / densIdad, "0.000") & " g/ml"
db.Close
End Sub

Private Sub Command4_Click()
    NombreA = InputBox("Ingrese el código con el cual quiere guardar esta modificación", "Nuevo Compuesto")
    
    Dim db As Database
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, False, ";pwd=flanflus")
    
    
    Set rs = db.OpenRecordset("Select * from partes_desarrollo where n_formula = '" & NombreA & "'")
    
    If rs.RecordCount <> 0 Then
        asdasd = MsgBox("El compuesto que está por guardar ya existe. Desea reemplazarlo?", vbCritical + vbYesNo, "Sobreescribir?")
        If asdasd = vbYes Then
        
            Do Until rs.EOF = True
                rs.Delete
                rs.MoveNext
            Loop
        Else
            Exit Sub
        End If
    End If
    
    
    
    Set rs = db.OpenRecordset("Select * from partes_desarrollo")
    
    componentes = Combo1.Count
    For k = 0 To componentes - 1
        Text1(k).Text = Replace(Text1(k).Text, ".", ",")
    Next
      
    
    For i = 0 To componentes - 1
        rs.AddNew
        rs.Fields("N_FORMULA") = NombreA
        
        Set rs1 = db.OpenRecordset("Select COD_PROD from producto where descrip = '" & Combo1(i).Text & "'")
        rs.Fields("cod_prod") = rs1.Fields("cod_prod")
        rs.Fields("partes") = punto_por_coma(Text1(i).Text)
        rs.Fields("etapa") = Combo2(i).Text
        rs.Update
    Next
    
    db.Close
    sdfsdf = MsgBox("Se ha guardado satisfactoriamente la formula", vbInformation + vbOKOnly, "Grabado de formula de desarrollo")
    Form1.Enabled = True
    Form1.Visible = True
     gi = Combo1.Count
 For i = 1 To gi - 1
    Unload Combo1(i)
    Unload Combo2(i)
    Unload Text1(i)
 Next
    Me.Hide
End Sub

Private Sub Command5_Click()
cuanTos = Text1.Count
cuanTos = cuanTos - 1
Load Combo1(cuanTos + 1)
Load Text1(cuanTos + 1)
Load Combo2(cuanTos + 1)
Combo1(cuanTos + 1).Top = Combo1(cuanTos).Top + Combo1(cuanTos).Height
Text1(cuanTos + 1).Top = Text1(cuanTos).Top + Text1(cuanTos).Height
Combo2(cuanTos + 1).Top = Combo2(cuanTos).Top + Combo2(cuanTos).Height
Combo1(cuanTos + 1).Visible = True
Text1(cuanTos + 1).Visible = True
Combo2(cuanTos + 1).Visible = True
IniCsEt = (cuanTos + 1) * 3
Combo1(cuanTos + 1).TabIndex = IniCsEt
Text1(cuanTos + 1).TabIndex = IniCsEt + 1
Combo2(cuanTos + 1).TabIndex = IniCsEt + 2

    Dim db As Database
    Dim rs As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")
    Set rs = db.OpenRecordset("Select cod_prod,descrip from producto order by descrip")
     
    Do Until rs.EOF = True
            frmDesarrFormulaIng.Combo1(cuanTos + 1).AddItem (rs.Fields("descrip"))
            rs.MoveNext
    Loop
    db.Close
    
    Combo2(cuanTos + 1).AddItem ("A")
    Combo2(cuanTos + 1).AddItem ("B")
    Combo1(cuanTos + 1).Text = ""
    Combo2(cuanTos + 1).Text = ""
    Text1(cuanTos + 1).Text = ""

End Sub

Private Sub Command6_Click()
iTm = Text1.Count - 1
If iTm = 0 Then
    sdfsdf = MsgBox("No se puede quitar este item por ser el primero", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Unload Combo1(iTm)
Unload Text1(iTm)
Unload Combo2(iTm)
End Sub

Private Sub Command7_Click()
Dim appp As New Excel.Application
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim r As Excel.Range

filos = Text1.Count

Set wb = appp.Workbooks.Open("\\Servidor2\e\EntornoBafir\Planillas\informedesarr.xls", , True)
Set ws = wb.Worksheets(1)
ws.Cells(3, 2) = Label5.Caption
For i = 1 To filos
'ws.Cells(7, 1).Select

    With appp
        .Range("A7", "B7").Select
        .Selection.Insert Shift:=xlDown
    End With
Next
For i = 0 To filos - 1
    ws.Cells(6 + i, 1) = Combo1(i).Text
    ws.Cells(6 + i, 2) = Text1(i).Text
Next

ws.Cells(6 + filos + 1, 2) = Label7.Caption
ws.Cells(6 + filos + 3, 2) = Label12.Caption
ws.Cells(6 + filos + 4, 2) = Label13.Caption



On Error Resume Next

ws.PrintOut

If Err.Number = 1004 Then
    dfsdf = MsgBox("Hay un problema con la impresora, la formula no se ha impreso", vbCritical + vbOKOnly, "Error")
Else
    DoEvents
    sdffsf = MsgBox("Presiones Ok cuando la impresión esté finalizada", vbInformation + vbOKOnly, "Imprimiendo")
End If
wb.Close (False)
appp.Quit
End Sub

