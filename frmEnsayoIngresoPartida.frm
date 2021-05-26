VERSION 5.00
Begin VB.Form frmEnsayoIngresoPartida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Ensayo de partida"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   12360
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   495
      Left            =   12360
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Traer ensayos"
      Height          =   495
      Left            =   12360
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   495
      Left            =   12360
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label probeta 
      Caption         =   "probeta"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Unidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Maximo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Mínimo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Medicion 
      Caption         =   "Medicion"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Toleranciamax 
      Caption         =   "ToleranciaMax"
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Toleranciamin 
      Caption         =   "Toleranciamin"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label unidades 
      Caption         =   "Unidad"
      Height          =   255
      Index           =   0
      Left            =   11280
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frmEnsayoIngresoPartida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
sPathBase1 = "\\Servidor2\e\EntornoBafir\Normas\" & Normatxt & ".mdb"


    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   With cnn1
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase1 & ";" & "Jet OLEDB:Database Password=zZz999;"
        .Open
    End With
    
        '''''''''
'crear la columna con la partida en tolerancias

'ADOCon.Execute "ALTER TABLE tblTest" _
'         & i & " ADD Field" & j & " Text"
'ALTER TABLE table ADD COLUMN field type [(size)] [NOT NULL] [CONSTRAINT constraint]


If Not ExisteCampo(partidaA, "\\Servidor2\e\EntornoBafir\Normas\" & Normatxt & ".mdb", "Tolerancias") Then
    'cnn1.Execute "ALTER TABLE Tolerancias ADD `" & frmModPartidasNuevas.parTida & "` varchar(50)"
    cnn1.Execute "ALTER TABLE Tolerancias ADD `" & partidaA & "` varchar(50) NULL"
End If




'cantidad de ensayos
'a = UBound(frmModPartidasNuevas.mediciones_DE_ensayo)
a = UBound(mediciones_DE_ensayo)

'esto es por cada ensayo
k = 1
For i = 1 To a
    'b es cantidad de mediciones para a
    'b = frmModPartidasNuevas.mediciones_DE_ensayo(i)
    b = mediciones_DE_ensayo(i)
    Ensayo = Label1(i)
    rst.Open "SELECT codigo FROM ensayos where referencia = '" & Ensayo & "'", cnn, adOpenStatic, adLockReadOnly
    ensayocod = rst.Fields("codigo")
    rst.Close
    
    For Y = 1 To b
        'esto carga los valores de los ensayos en la tabla tolerancias
        rMedicion = Medicion(k)
        rst.Open "SELECT codigo FROM mediciones where mediciones = '" & rMedicion & "'", cnn, adOpenStatic, adLockReadOnly
        medicioncod = rst.Fields("codigo")
        rst.Close
        rst1.Open "SELECT `" & partidaA & "` FROM tolerancias where mediciones = " & medicioncod & " and ensayo = " & ensayocod, cnn1, adOpenStatic, adLockOptimistic
        rst1.Fields(0) = Text1(k).Text
        rst1.Update
        rst1.Close
        k = k + 1
    Next
Next
asdasd = MsgBox("Se han cargado los datos para la partida " & parTida & ".", vbInformation + vbOKOnly, "Carga de partida")
cnn.Close
cnn1.Close
End Sub

Private Sub Command2_Click()
observacionesflag = 0
cantensayos = UBound(mediciones_DE_ensayo)
controles = 1
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
With cnn
     'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
     .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
     .Open
 End With
For fortraer = 1 To cantensayos
    cantmediCiones = mediciones_DE_ensayo(fortraer)
    'a = Label1(fortraer).Caption
    If Label1(fortraer).Caption = "Original" Then
        
        For formediciones = controles To controles + cantmediCiones - 1
            'a = Medicion(formediciones).Caption
            If Medicion(formediciones).Caption = "Traccion" Or Medicion(formediciones).Caption = "Elongacion" Or Medicion(formediciones).Caption = "Dureza" Then
                rst.Open "SELECT `" & Medicion(formediciones) & "`, traccion1, traccion2, elongacion1, elongacion2, observacion FROM traccion where referencia = '0' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                If rst.RecordCount <> 0 Then
                    If Medicion(formediciones).Caption = "Traccion" Then
                        If rst.Fields("traccion1") = 0 Then
                            Text1(formediciones).Text = Format(rst.Fields(Medicion(formediciones).Caption), "0.00")
                        ElseIf IsNull(rst.Fields("traccion1")) Then
                            Text1(formediciones).Text = Format(rst.Fields(Medicion(formediciones).Caption), "0.00")
                        Else
                            Text1(formediciones).Text = Format((CDbl(rst.Fields("traccion")) + CDbl(rst.Fields("traccion1")) + CDbl(rst.Fields("traccion2"))) / 3, "0.00")
                        End If
                    ElseIf Medicion(formediciones).Caption = "Elongacion" Then
'''''''''''
                        If rst.Fields("elongacion1") = 0 Then
                            Text1(formediciones).Text = Format(rst.Fields(Medicion(formediciones).Caption), "0.00")
                        ElseIf IsNull(rst.Fields("elongacion1")) Then
                            Text1(formediciones).Text = Format(rst.Fields(Medicion(formediciones).Caption), "0.00")
                        Else
                            Text1(formediciones).Text = Format((CDbl(rst.Fields("elongacion")) + CDbl(rst.Fields("elongacion1")) + CDbl(rst.Fields("elongacion2"))) / 3, "0.00")
                        End If
''''''''''''''
                    Else
                    Text1(formediciones).Text = Format(rst.Fields(Medicion(formediciones).Caption), "0.00")
                    End If
                End If
                If rst.Fields("observacion") <> 0 Or Not IsNull(rst.Fields("observacion")) Then
                    observacionesflag = 1
                End If
                rst.Close
                                
            Else 'si no es o traccion o elongacion o dureza ( por ej. desgarro )
                ''desgarro
                If Medicion(formediciones).Caption = "Desgarro" Then
                    
                    'rst.Open "SELECT `" & Medicion(formediciones) & "` FROM desgarros where partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    rst.Open "SELECT valor FROM desgarros where partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    If rst.RecordCount <> 0 Then
                        Text1(formediciones).Text = rst.Fields("valor")
                    End If
                    rst.Close
                End If
            End If
        
        Next 'formediciones
        controles = controles + cantmediCiones
    ElseIf Left(Label1(fortraer).Caption, 6) = "Compre" Then 'compresion
        
        For formediciones = controles To controles + cantmediCiones - 1
            
            rst.Open "SELECT compresion_porc FROM compresion where partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' and tiempo_temperatura = '" & Label1(fortraer).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                If rst.RecordCount <> 0 Then
                    Text1(formediciones).Text = rst.Fields("compresion_porc")
                End If
                rst.Close
        Next 'formediciones
        controles = controles + cantmediCiones
    'envejecimientos (aire, fluido, externos)
    ElseIf Left(Label1(fortraer).Caption, 6) = "Enveje" Then 'compresion
        'para envej aire
        ''If Right(Label1(fortraer).Caption, 4) = "Aire" Then
            For formediciones = controles To controles + cantmediCiones - 1

                'para valores solos
                If Medicion(formediciones).Caption = "Traccion" Or Medicion(formediciones).Caption = "Elongacion" Or Medicion(formediciones).Caption = "Dureza" Then
                    rst.Open "SELECT `" & Medicion(formediciones) & "`, traccion1, traccion2, elongacion1, elongacion2, observacion FROM traccion where referencia = '0' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    If rst.RecordCount <> 0 Then
                        Text1(formediciones).Text = rst.Fields(Medicion(formediciones).Caption)
                    End If
                If rst.Fields("observacion") <> 0 Or Not IsNull(rst.Fields("observacion")) Then
                    observacionesflag = 1
                End If
                    
                    
                    rst.Close
                End If
                ' para variacion
                ' var. traccion
                If Medicion(formediciones).Caption = "Var. Traccion" Then
                    'rst original
                    'rst1 envejecido
                    
                    rst.Open "SELECT traccion, traccion1, traccion2, observacion FROM traccion where referencia = '0' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    rst1.Open "SELECT traccion, traccion1, traccion2, observacion FROM traccion where referencia = '" & Label1(fortraer).Caption & "' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    If rst.RecordCount <> 0 And rst1.RecordCount <> 0 Then
                        If rst.Fields("traccion1") = 0 Then
                            Text1(formediciones).Text = Format((CDbl(rst1.Fields("traccion")) * 100 / rst.Fields("traccion")) - 100, "0.00")
                        ElseIf IsNull(rst.Fields("traccion1")) Then
                            Text1(formediciones).Text = Format((CDbl(rst1.Fields("traccion")) * 100 / rst.Fields("traccion")) - 100, "0.00")
                        Else
                            
                            t2 = rst1.Fields("traccion1")
                            t3 = rst1.Fields("traccion2")
                            If IsNull(t2) Then
                                t2 = rst1.Fields("traccion")
                            End If
                            If IsNull(t3) Then
                                t3 = rst1.Fields("traccion")
                            End If
                            
                            'o = Format(((CDbl(rst1.Fields("traccion")) + CDbl(t2) + CDbl(t3)) / 3) * 100 / ((CDbl(rst.Fields("traccion")) + CDbl(rst.Fields("traccion1")) + CDbl(rst.Fields("traccion2"))) / 3) - 100, "0.00")
                            
                            Text1(formediciones).Text = Format(((CDbl(rst1.Fields("traccion")) + CDbl(t2) + CDbl(t3)) / 3) * 100 / ((CDbl(rst.Fields("traccion")) + CDbl(rst.Fields("traccion1")) + CDbl(rst.Fields("traccion2"))) / 3) - 100, "0.00")
                        End If
                        
                    End If
                    If rst.Fields("observacion") <> 0 Or Not IsNull(rst.Fields("observacion")) Then
                        observacionesflag = 1
                    End If
                    
                    rst.Close
                    
                    
                    rst1.Close
                End If
                'var. traccion
                
                'var. elongacion
                If Medicion(formediciones).Caption = "Var. Elongacion" Then
                    rst.Open "SELECT elongacion, elongacion1, elongacion2 FROM traccion where referencia = '0' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    rst1.Open "SELECT elongacion, elongacion1, elongacion2 FROM traccion where referencia = '" & Label1(fortraer).Caption & "' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    If rst.RecordCount <> 0 And rst1.RecordCount <> 0 Then
                        If rst.Fields("elongacion1") = 0 Then
                            Text1(formediciones).Text = Format((CDbl(rst1.Fields("elongacion")) * 100 / rst.Fields("elongacion")) - 100, "0.00")
                        ElseIf IsNull(rst.Fields("elongacion1")) Then
                            Text1(formediciones).Text = Format((CDbl(rst1.Fields("elongacion")) * 100 / rst.Fields("elongacion")) - 100, "0.00")
                        Else
                            el2 = rst1.Fields("elongacion1")
                            el3 = rst1.Fields("elongacion2")
                            If IsNull(el2) Then
                                el2 = rst1.Fields("elongacion")
                            End If
                            If IsNull(el3) Then
                                el3 = rst1.Fields("elongacion")
                            End If
                        
                        
                        
                        
                        
                            Text1(formediciones).Text = Format((((CDbl(rst1.Fields("elongacion")) + CDbl(el2) + CDbl(el3)) / 3) * 100 / ((CDbl(rst.Fields("elongacion")) + CDbl(rst.Fields("elongacion1")) + CDbl(rst.Fields("elongacion2"))) / 3)) - 100, "0.00")
                        End If
                    End If
                    rst.Close
                    rst1.Close
                End If
                'var. elongacion
                
                'var dureza
                If Medicion(formediciones).Caption = "Var. Dureza" Then
                    rst.Open "SELECT dureza FROM traccion where referencia = '0' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    rst1.Open "SELECT dureza FROM traccion where referencia = '" & Label1(fortraer).Caption & "' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc", cnn, adOpenStatic, adLockReadOnly
                    If rst.RecordCount <> 0 And rst1.RecordCount <> 0 Then
                        Text1(formediciones).Text = Format((CDbl(rst1.Fields("dureza")) - rst.Fields("dureza")), "0.00")
                    End If
                    rst.Close
                    rst1.Close
                End If
                'var dureza
                
                'var volumen
                If Medicion(formediciones).Caption = "Var. Volumen" Then
                    'asd = MsgBox("SELECT var_vol FROM fluidos where tiemp_temp_oil = '" & Label1(fortraer).Caption & "' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo_ensayo desc")
                    rst.Open "SELECT var_vol, var_vol1, var_vol2 FROM fluidos where tiemp_temp_oil = '" & Label1(fortraer).Caption & "' and partida = '" & partidaA & "' and probeta = '" & probeta(formediciones).Caption & "' order by codigo desc", cnn, adOpenStatic, adLockReadOnly
                    If rst.RecordCount <> 0 Then
                        ''''''''''''''''''mod 080125
                        If IsNull(rst.Fields("var_vol1")) Then
                            Text1(formediciones).Text = Format(rst.Fields("var_vol"), "0.00")
                        Else
                            Text1(formediciones).Text = Format((CDbl(rst.Fields("var_vol")) + CDbl(rst.Fields("var_vol1")) + CDbl(rst.Fields("var_vol2"))) / 3, "0.00")
                        End If
                        ''''''''''''''''''mod 080125
                        
                    End If
                    rst.Close
                End If
                'var volumen

            Next 'formediciones
            controles = controles + cantmediCiones
        ''End If 'aire
    
    
    
    Else
        controles = controles + cantmediCiones
    
    End If 'todo


Next 'fortraer
cnn.Close
'evaluación

controles = Text1.Count
For i = 1 To controles - 1

    If Text1(i).Text <> "" Then
        OKEnsayo = True
        If frmEnsayoIngresoPartida.Toleranciamin(i).Caption <> "" Then
            If CDbl(Text1(i).Text) < CDbl(frmEnsayoIngresoPartida.Toleranciamin(i).Caption) Then
                OKEnsayo = False
            End If
        End If
        If frmEnsayoIngresoPartida.Toleranciamax(i).Caption <> "" Then
            If CDbl(Text1(i).Text) > CDbl(frmEnsayoIngresoPartida.Toleranciamax(i).Caption) Then
                OKEnsayo = False
            End If
        End If
    
        If OKEnsayo = False Then
            frmEnsayoIngresoPartida.Text1(i).BackColor = vbRed
        Else
            frmEnsayoIngresoPartida.Text1(i).BackColor = vbGreen
        End If
    
    End If
Next i









If observacionesflag = 1 Then
    sdfsdf = MsgBox("Atención, alguno de los ensayos presentan observaciones. Verificar antes de realizar la aprobación.", vbInformation + vbOKOnly, "Se han encontrado observaciones")
Else
    sdfsdf = MsgBox("Sin observaciones", vbInformation + vbOKOnly, "No se han encontrado observaciones")
End If


End Sub

Private Sub Command3_Click()
labelcount = Label1.Count
medicioncount = Medicion.Count
toleranciamincount = Toleranciamin.Count
Toleranciamaxcount = Toleranciamax.Count
text1count = Text1.Count
For i = 1 To labelcount - 1
    Unload Label1(i)
Next
For i = 1 To medicioncount - 1
    Unload Medicion(i)
    Unload probeta(i)
    Unload unidades(i)
Next
For i = 1 To toleranciamincount - 1
    Unload Toleranciamin(i)
Next
For i = 1 To Toleranciamaxcount - 1
    Unload Toleranciamax(i)
Next
For i = 1 To text1count - 1
    Unload Text1(i)
Next


Dim tmpform As Form
Dim tform As Form
        
For Each tmpform In Forms
    If tmpform.Name = formUlario Then
        Set tform = tmpform
    End If
Next
If tform.Enabled = False Then
    tform.Enabled = True
End If
If tform.Visible = False Then
    tform.Visible = True
End If
Me.Hide
End Sub

Private Sub Command4_Click()
Me.PrintForm
End Sub

Private Sub Label5_DblClick()
    Command2.Enabled = True
End Sub
