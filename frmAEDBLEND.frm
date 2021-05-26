VERSION 5.00
Begin VB.Form frmAEDBLEND 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blend AED"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "?"
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   5520
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Controlado?"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   4680
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "No aprobadas"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aprobadas"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Componentes"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label label1 
      Caption         =   "Batches"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmAEDBLEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mOdO As Integer
Public baTCh As Long
'modos
'0 no activar autoseleccion en list1
'1 activar autoseleccion en list1

Private Sub Command1_Click()
List2.AddItem (List1.Text)
End Sub

Private Sub Command2_Click()
If List2.ListIndex <> -1 Then
    List2.RemoveItem (List2.ListIndex)
End If
End Sub

Private Sub Command3_Click()
mOdO = 0
nuevobatch = InputBox("Ingrese el número del nuevo batch blend", "Crear Blend")
If nuevobatch = "" Then
    Exit Sub
End If
baTCh = nuevobatch
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = True
Check1.Enabled = True
Check1.Value = 0
List1.AddItem (nuevobatch)
List1.ListIndex = (List1.ListCount - 1)
List2.Clear
End Sub

Private Sub Command4_Click()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Command5.Enabled = True
Check1.Enabled = True
mOdO = 0
baTCh = List1.Text
End Sub

Private Sub Command5_Click()
If List2.ListCount = 0 Then
    asdasd = MsgBox("Debe agregar algun batch al blend", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If List2.ListCount = 1 Then
    asdasdasd = MsgBox("Debe de haber al menos 2 batches para constituir un blend.", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
cantidad_de_arrays = List2.ListCount
For i = 0 To (cantidad_de_arrays - 1)
    If i = 0 Then
        List2.ListIndex = 0
        batches = List2.Text & "@"
    End If
    If i <> 0 And i <> (cantidad_de_arrays - 1) Then
        List2.ListIndex = i
        batches = batches & List2.Text & "@"
    End If
    If i = (cantidad_de_arrays - 1) Then
        List2.ListIndex = (cantidad_de_arrays - 1)
        batches = batches & List2.Text
    End If
Next

Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    '''''busca que todos los constituyentes sean del mismo compuesto, sino aborta
    tooo = List2.ListCount
    ReDim compuestoB(0 To (tooo - 1))
    
    For i = 0 To (tooo - 1)
        List2.ListIndex = i
        rst1.Open "SELECT compuesto FROM mezclas where batch = '" & List2.Text & "'", cnn, adOpenStatic, adLockReadOnly
        compuestoB(i) = rst1.Fields("compuesto")
        rst1.Close
    Next
    
    cantidad = UBound(compuestoB)
    For i = 0 To cantidad
        If i = 0 Then
            If compuestoB(0) <> compuestoB(1) Then
                sdfsdf = MsgBox("Los batches constituyentes no pertenecen al mismo compuesto, no se puede continuar", vbCritical + vbOKOnly, "Error")
                Exit Sub
            End If
        End If
        If i <> 0 And i <> cantidad Then
            If compuestoB(i) <> compuestoB(i + 1) Then
                sdfsdf = MsgBox("Los batches constituyentes no pertenecen al mismo compuesto, no se puede continuar", vbCritical + vbOKOnly, "Error")
                Exit Sub
            End If
        End If
        If i = cantidad Then
            'nada por que ya terminaste pelotudo
        End If
        compuestodefinitivo = compuestoB(0)
    Next
          
    '''''
    
    rst1.Open "SELECT * FROM mezclas where batch = '" & baTCh & "'", cnn, adOpenStatic, adLockOptimistic
    If rst1.RecordCount = 0 Then
        rst1.AddNew
        rst1.Fields("fecha") = Format(Now, "MM/DD/YY")
        rst1.Fields("batch") = baTCh
        rst1.Fields("Compuesto") = compuestodefinitivo
        If Option1.Value = True Then
            aprobado = True
            rst1.Fields("fecha_aprobado") = Format(Now, "MM/DD/YY")
        Else
            aprobado = False
        End If
        rst1.Fields("aprobado") = aprobado
        
        rst1.Fields("blend") = True
        rst1.Fields("controlado") = Check1.Value
        rst1.Update
    Else
        'Si existe por ahora no hace nada, ya que de seguro es una modificacion
            
        
        
    End If
    rst.Open "SELECT * FROM blend where batch_final = '" & baTCh & "'", cnn, adOpenStatic, adLockOptimistic
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("batch_final") = baTCh
        rst.Fields("batches") = batches
        rst.Update
    Else
        rst.Fields("batches") = batches
        rst.Fields("controlado") = Check1.Value
        rst.Update
    End If
    rst.Close

cnn.Close
listacuenta = List1.ListCount
List1.ListIndex = List1.ListCount - 1
Command3.Enabled = True
Command4.Enabled = False
Command5.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
mOdO = 1
End Sub

Private Sub Command6_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Command7_Click()
If List1.Text = "" Then
    Exit Sub
End If
Form1.AEDPassword = ""
frmAEDPassword.Show (1)
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

rst.Open "SELECT dato FROM Reg WHERE funcion = 'masterkey'", cnn, adOpenStatic, adLockReadOnly
If rst.Fields("dato") <> Form1.AEDPassword Then
    cnn.Close
    Exit Sub
End If
cnn.Close

frmAEDDOC.formUlario = Me.Name
frmAEDDOC.AEDstr = "Select Observaciones From blend where batch_final = '" & List1.Text & "'"
Me.Enabled = False
frmAEDDOC.Show
End Sub

Private Sub List1_Click()
If mOdO <> 0 Then 'ver en public los modos
    List2.Clear
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rst1 As ADODB.Recordset
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst1 = New ADODB.Recordset
    
    sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"
    
        With cnn
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
            .Open
        End With
        
        rst.Open "SELECT * FROM blend where batch_final = '" & List1.Text & "'", cnn, adOpenStatic, adLockReadOnly
        constituyentes = Explode("@", rst.Fields("batches"))
        Check1.Value = Abs(CInt(rst.Fields("controlado")))
        tamaño = UBound(constituyentes) + 1
        If UBound(constituyentes) = 0 Then
            List2.Clear
        Else
            If tamaño <> 0 Then
                For i = 0 To (tamaño - 1)
                    List2.AddItem (constituyentes(i))
                Next
            End If
        End If
        
        
        'buscar compuesto
        rst1.Open "SELECT compuesto, fecha, controlado FROM mezclas where batch = '" & List1.Text & "'", cnn, adOpenStatic, adLockReadOnly
        Text1.Text = rst1.Fields("compuesto")
        Text2.Text = rst1.Fields("fecha")
        
        rst1.Close
        'buscar compuesto
    
    rst.Close
    cnn.Close
    Command4.Enabled = True
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    mOdO = 1
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rst1 As ADODB.Recordset
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst1 = New ADODB.Recordset
    
    sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"
    
        With cnn
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
            .Open
        End With
        
        rst.Open "SELECT batch FROM mezclas where aprobado = TRUE order by batch", cnn, adOpenStatic, adLockReadOnly
    
    frmAEDBLEND.List1.Clear
    frmAEDBLEND.List2.Clear
    frmAEDBLEND.Command4.Enabled = False
    frmAEDBLEND.Command5.Enabled = False
    frmAEDBLEND.Text1.Text = ""
    frmAEDBLEND.Text2.Text = ""
    frmAEDBLEND.Check1.Value = 0
    Do Until rst.EOF = True
        frmAEDBLEND.List1.AddItem (rst.Fields("batch"))
        rst.MoveNext
    Loop
    rst.Close
    cnn.Close
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    mOdO = 1
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rst1 As ADODB.Recordset
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst1 = New ADODB.Recordset
    
    sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"
    
        With cnn
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
            .Open
        End With
        
        rst.Open "SELECT batch FROM mezclas where aprobado = FALSE order by batch", cnn, adOpenStatic, adLockReadOnly
    
    frmAEDBLEND.List1.Clear
    frmAEDBLEND.List2.Clear
    frmAEDBLEND.Command4.Enabled = False
    frmAEDBLEND.Command5.Enabled = False
    frmAEDBLEND.Text1.Text = ""
    frmAEDBLEND.Text2.Text = ""
    frmAEDBLEND.Check1.Value = 0
    Do Until rst.EOF = True
        frmAEDBLEND.List1.AddItem (rst.Fields("batch"))
        rst.MoveNext
    Loop
    rst.Close
    cnn.Close
End If
End Sub
