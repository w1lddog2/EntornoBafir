VERSION 5.00
Begin VB.Form frmAEDLOTES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lotes AED"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "?"
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Controlado?"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox Text3 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   1215
      Left            =   1320
      TabIndex        =   8
      Top             =   3840
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Ensayo/Desarrollo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Producción"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Informe"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Informe"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Batch"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Lote"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmAEDLOTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
numeroDElote = InputBox("Ingrese el número de lote AED", "Nuevo Lote")
Text1.Text = Format(Now, "MM/DD/YY")
Text2.Text = numeroDElote
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check1.Value = 0
Check2.Value = 0
Text4.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Text3.Clear
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    rst.Open "SELECT batch FROM mezclas where aprobado =  TRUE order by batch", cnn, adOpenStatic, adLockReadOnly
    Do Until rst.EOF = True
        Text3.AddItem (rst.Fields("batch"))
        rst.MoveNext
    Loop





cnn.Close

End Sub

Private Sub Command2_Click()
loteAbuscar = InputBox("Ingrese el lote a buscar", "Buscar lote AED")

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = False

'''''''''''''''''busqueda
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    rst.Open "SELECT * FROM lotes where lote = '" & loteAbuscar & "'", cnn, adOpenStatic, adLockReadOnly
    If rst.RecordCount = 0 Then
        sdfsdsdf = MsgBox("El lote no se encuentra", vbCritical + vbOKOnly, "Error")
    
    Else
        Text2.Text = rst.Fields("lote")
        Text1.Text = rst.Fields("fecha")
        Text3.Text = rst.Fields("batch")
        Check1.Value = Abs(CInt(rst.Fields("informe")))
        Check2.Value = Abs(CInt(rst.Fields("controlado")))
        Text4.Text = rst.Fields("fecha_informe") & ""
        If rst.Fields("tipo") = True Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If
    End If
    cnn.Close
''''''''''''''''''''''''''''''''''''''''


End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
    sfsfsdfsdf = MsgBox("Debe buscar un lote primero", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

Text1.Enabled = True
Text3.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Text4.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
'''''''''''''carga el combo
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    rst.Open "SELECT batch FROM mezclas where aprobado =  TRUE", cnn, adOpenStatic, adLockReadOnly
    Do Until rst.EOF = True
        Text3.AddItem (rst.Fields("batch"))
        rst.MoveNext
    Loop
cnn.Close
'''''''''''''''''''''''''''''
End Sub

Private Sub Command4_Click()
If Text3.Text = "" Then
    asdsdfdf = MsgBox("Debe definir un batch", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Option1.Value = False And Option2.Value = False Then
    sdfsdfsdf = MsgBox("Debe definir al menos un tipo", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset


If Text1.Enabled = False Then
    ''''''para guardar uno nuevo
    
    
    sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"
    
        With cnn
            'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
            .Open
        End With
        'rst simulacion
        'rst1 real
        rst.Open "SELECT * FROM lotes", cnn, adOpenStatic, adLockOptimistic
        rst.AddNew
        rst.Fields("lote") = Text2.Text
        rst.Fields("fecha") = Text1.Text
        rst.Fields("batch") = Text3.Text
        rst.Fields("informe") = Check1.Value
        rst.Fields("controlado") = Check2.Value
        If Check1.Value = False Then
            fechainfo = Empty
        Else
            fechainfo = Text4.Text
        End If
        rst.Fields("fecha_informe") = fechainfo
        If Option1.Value = True Then
            rst.Fields("tipo") = True
        Else
            rst.Fields("tipo") = False
        End If
        
        On Error Resume Next
            rst.Update
            Erri = Err.Number
            If Erri = -2147217887 Then
                sdfsdfsfsdf = MsgBox("El lote ya existe", vbCritical + vbOKOnly, "Error")
                On Error GoTo 0
                cnn.Close
                frmAEDLOTES.Text1.Text = ""
                frmAEDLOTES.Text2.Text = ""
                frmAEDLOTES.Text3.Text = ""
                frmAEDLOTES.Text4.Text = ""
                frmAEDLOTES.Check1.Value = False
                frmAEDLOTES.Check2.Value = False
                frmAEDLOTES.Command3.Enabled = False
                frmAEDLOTES.Command4.Enabled = False
                Exit Sub
            End If
                
        
        cnn.Close
ElseIf Text1.Enabled = True Then
''''''para guardar una modificacion


    
    sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"
    
        With cnn
            'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
            .Open
        End With
        'rst simulacion
        'rst1 real
        rst.Open "SELECT * FROM lotes where lote = '" & Text2.Text & "'", cnn, adOpenStatic, adLockOptimistic
        rst.Fields("fecha") = Text1.Text
        rst.Fields("batch") = Text3.Text
        rst.Fields("informe") = Check1.Value
        rst.Fields("controlado") = Check2.Value
        If Check1.Value = 1 Then
            If Text4.Text = "" Then
                sdfsdfsdf = MsgBox("Debe ingresar una fecha de emisión de informe AED", vbCritical + vbOKOnly, "Error")
                Text4.SetFocus
                Exit Sub
            Else
                fechainfo = Text4.Text
            End If
        End If
        If Check1.Value = 0 Then
            fechainfo = Empty
        End If
        
        rst.Fields("fecha_informe") = fechainfo
        If Option1.Value = True Then
            rst.Fields("tipo") = True
        Else
            rst.Fields("tipo") = False
        End If
        rst.Update
        cnn.Close
End If

''''''''''''''''''''''

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Text4.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Option1.Value = True
End Sub

Private Sub Command5_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Command6_Click()
If Text2.Text = "" Then
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
frmAEDDOC.AEDstr = "Select Observaciones From lotes where lote = '" & Text2.Text & "'"
Me.Enabled = False
frmAEDDOC.Show
End Sub

Private Sub Form_Load()
Form1.AEDPassword = ""
End Sub

Private Sub Option1_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    rst.Open "SELECT batch FROM mezclas where aprobado =  TRUE order by batch", cnn, adOpenStatic, adLockReadOnly
    Text3.Clear
    Do Until rst.EOF = True
        Text3.AddItem (rst.Fields("batch"))
        rst.MoveNext
    Loop





cnn.Close
End Sub

Private Sub Option2_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    rst.Open "SELECT batch FROM mezclas order by batch", cnn, adOpenStatic, adLockReadOnly
    If Text3.Text <> "" Then
    
    Else
        Text3.Clear
    End If
    Do Until rst.EOF = True
        Text3.AddItem (rst.Fields("batch"))
        rst.MoveNext
    Loop
End Sub
