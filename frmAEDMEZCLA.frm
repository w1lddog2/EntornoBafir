VERSION 5.00
Begin VB.Form frmAEDMEZCLA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mezclas AED"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "?"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Controlado?"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox Text3 
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   240
      Width           =   1815
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
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3120
      Width           =   2055
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
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Blend?"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Aprobado"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha aprob."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Batch"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmAEDMEZCLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
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
    
    rst.Open "SELECT * FROM mezclas where BATCH = '" & Combo1.Text & "'", cnn, adOpenStatic, adLockReadOnly
    If rst.RecordCount = 0 Then
    
    Else
        frmAEDMEZCLA.Text1.Text = rst.Fields("fecha")
        'frmAEDMEZCLA.Text2.Text = rst.Fields("kilos")
        frmAEDMEZCLA.Text3.Text = rst.Fields("compuesto")
        frmAEDMEZCLA.Text4.Text = rst.Fields("fecha_aprobado") & ""
        frmAEDMEZCLA.Check3.Value = (-1) * (CInt(rst.Fields("controlado")))
        frmAEDMEZCLA.Check1.Value = (-1) * (CInt(rst.Fields("aprobado")))
        frmAEDMEZCLA.Check2.Value = (-1) * (CInt(rst.Fields("blend")))
        frmAEDMEZCLA.Command3.Enabled = True
    End If
cnn.Close
Text1.Enabled = False
'Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
End Sub
Private Sub Command1_Click()
baTCh = InputBox("Ingrese número de batch", "Batch AED")
Combo1.Enabled = False
Text1.Enabled = True

Text3.Enabled = True
Text4.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Combo1.Text = ""
Text1.Text = ""

Text3.Text = ""
Text4.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Combo1.Text = baTCh
End Sub

Private Sub Command2_Click()
Text1.Enabled = True

Text3.Enabled = True
Text4.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True

End Sub

Private Sub Command3_Click()
If Check1.Value = True Then
    If Text4.Text = "" Then
        sdasdasd = MsgBox("Debe designar la fecha de aprobación", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
End If
'rutina para guardar
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
    rst.Open "SELECT * FROM mezclas where batch = '" & Combo1.Text & "'", cnn, adOpenStatic, adLockOptimistic
    If rst.RecordCount = 0 Then
        rst.Close
        rst.Open "SELECT * FROM mezclas", cnn, adOpenStatic, adLockOptimistic
        rst.AddNew
        rst.Fields("batch") = Combo1.Text
        rst.Fields("fecha") = Text1.Text
        
        rst.Fields("compuesto") = Text3.Text
        If Text4.Text = "" Then
            fechaaprob = Empty
        Else
            fechaaprob = Text4.Text
        End If
        rst.Fields("fecha_aprobado") = fechaaprob
        rst.Fields("aprobado") = Check1.Value
        rst.Fields("blend") = Check2.Value
        rst.Fields("controlado") = Check3.Value
        rst.Update
        Combo1.AddItem (Combo1.Text)
    Else
        rst.Fields("fecha") = Text1.Text
                  
        rst.Fields("compuesto") = Text3.Text
        If Check1.Value = 1 Then
            If Text4.Text = "" Then
                dfsdfsdf = MsgBox("Debe ingresar un fecha de ensayo", vbCritical + vbOKOnly, "Error")
                Exit Sub
            Else
                faprob = Text4.Text
            End If
        Else
            faprob = Empty
        End If
        rst.Fields("fecha_aprobado") = faprob
        rst.Fields("aprobado") = Check1.Value
        rst.Fields("blend") = Check2.Value
        rst.Fields("controlado") = Check3.Value
        rst.Update
        
    End If
'******

Text1.Enabled = False

Text3.Enabled = False
Text4.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Combo1.Enabled = True


cnn.Close
End Sub

Private Sub Command4_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Command6_Click()
If Combo1.Text = "" Then
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
frmAEDDOC.AEDstr = "Select Observaciones From mezclas where batch = '" & Combo1.Text & "'"
Me.Enabled = False
frmAEDDOC.Show
End Sub

Private Sub Form_Load()
Form1.AEDPassword = ""
End Sub
