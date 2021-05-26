VERSION 5.00
Begin VB.Form frmAgregarMedicion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Medición"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese el nombre de la medición, por ejemplo Tracción"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmAgregarMedicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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

   
    rst.Open "SELECT mediciones FROM mediciones where mediciones = '" & Text1.Text & "'", cnn, adOpenStatic, adLockReadOnly
    
    If rst.RecordCount <> 0 Then
        sdsdf = MsgBox("Ya existe la medición ingresada", vbCritical + vbOKOnly, "Error")
    Else
        rst.Close
        rst.Open "SELECT * FROM mediciones order by codigo desc", cnn, adOpenStatic, adLockOptimistic
        Codigo = rst.Fields("codigo") + 1
        rst.AddNew
        rst.Fields("codigo") = Codigo
        rst.Fields("Mediciones") = Text1.Text
        rst.Update
    End If
    cnn.Close
    asdasd = MsgBox("La medición fue agregada satisfactoriamente", vbInformation + vbOKOnly, "Ingreso de medición")
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Me.Hide
End Sub
