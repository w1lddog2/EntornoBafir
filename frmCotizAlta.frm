VERSION 5.00
Begin VB.Form frmCotizAlta 
   Caption         =   "Alta de pieza para ensayos de cliente"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Norma"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Pieza Bafir"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmCotizAlta"
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

   
    rst.Open "SELECT * FROM piezas_en_produccion where pieza = '" & Text1.Text & "'", cnn, adOpenStatic, adLockOptimistic
    If rst.RecordCount <> 0 Then
        dadasd = MsgBox("La pieza mencionada ya existe", vbCritical + vbOKOnly, "Error")
        cnn.Close
        Exit Sub
    End If
    rst.AddNew
    rst.Fields("Pieza") = Text1.Text
    rst.Fields("Norma") = Text3.Text
    rst.Fields("cliente") = Text2.Text
    rst.Fields("compuesto") = Text4.Text
    rst.Fields("fecha_alta") = Now()
    rst.Update
    cnn.Close
    
frmBusca.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()
frmBusca.Enabled = True
Unload Me
End Sub

