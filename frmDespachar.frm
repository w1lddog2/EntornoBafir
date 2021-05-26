VERSION 5.00
Begin VB.Form frmDespachar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despachar mercadería"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1320
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   240
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Destino"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Partida"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Kilos"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Mercaderia"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmDespachar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "" Or Combo2.Text = "" Or Text1.Text = "" Or Text2.Text = "" Then
    sdfsdfgsfdg = MsgBox("Debe ingresar todos los datos antes de proseguir", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\centralpesado.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    
 
    rst.Open "SELECT COD_PROD FROM producto where descrip = '" & Combo1.Text & "'", cnn, adOpenStatic, adLockReadOnly
    codigomp = rst.Fields("cod_prod")
    rst.Close
    
    rst.Open "SELECT * FROM pesado", cnn, adOpenStatic, adLockOptimistic
    
    rst.AddNew
    rst.Fields("compuesto") = Combo2.Text
    rst.Fields("partida") = "-"
    rst.Fields("batch") = "ENVIO"
    rst.Fields("materia_prima") = codigomp
    rst.Fields("pesado") = Replace(Text1.Text, ".", ",")
    rst.Fields("loteMP") = Text2.Text
    rst.Fields("fecha") = Date
    rst.Fields("hora") = Time
    rst.Update
    
    rst.Close
    cnn.Close
    sdfgdsfgsghs = MsgBox("Se ha ingresado el despacho correctamente", vbInformation + vbOKOnly)
    Combo1.Text = ""
    Combo2.Text = ""
    Text1.Text = ""
    Text2.Text = ""

End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Unload Me
End Sub
