VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmespesores 
   Caption         =   "Espesores de matriz de tacos para cálculo de contracción"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   4965
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   $"frmespesores.frx":0000
      Height          =   1815
      Left            =   6840
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmespesores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\entornobafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    
    rst.Open "SELECT * FROM espesores where compuesto = '" & Combo1.Text & "'", cnn, adOpenStatic, adLockReadOnly
    If rst.RecordCount = 0 Then
        sdfssfg = MsgBox("No se han encontrado registros para el compuesto requerido", vbCritical + vbOKOnly, "Datos no disponibles")
    Else
    
    MSFlexGrid1.Rows = rst.RecordCount + 1
    
    MSFlexGrid1.TextMatrix(0, 0) = "Fecha"
    MSFlexGrid1.TextMatrix(0, 1) = "Compuesto"
    MSFlexGrid1.TextMatrix(0, 2) = "Partida"
    MSFlexGrid1.TextMatrix(0, 3) = "espesor"
    
    Fila = 1
    
    Do Until rst.EOF = True
        MSFlexGrid1.TextMatrix(Fila, 0) = rst.Fields("Fecha")
        MSFlexGrid1.TextMatrix(Fila, 1) = rst.Fields("Compuesto")
        MSFlexGrid1.TextMatrix(Fila, 2) = rst.Fields("Partida")
        MSFlexGrid1.TextMatrix(Fila, 3) = rst.Fields("Espesor")
        rst.MoveNext
        Fila = Fila + 1
    Loop
    
    AutoGrid frmespesores.MSFlexGrid1
    
    
    
    End If
End Sub

Private Sub Command1_Click()
Form1.Enabled = True
Unload Me
End Sub
