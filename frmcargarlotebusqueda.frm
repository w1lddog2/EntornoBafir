VERSION 5.00
Begin VB.Form frmcargarlotebusqueda 
   Caption         =   "Buscar Lote"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   6450
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   27
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar Lote"
      Height          =   495
      Left            =   1800
      TabIndex        =   23
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maquina"
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "Prensa"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Inyectora"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Lote"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Pieza"
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad"
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "O.T."
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Matriz"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Bocas"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Moldeadas"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Compuesto"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Partida"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label13 
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Cota de control"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "frmcargarlotebusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmCargarLote.buscar = True
a = frmCargarLote.Calcula_Muestreo(CLng(Label13.Caption))
b = frmCargarLote.calcula_controles_intermedios(a, CLng(Label13.Caption))

 j = frmCargarLote.Imprime_Muestreo(CLng(Label13.Caption), Text1.Text, Label4.Caption, Option1.Value, CLng(Text2.Text), CLng(a), CLng(b), Text7.Text, Text3.Text, Text6.Text)




End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click()
numero_lote = InputBox("Ingrese el número de lote a buscar", "Buscar Lote")

Dim db As Database
Dim db1 As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, True)
Set rs = db.OpenRecordset("select OT,NRO_LOTE,CDGO_PIEZA,CANT_PIEZA,COMPUESTO,FECHA,PARTIDA,OBSERVA1,NRO_MATRIZ,FRECUENCIA_CONTROL,NIVEL_DE_INSPECCION,NIVEL_DE_ACEPTACION,MAQUINA,COTA_CONTROL1 from LOTES where NRO_LOTE = " & numero_lote)

If rs.RecordCount = 0 Then
    sdfsdfsdf = MsgBox("El lote solicitado no existe", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If


Label2.Caption = rs.Fields("fecha")
Label4.Caption = numero_lote
Text1.Text = rs.Fields("cdgo_pieza")
Text2.Text = rs.Fields("cant_pieza")
Text3.Text = rs.Fields("ot") & ""
Text4.Text = rs.Fields("nro_matriz")
Set db1 = OpenDatabase("\\Servidor2\e\entornobafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs1 = db1.OpenRecordset("select bocas from tabla_piezas where NRO_pieza = '" & Text1.Text & "' and matriz = '" & Text4.Text & "'")
If rs1.RecordCount = 0 Then
    sdfsdfsdf = MsgBox("Es posible que el lote que está buscando sea anterior a la implementación de visual. Es posible que algunos datos no estén disponibles. Ante cualquier duda, consulte al administrador del sistema", vbCritical + vbOKOnly, "Error")
    Text5.Text = ""
    Label13.Caption = ""
Else
    Text5.Text = rs1.Fields("bocas")
    Label13.Caption = CInt((CDbl(Text2.Text / Text5.Text)) + 0.5)
End If
Text6.Text = rs.Fields("compuesto")
Label14.Caption = rs.Fields("partida")

If rs.Fields("maquina") = 1 Then
    Option1.Value = True
Else
    Option2.Value = True
End If
Text7.Text = rs.Fields("cota_control1") & ""

db.Close
End Sub

