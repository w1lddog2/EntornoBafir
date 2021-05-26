VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBuscaComp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar  Ensayo de compresion"
   ClientHeight    =   4935
   ClientLeft      =   5355
   ClientTop       =   3600
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label Label4 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Partida"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Busca Por:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmBuscaComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" Then
busqueda = "Compuesto"
clave = Text1.Text
End If
If Text2.Text <> "" Then
busqueda = "Partida"
clave = Text2.Text
End If
If Text3.Text <> "" Then
busqueda = "Codigo_ensayo"
clave = Text3.Text
End If
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
If Text3.Text = "" Then
    Set rs = db.OpenRecordset("Select probeta, compresion, codigo_ensayo, compuesto, partida, tiempo_temperatura, compresion_porc from compresion where " & busqueda & " = '" & clave & "' ;")
Else
    Set rs = db.OpenRecordset("Select probeta, compresion, codigo_ensayo, compuesto, partida, tiempo_temperatura, compresion_porc from compresion where " & busqueda & " = " & clave & " ;")
End If
'InputBox ("Select codigo_ensayo, compuesto, partida, tiempo_temperatura, compresion_porc from compresion where " & busqueda & " = '" & clave & "' ;")

If rs.RecordCount = 0 Then
    fgsdfda = MsgBox("No se han hallado valores coincidentes a su busqueda", vbCritical + vbOKOnly, "Error")
    frmBuscaComp.Text1.SetFocus
    db.Close
    Exit Sub
End If
frmBuscaComp.Height = 4485
MSFlexGrid1.TextMatrix(0, 0) = "Compuesto"
MSFlexGrid1.TextMatrix(0, 1) = "Partida"
MSFlexGrid1.TextMatrix(0, 2) = "Codigo"
MSFlexGrid1.TextMatrix(0, 3) = "Ensayo"
MSFlexGrid1.TextMatrix(0, 4) = "% de deformación"
MSFlexGrid1.TextMatrix(0, 5) = "Probeta"
MSFlexGrid1.TextMatrix(0, 6) = "Compresion"
rs.MoveLast
ultimo = rs.RecordCount
rs.MoveFirst
MSFlexGrid1.Rows = ultimo + 1
For indice = 1 To ultimo
    MSFlexGrid1.TextMatrix(indice, 0) = rs.Fields("compuesto")
    MSFlexGrid1.TextMatrix(indice, 1) = rs.Fields("Partida")
    MSFlexGrid1.TextMatrix(indice, 2) = rs.Fields("Codigo_ensayo")
    MSFlexGrid1.TextMatrix(indice, 3) = rs.Fields("tiempo_temperatura")
    MSFlexGrid1.TextMatrix(indice, 4) = rs.Fields("compresion_porc")
    MSFlexGrid1.TextMatrix(indice, 5) = rs.Fields("probeta") & ""
    MSFlexGrid1.TextMatrix(indice, 6) = rs.Fields("compresion") & ""
    
    rs.MoveNext
Next

AutoGrid MSFlexGrid1

db.Close
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
frmBuscaComp.Hide
End Sub

Private Sub Command3_Click()
frmBuscaTraccion.Enabled = True
frmBuscaTraccion.Visible = True
Command3.Visible = False
Command2.Visible = True
If Text4.Visible = True Then
    Text4.Visible = False
    Text5.Visible = False
    Text6.Visible = False
    Command1.Enabled = True
End If
frmBuscaComp.Hide
End Sub

Private Sub Text1_Change()
Text2.Text = ""
Text3.Text = ""
frmBuscaComp.Height = 2745
End Sub

Private Sub Text2_Change()
Text1.Text = ""
Text3.Text = ""
frmBuscaComp.Height = 2745
End Sub

Private Sub Text3_Change()
Text1.Text = ""
Text2.Text = ""
frmBuscaComp.Height = 2745
End Sub
