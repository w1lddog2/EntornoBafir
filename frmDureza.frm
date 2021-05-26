VERSION 5.00
Begin VB.Form frmDureza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de toma de durezas o ensayos varios"
   ClientHeight    =   3555
   ClientLeft      =   4470
   ClientTop       =   3135
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Text            =   "text4"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "text3"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Responsable"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Pieza"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Lote"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmDureza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset

If Text1.Text = "" Then
ds = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
ds = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
ds = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
Text3.SetFocus
Exit Sub
End If
If Text4.Text = "" Then
ds = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
Text4.SetFocus
Exit Sub
End If
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
'Set rs = db.OpenRecordset("Select LOTE, PIEZA, COMPUESTO, RESPONSABLE_SOLICITUD, FECHA_SOLICITUD, DUREZA_LABORATORIO, RESPONSABLE_LABORATORIO, FECHA_LABORATORIO From Durezas_Produccion Where LOTE = '" & Form3.Text1.Text & "' ;")
Set rs = db.OpenRecordset("Durezas_produccion", dbOpenTable)
rs.Index = "primarykey"
rs.Seek "=", Text1.Text
If rs.NoMatch = False Then
    ars = MsgBox("El lote ya ha sido cargado con anterioridad. Desea resolicitar la medición?.", vbCritical + vbYesNo, "Lote existente")
    If ars = vbYes Then
    rs.Edit
    rs.Fields("responsable_solicitud") = rs.Fields("responsable_solicitud") & " R"
    rs.Fields("fecha_solicitud") = Date & " R"
    rs.Fields("dureza_laboratorio") = 0
    rs.Fields("responsable_laboratorio") = 0
    rs.Fields("fecha_laboratorio") = 0
    rs.Fields("tipo_ensayo") = "Especificar"
    rs.Update
    Else
    Exit Sub
    End If
Else
rs.AddNew
rs.Fields("lote") = Text1.Text
rs.Fields("pieza") = Text2.Text
rs.Fields("Compuesto") = Text3.Text
rs.Fields("responsable_solicitud") = Text4.Text
rs.Fields("fecha_solicitud") = Date
rs.Fields("dureza_laboratorio") = 0
rs.Fields("responsable_laboratorio") = 0
rs.Fields("fecha_laboratorio") = 0
rs.Fields("tipo_ensayo") = "Especificar"
rs.Update
ert = MsgBox("El lote ha sido solicitado", vbInformation + vbOKOnly, "Carga de datos")
End If
db.Close


ReDim destinatarios(1 To 4)

indicedestinatarios = 4
asunto = "Entorno Bafir: Solicitud de medición para pieza " & Text2.Text
mail = ": Se ha solicitado la medición de la pieza " & Text2.Text & " por " & Text4.Text

destinatarios(1) = "pablopirri@bafir.com.ar"
destinatarios(2) = "laboratorio@bafir.com.ar"
destinatarios(3) = "entornobafir@gmail.com"
destinatarios(4) = "laboratoriobafir@gmail.com"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Label6.Caption = ""

'frmSendinfo.Show
'frmSendinfo.Hide
frmSendinfo.Show
frmSendinfo.Visible = False
Call Moduloenvio
frmSendinfo.Hide
Text1.SetFocus

End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
frmDureza.Hide
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text1.TabIndex = 1
Text2.Text = ""
Text2.TabIndex = 2
Text3.Text = ""
Text3.TabIndex = 3
Text4.Text = ""
Text4.TabIndex = 4
Label6.Caption = ""
Command1.TabIndex = 5
Command2.TabIndex = 6
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
    Label6.Caption = ""
Else
    Label6.Caption = Date
End If
End Sub

