VERSION 5.00
Begin VB.Form frmCargarDureza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar durezas de piezas"
   ClientHeight    =   3195
   ClientLeft      =   2610
   ClientTop       =   2895
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5400
      TabIndex        =   21
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Tipo ensayo"
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Label14"
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Responsable"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Resultado"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Responsable"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Pieza"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Lote"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmCargarDureza"
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

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Durezas_produccion", dbOpenTable)
rs.Index = "primarykey"
rs.Seek "=", Text1.Text
If rs.NoMatch = True Then
    ars = MsgBox("El lote no ha sido cargado.", vbCritical + vbOKOnly, "Lote no existente")
    Label14.Visible = False
    Text1.SetFocus
    Exit Sub
Else

Label6.Caption = rs.Fields("pieza")
Label7.Caption = rs.Fields("Compuesto")
Label8.Caption = rs.Fields("responsable_solicitud")
Label9.Caption = rs.Fields("fecha_solicitud")
If rs.Fields("dureza_laboratorio") = 0 Then
    Label14.Caption = "Lote Abierto"
    Label14.BackColor = &HFF00&
    Label14.Visible = True
    Text2.Enabled = True
    Text3.Enabled = True
    Command3.Enabled = True
    Text2.Text = ""
     Text3.Text = ""
     Label13.Caption = ""
Else
    Label14.Caption = "Lote Cerrado"
    Label14.BackColor = &HFF&
    Label14.Visible = True
    Text2.Enabled = False
    Text3.Enabled = False
    Command3.Enabled = False
    
     Text2.Text = rs.Fields("dureza_laboratorio")
     Text3.Text = rs.Fields("responsable_laboratorio")
     Label13.Caption = rs.Fields("fecha_laboratorio")
End If
End If
db.Close
If Text2.Enabled = True Then
Text2.SetFocus
End If
End Sub

Private Sub Command2_Click()
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label13.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Form1.Enabled = True
Form1.Visible = True
frmCargarDureza.Hide
End Sub

Private Sub Command3_Click()
Dim db As Database
Dim rs As Recordset

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
If Combo1.Text = "" Then
ds = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
Combo1.SetFocus
Exit Sub
End If

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Durezas_produccion", dbOpenTable)
rs.Index = "primarykey"
rs.Seek "=", Text1.Text
rs.Edit
rs.Fields("dureza_laboratorio") = Text2.Text
rs.Fields("responsable_laboratorio") = Text3.Text
rs.Fields("fecha_laboratorio") = Label13.Caption
rs.Fields("tipo_ensayo") = Combo1.Text
Label14.Caption = "Lote Cerrado"
Label14.BackColor = &HFF&
Label14.Visible = True
Text2.Enabled = False
Text3.Enabled = False
Command3.Enabled = False
rs.Update
db.Close

ReDim destinatarios(1 To 5)

indicedestinatarios = 5
asunto = "Entorno Bafir: Carga de medición para pieza " & Label6.Caption
mail = ": Se ha respondido la medición de la pieza " & Label6.Caption & " por " & Text3.Text

destinatarios(1) = "pablopirri@bafir.com.ar"
destinatarios(2) = "laboratorio@bafir.com.ar"
destinatarios(3) = "produccion2@bafir.com.ar"
destinatarios(4) = "entornobafir@gmail.com"
destinatarios(5) = "laboratoriobafir@gmail.com"

frmSendinfo.Show
frmSendinfo.Visible = False

Call Moduloenvio

frmSendinfo.Hide
Command2.SetFocus
End Sub

Private Sub Form_Load()
Text1.TabIndex = 1
Command1.TabIndex = 2
Command2.TabIndex = 3
Text2.TabIndex = 4
Text3.TabIndex = 5
Command3.TabIndex = 6
End Sub

Private Sub Text1_Change()
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Text2.Text = ""
Text3.Text = ""

End Sub



Private Sub Text2_Change()
If Text2.Text = "" Then
    Label13 = ""
Else
    Label13 = Date
End If
End Sub
