VERSION 5.00
Begin VB.Form frmCuerdaInforme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de cuerda realizado"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Informe Realizado"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Lote"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Cuerda"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmCuerdaInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim db As Database
Dim rs As Recordset
Combo2.Clear
Combo3.Clear

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select  cuerda from cuerdas where proveedor ='" & Combo1.Text & "' group by cuerda")
If rs.RecordCount = 0 Then
sfsfsfdf = MsgBox("No se encuentra el proveedor", vbCritical + vbOKOnly, "Error")
    Combo1.SetFocus
    Exit Sub
End If
rs.MoveFirst
Do Until rs.EOF = True
Combo2.AddItem (rs.Fields("cuerda"))
rs.MoveNext
Loop
Combo2.SetFocus
db.Close
End Sub

Private Sub Combo2_Click()
Dim db As Database
Dim rs As Recordset
Combo3.Clear
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select  lote, cuerda, proveedor from cuerdas where proveedor ='" & Combo1.Text & "' and cuerda = '" & Combo2.Text & "'")
If rs.RecordCount = 0 Then
sfsfsfdf = MsgBox("No se encuentra.", vbCritical + vbOKOnly, "Error")
    Combo2.SetFocus
    Exit Sub
End If
rs.MoveFirst
Do Until rs.EOF = True
Combo3.AddItem (rs.Fields("lote"))
rs.MoveNext
Loop
db.Close
End Sub

Private Sub Combo3_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select  informe_realizado, lote, cuerda, proveedor, remito from cuerdas where proveedor ='" & Combo1.Text & "' and cuerda = '" & Combo2.Text & "' and lote = '" & Combo3.Text & "'")

If rs.RecordCount = 0 Then
    Label4.Caption = "No encontrado"
    Combo3.SetFocus
    Exit Sub
End If
Label4.Caption = "Remito Nº " & rs.Fields("remito")
Check1.Value = Abs(CInt(rs.Fields("informe_realizado")))
db.Close
End Sub

Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select  informe_realizado, lote, cuerda, proveedor, remito from cuerdas where proveedor ='" & Combo1.Text & "' and cuerda = '" & Combo2.Text & "' and lote = '" & Combo3.Text & "'")
If rs.RecordCount = 0 Then
    Label4.Caption = "Error, No se ha grabado"
    Combo3.SetFocus
    Exit Sub
End If
rs.Edit
rs.Fields("informe_realizado") = Check1.Value
rs.Update
db.Close
Combo2.Clear
Combo3.Clear
Check1.Value = 0
Combo1.SetFocus
Label4.Caption = ""
Combo1.Text = ""
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
frmCuerdaInforme.Hide
End Sub
