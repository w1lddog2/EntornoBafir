VERSION 5.00
Begin VB.Form frmLogger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logger"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Purgar"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM sys1 where maquina = '" & List1.Text & "'")

'Manipular strings

todos = Explode("@", rs.Fields("datos"))

cantidad = UBound(todos) ' establece la cantidad de registros
Open "c:\windowt\desktop\" & List1.Text & "_export.txt" For Output As #1
For i = 1 To cantidad
    
    Print #1, todos(i)
    
Next
    Close #1
End Sub

Private Sub Command2_Click()
frmLogger.Hide
End Sub

Private Sub Command3_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT maquina FROM sys1")

Do Until rs.EOF = True
    List1.AddItem (rs.Fields("maquina"))
    rs.MoveNext
Loop
db.Close
End Sub

Private Sub Command4_Click()
If List1.Text = "" Then
    sdfsdf = MsgBox("Selecciona una maquina", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM sys1 where maquina = '" & List1.Text & "'")

rs.Edit
rs.Fields("datos") = " "
rs.Update
db.Close
End Sub
