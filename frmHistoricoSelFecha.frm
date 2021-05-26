VERSION 5.00
Begin VB.Form frmHistoricoSelFecha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione rango de fecha"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmHistoricoSelFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public deSde
Public haSta

Private Sub Combo1_Click()


Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Historico_precios", dbOpenTable)

frmHistoricoSelFecha.Combo2.Clear

iniciode = Combo1.ListIndex
Columnas = rs.Fields.Count
For i = iniciode + 2 To Columnas - 1
    frmHistoricoSelFecha.Combo2.AddItem (rs.Fields(i).Name)
Next
End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Or Combo2.Text = "" Then
    asdasd = MsgBox("Debe seleccionar las fechas", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If


deSde = Combo1.ListIndex + 2
haSta = deSde + Combo2.ListIndex + 1
Me.Hide
End Sub

