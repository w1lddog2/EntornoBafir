VERSION 5.00
Begin VB.Form frmConsultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Desconoce"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Desconoce"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Desconoce"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Text            =   "Combo3"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Text            =   "Combo3"
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Desconoce"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmConsultas.frx":0000
      Top             =   4560
      Width           =   7695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Indicar unidad"
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Presión"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Uso"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Medio/Fluido"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "ºC"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Temperatura de uso"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Receptor de consulta"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de consulta"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo de consulta"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Controla_consultas()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select codigo, respuesta from consultas where estado = False")

If rs.RecordCount <> 0 Then
    textt = "Las siguientes consultas de cliente aún no han sido contestadas:"
    Do Until rs.EOF = True
    textt = textt & " " & rs.Fields("codigo") & ";"
    rs.MoveNext
    Loop
    kgl = MsgBox(textt, vbCritical + vbOKOnly, "Consultas de cliente")
End If



db.Close
End Sub











Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text4.Enabled = False
    Text4.Text = ""
Else
    Text4.Enabled = True
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Text5.Enabled = False
    Text5.Text = ""
Else
    Text5.Enabled = True
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Combo3.Enabled = False
    Combo3.Text = ""
Else
    Combo3.Enabled = True
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    Combo4.Enabled = False
    Combo4.Text = ""
Else
    Combo4.Enabled = True
End If
End Sub

Private Sub Command1_Click()

Dim medio As String
Dim uso As String
Dim presion As String

If Combo1.Text = "" Then
    sdfasdfs = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "error")
    Exit Sub
End If

If Combo2.Text = "" Then
    sdfasdfs = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "error")
    Exit Sub
End If




If Check1.Value = 1 Then
    'temperatura = Nothing
    temperatura = ""
Else
    If Text4.Text = "" Then
        sdfasdfs = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "error")
        Exit Sub
    End If
    temperatura = Text4.Text
    cont = cont + 1
End If

If Check2.Value = 1 Then
    presion = ""
Else
    If Text5.Text = "" Then
        sdfasdfs = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "error")
        Exit Sub
    End If
    presion = Text5.Text
    cont = cont + 1
End If

If Check3.Value = 1 Then
    medio = ""
Else
    If Combo3.Text = "" Then
        sdfasdfs = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "error")
        Exit Sub
    End If
    medio = Combo3.Text
    cont = cont + 1
End If

If Check4.Value = 1 Then
    uso = ""
Else
    If Combo4.Text = "" Then
        sdfasdfs = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "error")
        Exit Sub
    End If
    uso = Combo4.Text
    cont = cont + 1
End If

If Not cont >= 1 Then
asdasd = MsgBox("Debe Ingresar al menos un dato", vbCritical + vbOKOnly, "Error")
Exit Sub
End If


Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * from consultas")

rs.AddNew
rs.Fields("codigo") = CInt(Text1.Text)
rs.Fields("fecha_consulta") = Text2.Text
rs.Fields("cliente") = Combo1.Text
rs.Fields("responsable_cons") = Combo2.Text
If Text3.Text = "" Then
    rs.Fields("consulta") = ""
Else
    rs.Fields("consulta") = Text3.Text
End If
If medio = "" Then
    rs.Fields("medio") = ""
Else
    rs.Fields("medio") = medio
End If
If temperatura = "" Then
    rs.Fields("temperatura") = ""
Else
    rs.Fields("temperatura") = temperatura
End If
If presion = "" Then
    rs.Fields("presion") = ""
Else
    rs.Fields("presion") = presion
End If
If uso = "" Then
    rs.Fields("uso") = ""
Else
    rs.Fields("uso") = uso
End If
'rs.Fields("respuesta") = ""
'rs.Fields("compuesto_elastomero") = ""
'rs.Fields("responsable_resp") = ""
'rs.Fields("fecha_resp") = Nothing
rs.Fields("estado") = False
rs.Update

sad = MsgBox("Se han ingresado satisfatoriamente los datos", vbInformation + vbOKOnly, "Ingreso de consulta")
Form1.Enabled = True
Form1.Visible = True
db.Close
frmConsultas.Hide
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
frmConsultas.Hide
End Sub

