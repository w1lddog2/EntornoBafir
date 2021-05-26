VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Entorno Bafir"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Image imgLogo 
      Height          =   1305
      Left            =   240
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select contraseña, email from loginusuarios where email = '" & Text1.Text & "'")

If rs.RecordCount = 0 Then
fsdffsf = MsgBox("Nombre de usuario o contraseña incorrecta. Por favor intente de nuevo y compruebe que no tenga activado CapsLock (Bloq Mayus).", vbCritical + vbOKOnly, "Error")
Text2.SetFocus
Exit Sub
Else
frmLogin.Hide
End If

End Sub

Private Sub Command2_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select conexion, Fechaaccion, maquina, Usuario, accion from loginonline")
    rs.AddNew
    rs.Fields("fechaaccion") = Now()
    rs.Fields("maquina") = Form1.ComputerName
    rs.Fields("Usuario") = GetSetting("EBafir", "Valores", "Usuario")
    rs.Fields("Accion") = "Cancelado"
    rs.Fields("conexion") = Form1.coneXion
    rs.Update
    
    Do
        rs.MoveLast
        If rs.RecordCount > 100 Then
            rs.MoveFirst
            rs.Delete
        Else
            Exit Do
        End If
    Loop



    db.Close


End





End Sub

