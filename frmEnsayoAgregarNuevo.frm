VERSION 5.00
Begin VB.Form frmEnsayoAgregarNuevo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Nuevo ensayo"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   4200
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   7680
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Medio"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Temp ºC"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Horas"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmEnsayoAgregarNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public existe As Boolean
Public complety
Private Sub Command1_Click()
    AgregarEnsayo
    If frmEnsayoAgregarNuevo.complety = False Then
        Exit Sub
    End If
    Form1.Enabled = True
    Me.Hide
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Me.Hide
End Sub

Private Sub Command3_Click()
    AgregarEnsayo
    If frmEnsayoAgregarNuevo.complety = False Then
        Exit Sub
    End If
    If frmEnsayoAgregarNuevo.existe = False Then
        frmTraccion.Text2.AddItem List1.Text & " " & Text2.Text & " hs " & Text3.Text & " ºC " & List2.Text
    End If
    frmTraccion.Enabled = True
    Me.Hide
End Sub

Private Sub Command4_Click()
frmTraccion.Enabled = True
Me.Hide
End Sub

Sub AgregarEnsayo()
    existe = False
    complety = True
If Text2.Text = "" Or Text3.Text = "" Or List1.Text = "" Or List2.Text = "" Then
    sdfsdfs = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    frmEnsayoAgregarNuevo.complety = False
    Exit Sub
End If
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

    rst.Open "SELECT * FROM ensayos where referencia ='" & List1.Text & " " & Text2.Text & " hs " & Text3.Text & " ºC " & List2.Text & "' and tipo = '" & List1.Text & "'", cnn, adOpenStatic, adLockOptimistic
    If rst.RecordCount <> 0 Then
        dfsdfsdfs = MsgBox("El ensayo ingresado ya existe", vbCritical + vbOKOnly, "Error")
        existe = True
        cnn.Close
        Exit Sub
    End If
    rst.Close
    
    rst.Open "SELECT * FROM ensayos", cnn, adOpenStatic, adLockOptimistic
    rst.AddNew
    rst.Fields("codigo") = Text1.Text
    rst.Fields("referencia") = List1.Text & " " & Text2.Text & " hs " & Text3.Text & " ºC " & List2.Text
    rst.Fields("tipo") = List1.Text
    rst.Update
    cnn.Close
End Sub

Private Sub Command5_Click()
frmSeleccionarEnsayo.Enabled = True
Me.Hide
End Sub

Private Sub Command6_Click()
    AgregarEnsayo
    If frmEnsayoAgregarNuevo.complety = False Then
        Exit Sub
    End If
    If frmEnsayoAgregarNuevo.existe = False Then
        frmSeleccionarEnsayo.List1.AddItem List1.Text & " " & Text2.Text & " hs " & Text3.Text & " ºC " & List2.Text
    End If
    frmSeleccionarEnsayo.Enabled = True
    Me.Hide
End Sub

Private Sub Command7_Click()
    AgregarEnsayo
    If frmEnsayoAgregarNuevo.complety = False Then
        Exit Sub
    End If
    If frmEnsayoAgregarNuevo.existe = False Then
        frmfluido.List1.AddItem List1.Text & " " & Text2.Text & " hs " & Text3.Text & " ºC " & List2.Text
    End If
    frmfluido.Enabled = True
    Me.Hide
End Sub

Private Sub Command8_Click()
frmfluido.Enabled = True
Me.Hide
End Sub

