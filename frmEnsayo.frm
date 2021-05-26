VERSION 5.00
Begin VB.Form frmEnsayo 
   Caption         =   "Ingresar Ensayo Nuevo"
   ClientHeight    =   2085
   ClientLeft      =   4950
   ClientTop       =   5025
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   2085
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Ingrese solo los números, sin unidad y sin espacios"
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Fluido"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "ºC"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Horas"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmEnsayo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'FrmFluidoNuevo.Enabled = True
'FrmFluidoNuevo.List1.AddItem (frmEnsayo.Text1.Text & "hs" & " " & frmEnsayo.Text2.Text & "ºC" & " " & frmEnsayo.Text3.Text)
'FrmFluidoNuevo.List1.Text = frmEnsayo.Text1.Text & "hs" & " " & frmEnsayo.Text2.Text & "ºC" & " " & frmEnsayo.Text3.Text
'FrmFluidoNuevo.Visible = True
'frmEnsayo.Hide
frmfluido.List1.AddItem (Text1.Text & " hs " & Text2.Text & " ºC " & Text3.Text)
frmfluido.List1.Text = (Text1.Text & " hs " & Text2.Text & " ºC " & Text3.Text)
frmfluido.Enabled = True
frmfluido.Show
frmEnsayo.Hide
End Sub

Private Sub Command2_Click()
'FrmFluidoNuevo.Enabled = True
'FrmFluidoNuevo.Visible = True
'frmEnsayo.Hide
frmfluido.Enabled = True
frmfluido.Show
frmEnsayo.Hide
End Sub

