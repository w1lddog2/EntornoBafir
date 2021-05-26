VERSION 5.00
Begin VB.Form frmBackSelDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione Ruta"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Carpeta"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmBackSelDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmBack.Label7.Caption = Text1.Text
frmBackSelDir.Hide
frmBack.Enabled = True
frmBack.SetFocus
End Sub

Private Sub Command2_Click()
Dim fso As FileSystemObject

Set fso = New FileSystemObject

Carpeta = InputBox("Ingrese el nombre de la carpeta", "Crear Carpeta")
fso.CreateFolder (Text1.Text & "\" & Carpeta)
Dir1.Path = Text1.Text & "\" & Carpeta
End Sub

Private Sub Command3_Click()
frmBackSelDir.Hide
frmBack.Enabled = True
frmBack.SetFocus
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
