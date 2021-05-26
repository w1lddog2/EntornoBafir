VERSION 5.00
Begin VB.Form frmconsultaver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   26
      Text            =   "Text11"
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   24
      Text            =   "Text10"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   22
      Text            =   "Text9"
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   495
      Left            =   7560
      TabIndex        =   21
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   6600
      Width           =   5415
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmconsultaver.frx":0000
      Top             =   4440
      Width           =   8775
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   3000
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   840
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmconsultaver.frx":0008
      Top             =   1680
      Width           =   8775
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Presion"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Uso"
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Resp."
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Responsable"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Compuesto"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Respuesta"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Temperatura"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Medio"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo de consulta"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de consulta"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Receptor de consulta"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmconsultaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
frmconsultabusca.Enabled = True
Me.Hide
End Sub
