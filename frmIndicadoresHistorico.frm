VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmIndicadoresHistorico 
   Caption         =   "Historico de Indicadores"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11595
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   8955
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8880
      TabIndex        =   22
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Graficar"
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Graficar"
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Graficar"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   8160
      Width           =   1935
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   8160
      Width           =   1815
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   7680
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   5400
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2040
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   4
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   4
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   1215
      Left            =   240
      TabIndex        =   13
      Top             =   6360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   4
   End
   Begin VB.Label Label9 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Desde"
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Indicadores de Desarrollo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Desde"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Indicadores de Partidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Indicadores de Reometría"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmIndicadoresHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmIndicadores.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click()
frmIndicadores.Flag = True
frmIndicadores.Actualizar_historico
End Sub

Private Sub Command4_Click()
Me.Enabled = False
frmIndicadoresGraph.Show
End Sub
