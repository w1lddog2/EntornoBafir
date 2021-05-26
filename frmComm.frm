VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmComm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor de puerto"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Monitorear"
      Height          =   615
      Left            =   3360
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   3240
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   600
      Width           =   1455
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4080
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Actualizar datos (ms)"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Bits de parada"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Bits"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Paridad"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Velocidad"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "COM"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo1.AddItem ("1")
Combo1.AddItem ("2")
Combo1.AddItem ("3")
Combo1.AddItem ("4")

Combo2.AddItem ("110")
Combo2.AddItem ("300")
Combo2.AddItem ("600")
Combo2.AddItem ("1200")
Combo2.AddItem ("2400")
Combo2.AddItem ("4800")
Combo2.AddItem ("9600")
Combo2.AddItem ("14400")
Combo2.AddItem ("19200")
Combo2.AddItem ("28800")
Combo2.AddItem ("38400")
Combo2.AddItem ("56000")
Combo2.AddItem ("128000")
Combo2.AddItem ("256000")

Combo3.AddItem ("E")
Combo3.AddItem ("M")
Combo3.AddItem ("N")
Combo3.AddItem ("O")
Combo3.AddItem ("S")

Combo4.AddItem ("4")
Combo4.AddItem ("5")
Combo4.AddItem ("6")
Combo4.AddItem ("7")
Combo4.AddItem ("8")

Combo5.AddItem ("1")
Combo5.AddItem ("1.5")
Combo5.AddItem ("2")





End Sub
