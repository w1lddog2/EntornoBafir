VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGridDur 
   Caption         =   "Form4"
   ClientHeight    =   3165
   ClientLeft      =   4605
   ClientTop       =   4305
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   3165
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar a Excel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   4
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmGridDur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmGridDur.MSFlexGrid1.Clear
FrmLotedureza.Enabled = True
FrmLotedureza.Visible = True
frmGridDur.Caption = ""
frmGridDur.Hide
End Sub

