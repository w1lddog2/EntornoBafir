VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEnsayosVer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Ensayos"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7011
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEnsayosVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Me.Hide
End Sub
