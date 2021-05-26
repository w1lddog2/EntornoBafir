VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajando"
   ClientHeight    =   1425
   ClientLeft      =   6330
   ClientTop       =   3600
   ClientWidth     =   3075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1425
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Trabajando. Por favor espere."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ProgressBar1.Value = 100
End Sub
