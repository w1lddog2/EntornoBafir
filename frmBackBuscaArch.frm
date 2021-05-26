VERSION 5.00
Begin VB.Form frmBackBuscaArch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Archivo a la lista de Backing Up"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmBackBuscaArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
Dir1.Path = File1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Drive1.Drive

End Sub
