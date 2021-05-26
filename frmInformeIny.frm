VERSION 5.00
Begin VB.Form frmInformeIny 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de productividad"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Generar Reporte"
      Height          =   315
      Left            =   5400
      TabIndex        =   27
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   2760
      TabIndex        =   26
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label26 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label25 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   8040
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   8040
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   8040
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   8040
      TabIndex        =   20
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   8040
      TabIndex        =   19
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000009&
      Caption         =   "Label14"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Maquina trabajando en orden no productivo"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Falta de trabajo"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Falta de operador"
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Falta de materia prima"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Dificultad de materia prima"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Lavado de matriz"
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Matriz Rota"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "En mantenimiento preventivo"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "En mantenimiento correctivo"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "En preparación"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Trabajando"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Inyectora"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Rango"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmInformeIny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmInformeIny.Hide
frmIny.Enabled = True
End Sub

Private Sub Command2_Click()
Dim appp As New Excel.Application
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim r As Excel.Range

Set wb = appp.Workbooks.Open("\\Servidor2\e\EntornoBafir\Planillas\informeiny.xls", , True)
Set ws = wb.Worksheets(1)

ws.Cells(5, 5) = Label14.Caption
ws.Cells(6, 5) = Label26.Caption
ws.Cells(7, 5) = Date
ws.Cells(17, 4) = Label15.Caption
ws.Cells(19, 4) = Label16.Caption
ws.Cells(21, 4) = Label17.Caption
ws.Cells(23, 4) = Label18.Caption
ws.Cells(25, 4) = Label19.Caption
ws.Cells(17, 8) = Label20.Caption
ws.Cells(19, 8) = Label21.Caption
ws.Cells(21, 8) = Label22.Caption
ws.Cells(23, 8) = Label23.Caption
ws.Cells(25, 8) = Label24.Caption
ws.Cells(27, 8) = Label25.Caption
ws.PrintOut
DoEvents
wb.Close (False)
End Sub
