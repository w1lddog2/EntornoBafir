VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmIndicadoresGraph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grafico de indicadores"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   6960
      Width           =   2175
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6015
      Left            =   480
      OleObjectBlob   =   "frmIndicadoresGraph.frx":0000
      TabIndex        =   0
      Top             =   480
      Width           =   9975
   End
End
Attribute VB_Name = "frmIndicadoresGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Grafica()
ReDim vaLores(1 To registros, 1 To 2)
Xval = 0

registros = frmIndicadoresHistorico.MSFlexGrid1.col

For i = 2 To registros
    vaLores(i, 1) = frmIndicadoresHistorico.MSFlexGrid1.TextMatrix(1, i)
    vaLores(i, 2) = frmIndicadoresHistorico.MSFlexGrid1.TextMatrix(2, i)
Next
MSChart1.ColumnCount = 2
MSChart1.ColumnLabelCount = 2
MSChart1.Column = 1
MSChart1.ColumnLabel = "Grafico de indicadores, Mezclas aprobadas"


MSChart1.RowCount = 2
MSChart1.Plot.Axis(VtChAxisIdX).AxisScale.Type = VtChScaleTypeLinear
MSChart1.Plot.Axis(VtChAxisIdY).AxisScale.Type = VtChScaleTypeLinear
MSChart1.Plot.UniformAxis = False

With MSChart1
    With .Plot.Axis(VtChAxisIdX).AxisTitle
        .Text = "Periodo"
    End With
    With .Plot.Axis(VtChAxisIdY).AxisTitle
        .Text = "% Aprobación"
    End With

End With



    


For i = 1 To registros
    MSChart1.DataGrid.SetData i, 1, CLng(CDate(vaLores(i, 1))), False
    MSChart1.DataGrid.SetData i, 2, vaLores(i, 2), False
Next

MSChart1.Refresh

End Sub

Private Sub Command1_Click()
frmIndicadoresHistorico.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Call Grafica
End Sub
