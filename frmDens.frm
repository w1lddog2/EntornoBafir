VERSION 5.00
Begin VB.Form frmDens 
   Caption         =   "Generar listado de densidades"
   ClientHeight    =   3795
   ClientLeft      =   4710
   ClientTop       =   2220
   ClientWidth     =   5505
   LinkTopic       =   "Form4"
   ScaleHeight     =   3795
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   360
      Pattern         =   "*.xls"
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frmDens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frmDens.Text1.Text = "" Then
    u = MsgBox("Debe seleccionar un archivo para grabar.", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
ruta = frmDens.Dir1.path & frmDens.Text1.Text
Form2.Show
Form2.MousePointer = 11
Form1.Enabled = False
Form1.Visible = False
frmDens.Enabled = False
frmDens.Visible = False
Dim appp As New Excel.Application
Dim ws As New Excel.Worksheet
Dim wb As New Excel.Workbook
Dim r As Excel.Range
Dim db As Database
Dim rs As Recordset


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select N_formula, Densidad From FormBase")


Set wb = appp.Workbooks.Add
Set ws = wb.Worksheets.Add

rs.MoveLast
ultimo = rs.RecordCount
ws.Activate
rs.MoveFirst
Form2.ProgressBar1.Visible = True
rs.MoveFirst
For contadorfila = 1 To ultimo - 1
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / ultimo
ws.Cells(1, 1) = "Compuesto"
ws.Cells(1, 2) = "Densidad"
ws.Cells(contadorfila + 1, 2) = rs.Fields("Densidad")
ws.Cells(contadorfila + 1, 1) = rs.Fields("N_formula")
rs.MoveNext
Next
ws.Cells(contadorfila + 3, 1) = "Fecha de emisión " & Date
Form2.Hide
ws.SaveAs ruta
appp.Quit
Form1.Enabled = True
rt = MsgBox("Se ha grabado el archivo como " & ruta, vbInformation + vbOKOnly, "Archivo Guardado")
Form1.Enabled = True
Form1.Visible = True
frmDens.Hide
Form1.Enabled = True
Form1.Visible = True
frmDens.Hide
db.Close
End Sub

Private Sub Command2_Click()

Form1.Enabled = True
Form1.Visible = True
frmDens.Hide
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
frmDens.Text1.Text = ""
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
frmDens.Text1.Text = ""
End Sub

Private Sub File1_Click()
frmDens.Text1.Text = File1.FileName
End Sub

