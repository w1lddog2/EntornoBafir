VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione el nombre del archivo a exportar"
   ClientHeight    =   3870
   ClientLeft      =   5100
   ClientTop       =   3780
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   120
      Pattern         =   "*.xls"
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Form3.Text1.Text = "" Then
    u = MsgBox("Debe seleccionar un archivo para grabar.", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
ruta = Form3.Dir1.path & "\" & Form3.Text1.Text
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.MousePointer = 11
Form1.Enabled = False
Form1.Visible = False
Form3.Enabled = False
Form3.Visible = False
Dim appp As New Excel.Application
Dim ws As New Excel.Worksheet
Dim wb As New Excel.Workbook
Dim r As Excel.Range
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Historico_precios", dbOpenTable)

Set wb = appp.Workbooks.Add
Set ws = wb.Worksheets.Add


frmHistoricoSelFecha.Combo1.Clear
frmHistoricoSelFecha.Combo2.Clear



Columnas = rs.Fields.Count
For i = 1 To Columnas - 1
    frmHistoricoSelFecha.Combo1.AddItem (rs.Fields(i).Name)
Next

frmHistoricoSelFecha.Show (1)

d = frmHistoricoSelFecha.deSde
h = frmHistoricoSelFecha.haSta



rs.MoveLast
ultimo = rs.RecordCount
ws.Activate
rs.MoveFirst


For contadorfila = 1 To ultimo - 1
Form2.ProgressBar1.Value = contadorfila / ultimo * 100
    'For contadorcol = 1 To Columnas
    
        
    For contadorcol = d To h
        ws.Cells(1, contadorcol - (d - 1)) = Format((rs.Fields(contadorcol - 1).Name), "MM/DD/YY")
        ws.Cells(contadorfila + 1, contadorcol - (d - 1)) = rs.Fields(contadorcol - 1)
    Next
rs.MoveNext
Next
rs.MoveFirst
    Dim fff As Integer
    fff = 2
    Do Until rs.EOF = True
        ws.Cells(fff, 1) = rs.Fields(0)
        rs.MoveNext
        fff = fff + 1
    Loop
        ws.Cells(1, 1) = ""
Form2.Hide
ws.SaveAs ruta
appp.Quit
Form1.Enabled = True
rt = MsgBox("Se ha grabado el archivo como " & ruta, vbInformation + vbOKOnly, "Archivo Guardado")
Form1.Enabled = True
Form1.Visible = True
Form3.Hide
db.Close
Form2.ProgressBar1.Visible = False
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
Form3.Hide
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
Form3.Text1.Text = ""
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
Form3.Text1.Text = ""
End Sub

Private Sub File1_Click()
Form3.Text1.Text = File1.FileName
End Sub
Private Sub Form_Load()
Dir1.path = Drive1.Drive
File1.path = Dir1.path
End Sub

