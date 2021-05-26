VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAEDfinal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe final"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Inspec. Visual"
      Height          =   255
      Left            =   3960
      TabIndex        =   52
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   4800
      TabIndex        =   51
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   495
      Left            =   2280
      TabIndex        =   50
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox Text25 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   49
      Text            =   "25"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text24 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   48
      Text            =   "24"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text23 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   47
      Text            =   "23"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   46
      Text            =   "22"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text21 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   45
      Text            =   "21"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text20 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   44
      Text            =   "20"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   43
      Text            =   "19"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   42
      Text            =   "18"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   33
      Text            =   "17"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   32
      Text            =   "16"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text15 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   31
      Text            =   "15"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   30
      Text            =   "14"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   29
      Text            =   "13"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   28
      Text            =   "12"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   27
      Text            =   "11"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   26
      Text            =   "10"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   25
      Text            =   "9"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   24
      Text            =   "8"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   23
      Text            =   "7"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   22
      Text            =   "6"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   21
      Text            =   "5"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Text            =   "4"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Text            =   "3"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   18
      Text            =   "2"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label26 
      Height          =   375
      Left            =   2640
      TabIndex        =   53
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8400
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label25 
      Caption         =   "Variación de módulo"
      Height          =   255
      Left            =   3960
      TabIndex        =   41
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label24 
      Caption         =   "Variación de elongación"
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label23 
      Caption         =   "Variación de tracción"
      Height          =   255
      Left            =   3960
      TabIndex        =   39
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label22 
      Caption         =   "Variación de dureza"
      Height          =   255
      Left            =   3960
      TabIndex        =   38
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label21 
      Caption         =   "Variación de densidad"
      Height          =   255
      Left            =   3960
      TabIndex        =   37
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label20 
      Caption         =   "Variación de peso"
      Height          =   255
      Left            =   3960
      TabIndex        =   36
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label19 
      Caption         =   "Variación de diámetro"
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label18 
      Caption         =   "Variación  de  espesor"
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Módulo Final Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "Módulo Inicial Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "Elongación Final Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "Elongación Inicial Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Tracción Final Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Tracción Inicial Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Dureza Final Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Dureza Inicial Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Densidad Final Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Densidad Inicial Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Peso Final Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Peso Inicial Promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Diámetro final promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Diámetro Inicial promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Espesor Final promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Espesor Inicial promedio"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Ensayo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmAEDfinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Form1.Visible = True
frmAEDfinal.Hide

End Sub

Private Sub Command2_Click()
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDconst where cod = " & Text1.Text)
Set rs1 = db.OpenRecordset("Select * from AEDenv where ensayo = " & Text1.Text)
Set rs2 = db.OpenRecordset("Select * from AEDorig where ensayo = " & Text1.Text)






Dim appp As New Excel.Application
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim r As Excel.Range

Set wb = appp.Workbooks.Open("\\Servidor2\e\EntornoBafir\Planillas\informeAED.xls", , True)
Set ws = wb.Worksheets(1)

ws.Cells(12, 3) = rs.Fields("titulo")
ws.Cells(73, 3) = rs.Fields("titulo")

ws.Cells(14, 3) = rs.Fields("proyecto")
ws.Cells(75, 3) = rs.Fields("proyecto")

ws.Cells(14, 6) = "nº " & rs.Fields("cod")
ws.Cells(75, 6) = "nº " & rs.Fields("cod")

ws.Cells(14, 9) = rs.Fields("doc")
ws.Cells(75, 9) = rs.Fields("doc")

ws.Cells(16, 4) = rs1.Fields("fecha")
ws.Cells(77, 4) = rs1.Fields("fecha")

ws.Cells(16, 8) = rs1.Fields("responsable")
ws.Cells(77, 8) = rs1.Fields("responsable")

ws.Cells(20, 8) = rs.Fields("especificacion")
ws.Cells(21, 6) = rs.Fields("norma")
ws.Cells(24, 5) = rs.Fields("material")
ws.Cells(25, 5) = rs.Fields("codigo")
ws.Cells(26, 5) = rs.Fields("material")
ws.Cells(27, 5) = rs.Fields("lote")
ws.Cells(28, 5) = rs2.Fields("fecha")
ws.Cells(29, 5) = rs1.Fields("fecha")

ws.Cells(33, 4) = rs.Fields("temperatura")
ws.Cells(34, 4) = rs.Fields("presion")
ws.Cells(35, 4) = rs.Fields("medio")
ws.Cells(37, 5) = rs.Fields("instrumento")
ws.Cells(38, 5) = rs.Fields("loteinst")
ws.Cells(33, 9) = rs.Fields("ciclos")
ws.Cells(34, 9) = rs.Fields("tiempodeciclo")
ws.Cells(35, 9) = rs.Fields("tiempodedescompresion")


ws.Cells(54, 7) = rs.Fields("alojam")
ws.Cells(55, 7) = rs.Fields("comp")
'ws.Cells(56, 7) = rs.Fields("llenado")

espesorl = (CDbl(rs2.Fields("espesor1")) + CDbl(rs2.Fields("espesor2")) + CDbl(rs2.Fields("espesor3")) + CDbl(rs2.Fields("espesor4")) + CDbl(rs2.Fields("espesor5"))) / 5
diametroinl = (CDbl((rs2.Fields("perimetro1")) + CDbl(rs2.Fields("perimetro2")) + CDbl(rs2.Fields("perimetro3")) + CDbl(rs2.Fields("perimetro4")) + CDbl(rs2.Fields("perimetro5"))) / 5) / 3.14 - (2 * espesorl)
diametroexl = (CDbl((rs2.Fields("perimetro1")) + CDbl(rs2.Fields("perimetro2")) + CDbl(rs2.Fields("perimetro3")) + CDbl(rs2.Fields("perimetro4")) + CDbl(rs2.Fields("perimetro5"))) / 5) / 3.14
volumenoring = ((3.14) ^ 2 * (diametroexl - espesorl) * espesorl ^ 2 / 4) / 1000
volumenranura = 40.01
'volumen de ranura segun V = (De^2 - Di^2)* (PI()/4) * altura dividido mil para cm3
llenado = volumenoring * 100 / volumenranura

ws.Cells(56, 7) = Format(llenado, "0.00") & "%"

ws.Cells(58, 7) = rs.Fields("dimen")
ws.Cells(59, 7) = rs.Fields("volumen")

ws.Cells(43, 4) = rs2.Fields("espesor1")
ws.Cells(43, 5) = rs2.Fields("espesor2")
ws.Cells(43, 6) = rs2.Fields("espesor3")
ws.Cells(43, 7) = rs2.Fields("espesor4")
ws.Cells(43, 8) = rs2.Fields("espesor5")
ws.Cells(43, 9) = (CDbl(rs2.Fields("espesor1")) + CDbl(rs2.Fields("espesor2")) + CDbl(rs2.Fields("espesor3")) + CDbl(rs2.Fields("espesor4")) + CDbl(rs2.Fields("espesor5"))) / 5

'diam1 = (CDbl(rs2.Fields("perimetro1")) / 3.14) - (2 * CDbl(rs2.Fields("espesor1")))
'diam2 = (CDbl(rs2.Fields("perimetro2")) / 3.14) - (2 * CDbl(rs2.Fields("espesor2")))
'diam3 = (CDbl(rs2.Fields("perimetro3")) / 3.14) - (2 * CDbl(rs2.Fields("espesor3")))
'diam4 = (CDbl(rs2.Fields("perimetro4")) / 3.14) - (2 * CDbl(rs2.Fields("espesor4")))
'diam5 = (CDbl(rs2.Fields("perimetro5")) / 3.14) - (2 * CDbl(rs2.Fields("espesor5")))

diam1 = CDbl(rs2.Fields("diamint1"))
diam2 = CDbl(rs2.Fields("diamint2"))
diam3 = CDbl(rs2.Fields("diamint3"))
diam4 = CDbl(rs2.Fields("diamint4"))
diam5 = CDbl(rs2.Fields("diamint5"))

ws.Cells(44, 4) = diam1
ws.Cells(44, 5) = diam2
ws.Cells(44, 6) = diam3
ws.Cells(44, 7) = diam4
ws.Cells(44, 8) = diam5
ws.Cells(44, 9) = (diam1 + diam2 + diam3 + diam4 + diam5) / 5

ws.Cells(45, 4) = rs2.Fields("peso1")
ws.Cells(45, 5) = rs2.Fields("peso2")
ws.Cells(45, 6) = rs2.Fields("peso3")
ws.Cells(45, 7) = rs2.Fields("peso4")
ws.Cells(45, 8) = rs2.Fields("peso5")
ws.Cells(45, 9) = (CDbl(rs2.Fields("peso1")) + CDbl(rs2.Fields("peso2")) + CDbl(rs2.Fields("peso3")) + CDbl(rs2.Fields("peso4")) + CDbl(rs2.Fields("peso5"))) / 5

ws.Cells(46, 4) = CDbl(rs2.Fields("densidad1"))
ws.Cells(46, 5) = CDbl(rs2.Fields("densidad2"))
ws.Cells(46, 6) = CDbl(rs2.Fields("densidad3"))
ws.Cells(46, 7) = CDbl(rs2.Fields("densidad4"))
ws.Cells(46, 8) = CDbl(rs2.Fields("densidad5"))
ws.Cells(46, 9) = (CDbl(rs2.Fields("densidad1")) + CDbl(rs2.Fields("densidad2")) + CDbl(rs2.Fields("densidad3")) + CDbl(rs2.Fields("densidad4")) + CDbl(rs2.Fields("densidad5"))) / 5

ws.Cells(47, 4) = rs2.Fields("dureza1")
ws.Cells(47, 5) = rs2.Fields("dureza2")
ws.Cells(47, 6) = rs2.Fields("dureza3")
ws.Cells(47, 7) = rs2.Fields("dureza4")
ws.Cells(47, 8) = rs2.Fields("dureza5")
ws.Cells(47, 9) = (CDbl(rs2.Fields("dureza1")) + CDbl(rs2.Fields("dureza2")) + CDbl(rs2.Fields("dureza3")) + CDbl(rs2.Fields("dureza4")) + CDbl(rs2.Fields("dureza5"))) / 5

ws.Cells(48, 4) = CDbl(rs2.Fields("MPa1"))

ws.Cells(49, 4) = CDbl(rs2.Fields("elong1"))

ws.Cells(50, 4) = CDbl(rs2.Fields("modulo501"))

If Not Len(rs1.Fields("visual")) >= 255 Then
    ws.Cells(86, 2) = CVar(rs1.Fields("Visual"))
Else
    ws.Cells(86, 2) = "VER INFORME ADJUNTO"
End If
rangi = Len(rs1.Fields("visual"))






ws.Cells(101, 4) = CDbl(rs1.Fields("espesor2"))
ws.Cells(101, 5) = CDbl(rs1.Fields("espesor3"))
ws.Cells(101, 6) = CDbl(rs1.Fields("espesor4"))
ws.Cells(101, 7) = (CDbl(rs1.Fields("espesor2")) + CDbl(rs1.Fields("espesor3")) + CDbl(rs1.Fields("espesor4"))) / 3

t18 = Format(CDbl(Text18.Text), "0.00")
If t18 > 0 Then
t18 = "+" & t18
End If
ws.Cells(101, 8) = t18

ws.Cells(102, 4) = CDbl(rs1.Fields("diamint2"))
ws.Cells(102, 5) = CDbl(rs1.Fields("diamint3"))
ws.Cells(102, 6) = CDbl(rs1.Fields("diamint4"))
ws.Cells(102, 7) = (CDbl(rs1.Fields("diamint2")) + CDbl(rs1.Fields("diamint3")) + CDbl(rs1.Fields("diamint4"))) / 3

t19 = Format(CDbl(Text19.Text), "0.00")
If t19 > 0 Then
t19 = "+" & t19
End If
ws.Cells(102, 8) = t19

ws.Cells(103, 4) = rs1.Fields("peso2")
ws.Cells(103, 5) = rs1.Fields("peso3")
ws.Cells(103, 6) = rs1.Fields("peso4")
ws.Cells(103, 7) = (CDbl(rs1.Fields("peso2")) + CDbl(rs1.Fields("peso3")) + CDbl(rs1.Fields("peso4"))) / 3
t20 = Format(CDbl(Text20.Text), "0.00")
If t20 > 0 Then
t20 = "+" & t20
End If
ws.Cells(103, 8) = t20

ws.Cells(104, 4) = CDbl(rs1.Fields("densidad2"))
ws.Cells(104, 5) = CDbl(rs1.Fields("densidad3"))
ws.Cells(104, 6) = CDbl(rs1.Fields("densidad4"))
ws.Cells(104, 7) = (CDbl(rs1.Fields("densidad2")) + CDbl(rs1.Fields("densidad3")) + CDbl(rs1.Fields("densidad4"))) / 3
t21 = Format(CDbl(Text21.Text), "0.00")
If t21 > 0 Then
t21 = "+" & t21
End If
ws.Cells(104, 8) = t21

ws.Cells(105, 4) = rs1.Fields("dureza2")
ws.Cells(105, 5) = rs1.Fields("dureza3")
ws.Cells(105, 6) = rs1.Fields("dureza4")
ws.Cells(105, 7) = (CDbl(rs1.Fields("dureza2")) + CDbl(rs1.Fields("dureza3")) + CDbl(rs1.Fields("dureza4"))) / 3
t22 = Format(CDbl(Text22.Text), "0.00")
If t22 > 0 Then
t22 = "+" & t22
End If
ws.Cells(105, 8) = t22

ws.Cells(106, 4) = CDbl(rs1.Fields("MPa2"))
t23 = Format(CDbl(Text23.Text), "0.00")
If t23 > 0 Then
t23 = "+" & t23
End If
ws.Cells(106, 8) = t23

ws.Cells(107, 4) = CDbl(rs1.Fields("elong2"))
t24 = Format(CDbl(Text24.Text), "0.00")
If t24 > 0 Then
t24 = "+" & t24
End If
ws.Cells(107, 8) = t24

ws.Cells(108, 4) = CDbl(rs1.Fields("modulo502"))
t25 = Format(CDbl(Text25.Text), "0.00")
If t25 > 0 Then
t25 = "+" & t25
End If
ws.Cells(108, 8) = t25
ws.Cells(112, 2) = Label26.Caption








ws.PrintOut
sdffsf = MsgBox("Presiones Ok cuando la impresión esté finalizada", vbInformation + vbOKOnly, "Imprimiendo")
graba = MsgBox("Desea guardar el informe?", vbInformation + vbYesNo, "Grabar")
If graba = vbYes Then
    CommonDialog1.ShowSave
    ruta = CommonDialog1.FileName & CommonDialog1.Filter

    wb.SaveAs (ruta), "*.xls"
End If
wb.Close (False)



If Len(rs1.Fields("visual")) >= 255 Then

Dim w As New Word.Application


w.Documents.Open ("\\Servidor2\e\EntornoBafir\Planillas\membrete.doc"), , True
w.Documents(1).Activate
w.Selection.TypeText ("4.2 Inspección Visual" & Chr(13) & Chr(13))
w.Selection.TypeText (rs1.Fields("visual"))
'w.ActiveDocument.Range(rangi+1,rangi+1)
'w.ActiveDocument.Shapes.AddPicture "C:\WINDOWS\Escritorio\plaza\060421\www\images\home1.gif", LinkToFile:= _
'        False, SaveWithDocument:=True
'w.ActiveDocument.Shapes.AddPicture "C:\Mis documentos\pablo\1\literatura\Informatica\Visual basic\next.gif", LinkToFile:= _
'        False, SaveWithDocument:=True
'w.ActiveDocument.SaveAs FileName:="c:\windows\escritorio\prueba1.doc"
w.PrintOut
fg = MsgBox("Presione OK cuando se haya terminado de imprimir el informe", vbInformation + vbOKOnly, "Imprimiendo informe visual adjunto")

w.Application.Quit (False)

End If


db.Close
End Sub

Private Sub Command3_Click()



Dim db As Database
Dim rs As Recordset


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDenv where ensayo = " & Text1.Text)

MsgBox (rs.Fields("visual"))



db.Close
End Sub
