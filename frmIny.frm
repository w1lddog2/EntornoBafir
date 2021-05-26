VERSION 5.00
Begin VB.Form frmIny 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productividad de Inyectoras"
   ClientHeight    =   3120
   ClientLeft      =   5655
   ClientTop       =   3015
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Obtener datos de Inyectora:"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmIny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Inicio
Public final

Private Sub Command1_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from Inyectora1 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from Inyectora1")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False ' Or rs.AbsolutePosition <> finalizaren
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren
    
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'If CStr(tiempo1) = "01/09/05 08:19:37 a.m." Then
    'MsgBox ("chan")
    'End If
    
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False 'Xor rs.AbsolutePosition <> finalizaren
     If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        
        contador1 = rs.Fields("cont_total")
        'tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        
                
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from HistInyectora1")
Form2.Hide
rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la Inyectora1", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "1"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide
End Sub

Private Sub Command2_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from inyectora2 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from inyectora2")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren
    
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
    
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        contador1 = rs.Fields("cont_total")
        
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Form2.Hide
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from Histinyectora2")

rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la inyectora2", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "2"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide

End Sub

Private Sub Command3_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from inyectora3 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from inyectora3")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren

If rs.AbsolutePosition = finalizaren Then
    Exit Do
End If
    
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        contador1 = rs.Fields("cont_total")
        
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Form2.Hide

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from Histinyectora3")

rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la inyectora3", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "3"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide

End Sub

Private Sub Command4_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from inyectora4 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from inyectora4")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren

If rs.AbsolutePosition = finalizaren Then
    Exit Do
End If
    
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        contador1 = rs.Fields("cont_total")
        
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Form2.Hide
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from Histinyectora4")

rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la inyectora4", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "4"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide

End Sub



Private Sub Command5_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from inyectora5 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from inyectora5")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren

If rs.AbsolutePosition = finalizaren Then
    Exit Do
End If
    
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        contador1 = rs.Fields("cont_total")
        
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Form2.Hide
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from Histinyectora5")

rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la inyectora5", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "5"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide

End Sub

Private Sub Command6_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from inyectora6 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from inyectora6")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren
If rs.AbsolutePosition = finalizaren Then
    Exit Do
End If
    
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        contador1 = rs.Fields("cont_total")
        
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Form2.Hide
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from Histinyectora6")

rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la inyectora6", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "6"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide

End Sub

Private Sub Command7_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from inyectora7 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from inyectora7")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren
If rs.AbsolutePosition = finalizaren Then
    Exit Do
End If
    
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        contador1 = rs.Fields("cont_total")
        
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Form2.Hide
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from Histinyectora7")

rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la inyectora7", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "7"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide

End Sub

Private Sub Command8_Click()
FrmDateIny.Combo1.Clear
FrmDateIny.Combo2.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha from inyectora8 Group by fecha")
rs.MoveLast
cuantasfechas = rs.RecordCount - 1
rs.MoveFirst
ReDim convierte(cuantasfechas)
For subir = 0 To cuantasfechas
convierte(subir) = CDate(rs.Fields("fecha"))
rs.MoveNext
Next
rs.MoveFirst
'a = rs.RecordCount
'rs.MoveLast
'b = rs.RecordCount
Set rs = db.OpenRecordset("Select fecha from Fechatemp")
For subir = 0 To cuantasfechas
rs.AddNew
rs.Fields("fecha") = convierte(subir)
rs.Update
Next

Set rs = db.OpenRecordset("Select fecha from fechatemp Group by fecha order by fecha")
rs.MoveFirst

Do While rs.EOF = False
     FrmDateIny.Combo1.AddItem (rs.Fields("fecha"))
     rs.MoveNext
Loop

rs.MoveLast
Do While rs.BOF = False
   FrmDateIny.Combo2.AddItem (rs.Fields("fecha"))
    rs.MovePrevious
Loop
Set rs = db.OpenRecordset("Select fecha from fechatemp")
rs.MoveFirst
rs.Edit
Do While rs.EOF = False
rs.Delete
rs.MoveNext
Loop



FrmDateIny.Show (1)
frmIny.Enabled = False

Dim batch00 As Double
Dim batch01 As Double
Dim batch02 As Double
Dim batch03 As Double
Dim batch04 As Double
Dim batch05 As Double
Dim batch06 As Double
Dim batch07 As Double
Dim batch08 As Double
Dim batch09 As Double
Dim batch10 As Double


Set rs = db.OpenRecordset("Select fecha, hora, batch, cont_total from inyectora8")

rs.MoveFirst
Do Until rs.Fields("Fecha") = Inicio
rs.MoveNext
Loop
iniciaren = rs.AbsolutePosition
rs.MoveLast

Do Until rs.Fields("Fecha") = final
rs.MovePrevious
Loop
finalizaren = rs.AbsolutePosition + 1

rs.MoveFirst
rs.Move (iniciaren)

Do While rs.EOF = False
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Value = (rs.AbsolutePosition + 1) * 100 / finalizaren
If rs.AbsolutePosition = finalizaren Then
    Exit Do
End If
    
    tiempo1 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    batch1 = rs.Fields("batch")
    
    Do While rs.EOF = False
    If rs.AbsolutePosition = finalizaren Then
        Exit Do
    End If
     '   If rs.Fields("Cont_total") = 1149 Then
     '   MsgBox ("dont")
     '   End If
        contador1 = rs.Fields("cont_total")
        
        rs.MoveNext
        If rs.EOF = True Then
            batch2 = ""
        Else
            batch2 = rs.Fields("batch")
        End If
        If rs.EOF = True Then
            Exit Do
        End If
        contador2 = rs.Fields("cont_total")
        If contador1 <= contador2 Then
        Else
        Exit Do
        End If
        
        If batch1 <> batch2 Then
            Exit Do
        End If
    Loop
    
    rs.MovePrevious
    
    tiempo2 = CDate(rs.Fields("fecha") & " " & rs.Fields("hora"))
    
    'rs.MovePrevious
    a = Format(tiempo2 - tiempo1, "hh:mm:ss")
      
    
    a = (Hour(a) * 60) + (Minute(a)) + (Second(a) / 60)
    
    If rs.Fields("Batch") = "00          " Then
    batch00 = batch00 + a
    End If
    If rs.Fields("Batch") = "01          " Then
    batch01 = batch01 + a
    End If
    If rs.Fields("Batch") = "02          " Then
    batch02 = batch02 + a
    End If
    If rs.Fields("Batch") = "03          " Then
    batch03 = batch03 + a
    End If
    If rs.Fields("Batch") = "04          " Then
    batch04 = batch04 + a
    End If
    If rs.Fields("Batch") = "05          " Then
    batch05 = batch05 + a
    End If
    If rs.Fields("Batch") = "06          " Then
    batch06 = batch06 + a
    End If
    If rs.Fields("Batch") = "07          " Then
    batch07 = batch07 + a
    End If
    If rs.Fields("Batch") = "08          " Then
    batch08 = batch08 + a
    End If
    If rs.Fields("Batch") = "09          " Then
    batch09 = batch09 + a
    End If
    If rs.Fields("Batch") = "10          " Then
    batch10 = batch10 + a
    End If
    rs.MoveNext
    
    If rs.AbsolutePosition = finalizaren Or rs.EOF = True Then
        Exit Do
    End If
Loop
Form2.Hide
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\prod_iny.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select fecha, batch00, batch01, batch02, batch03,batch04, batch05, batch06,batch07, batch08,batch09,batch10 from Histinyectora8")

rs.AddNew
rs.Fields("fecha") = Inicio & "-" & final
rs.Fields("batch00") = batch00
rs.Fields("batch01") = batch01
rs.Fields("batch02") = batch02
rs.Fields("batch03") = batch03
rs.Fields("batch04") = batch04
rs.Fields("batch05") = batch05
rs.Fields("batch06") = batch06
rs.Fields("batch07") = batch07
rs.Fields("batch08") = batch08
rs.Fields("batch09") = batch09
rs.Fields("batch10") = batch10
rs.Update


db.Close

dfd = MsgBox("Se ha generado el registro de la inyectora8", vbInformation + vbOKOnly, "Registro Guardado")


inyectora = "8"
frmInformeIny.Label26 = Inicio & "-" & final
frmInformeIny.Label14 = inyectora
frmInformeIny.Label15 = Format(batch00, "#.##")
frmInformeIny.Label16 = Format(batch01, "#.##")
frmInformeIny.Label17 = Format(batch02, "#.##")
frmInformeIny.Label18 = Format(batch03, "#.##")
frmInformeIny.Label19 = Format(batch04, "#.##")
frmInformeIny.Label20 = Format(batch05, "#.##")
frmInformeIny.Label21 = Format(batch06, "#.##")
frmInformeIny.Label22 = Format(batch07, "#.##")
frmInformeIny.Label23 = Format(batch08, "#.##")
frmInformeIny.Label24 = Format(batch09, "#.##")
frmInformeIny.Label25 = Format(batch10, "#.##")

frmInformeIny.Show (1)

Form1.Enabled = True
Form1.Visible = True
frmIny.Hide

End Sub

Private Sub Command9_Click()
Form1.Enabled = True
Form1.Visible = True
frmIny.Hide
End Sub

