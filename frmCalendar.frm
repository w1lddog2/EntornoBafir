VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   4530
   ClientLeft      =   3510
   ClientTop       =   1770
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Quitar"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar Fecha"
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "frmCalendar.frx":0000
      Left            =   5400
      List            =   "frmCalendar.frx":0002
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _Version        =   524288
      _ExtentX        =   8493
      _ExtentY        =   6165
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2005
      Month           =   9
      Day             =   29
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   -1  'True
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "hh:mm"
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmBack.Enabled = True
frmBack.Visible = True
frmCalendar.Hide
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
gfg = MsgBox("Debe ingresar una hora", vbCritical + vbOKOnly, "Error")
Text1.SetFocus
Exit Sub
End If


Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")

Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha Where Fecha = '" & Calendar1.Value & "'")
If rs.RecordCount <> 0 Then
    fsdf = MsgBox("En este día ya se ha programado un back up", vbInformation + vbOKOnly, "Día ya programado")
    Exit Sub
End If


Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha")

rs.AddNew
rs.Fields("Fecha") = Calendar1.Value
rs.Fields("Hora") = Format(Text1.Text, "hh:mm:ss")
rs.Update
rs.MoveFirst

ordena_fechas

Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha")

List1.Clear

Do
List1.AddItem (rs.Fields("Fecha"))
rs.MoveNext
Loop Until rs.EOF = True
db.Close
End Sub

Private Sub Command4_Click()

'List1.RemoveItem List.listindex

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha Where fecha = '" & List1.Text & "'")


rs.Delete

List1.RemoveItem List1.ListIndex
db.Close
End Sub

Sub ordena_fechas()
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha") ' Where fecha = " & CDate(List1.Text))
Set rs1 = db.OpenRecordset("Select fecha, hora from fecha_temp")

rs.MoveFirst

Do Until rs.EOF = True
rs1.AddNew
rs1.Fields("fecha") = rs.Fields("Fecha")
rs1.Fields("Hora") = rs.Fields("hora")
rs1.Update
rs.MoveNext
Loop

Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha") ' Where fecha = " & CDate(List1.Text))
Set rs1 = db.OpenRecordset("Select fecha, hora from fecha_temp order by fecha, hora")

rs1.MoveFirst
rs.MoveFirst
Do Until rs.EOF = True
rs.Delete
rs.MoveNext
Loop
Do Until rs1.EOF = True
rs.AddNew
rs.Fields("fecha") = rs1.Fields("fecha")
rs.Fields("hora") = rs1.Fields("hora")
rs.Update
rs1.MoveNext
Loop

rs1.MoveFirst
Do Until rs1.EOF = True
rs1.MoveFirst
If rs1.EOF = True Then
    Exit Do
End If
rs1.Delete
Loop
db.Close
End Sub
Private Sub List1_Click()
a = 0
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha") ' Where fecha = " & CDate(List1.Text))

rs.MoveFirst
Do Until a = 1
If CStr(rs.Fields("fecha")) = List1.Text Then
a = 1
Else
rs.MoveNext
End If
Loop

Text1.Text = rs.Fields("Hora")

db.Close
End Sub
