VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInformefondos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Fondos"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   ".xls"
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exportar a excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Editar registro seleccionado"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Modificar"
      Height          =   735
      Left            =   6720
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Informe por fondos/conceptos día por día"
      Height          =   735
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Informe por fondos"
      Height          =   735
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Informe por conceptos"
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7646
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Informe por fondos/conceptos"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   9960
      TabIndex        =   4
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Período"
      Height          =   255
      Left            =   9960
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmInformefondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()



inicial = InputBox("Ingrese fecha de búsqueda inicial", "Fecha Inicial")
final = InputBox("Ingrese fecha de búsqueda final", "Fecha Final")

Label2.Caption = inicial & " - " & final

inicial = Format(inicial, "YYYY/MM/DD")
final = Format(final, "YYYY/MM/DD")

a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
Dim rcset1 As ADODB.Recordset
Set rcset1 = New ADODB.Recordset
rcset1.CursorLocation = adUseClient
Dim rcset2 As ADODB.Recordset
Set rcset2 = New ADODB.Recordset
rcset2.CursorLocation = adUseClient
Dim rcset3 As ADODB.Recordset
Set rcset3 = New ADODB.Recordset
rcset3.CursorLocation = adUseClient

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close

rcset.Open "SELECT * FROM flujo_fondos_asiento where fecha between '" & inicial & "' and '" & final & "'", CONN, adOpenStatic, adLockReadOnly

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset.EOF = True
    rcset1.AddNew
    rcset1.Fields("fecha") = rcset.Fields("fecha")
    rcset1.Fields("fondo") = rcset.Fields("fondo")
    rcset1.Fields("concepto") = rcset.Fields("concepto")
    rcset1.Fields("monto") = rcset.Fields("monto")
    rcset1.Update
    rcset.MoveNext
Loop
rcset.Close
rcset1.Close
rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close
rcset3.Open "SELECT * FROM flujo_fondos_asiento_temp3", CONN, adOpenStatic, adLockOptimistic
Do Until rcset3.EOF = True
    rcset3.Delete
    rcset3.MoveNext
Loop
rcset3.Close


'group fondo
rcset.Open "SELECT fondo FROM flujo_fondos_asiento_temp1 group by fondo", CONN, adOpenStatic, adLockReadOnly
'group concepto
rcset2.Open "SELECT concepto FROM flujo_fondos_asiento_temp1 group by concepto", CONN, adOpenStatic, adLockReadOnly
'temporal 2
rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockOptimistic
'seleccion en bloque
Do Until rcset.EOF = True 'loop fondo
    rcset2.MoveFirst
    Do Until rcset2.EOF = True 'loop concepto
    
    rcset3.Open "SELECT SUM(monto) as suma FROM flujo_fondos_asiento_temp1 where fondo = '" & rcset.Fields("fondo") & "' AND concepto = '" & rcset2.Fields("concepto") & "'", CONN, adOpenStatic, adLockReadOnly
    If IsNumeric(rcset3.Fields("suma")) Then
        rcset1.AddNew
        rcset1.Fields("fondo") = rcset.Fields("fondo")
        rcset1.Fields("concepto") = rcset2.Fields("concepto")
        rcset1.Fields("monto") = rcset3.Fields("suma")
        rcset1.Fields("fecha") = Date
        rcset1.Update
        
    End If
        rcset3.Close
    rcset2.MoveNext
    Loop 'loop concepto
rcset.MoveNext
Loop 'loop fondo
rcset1.Close
rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockReadOnly
filas = rcset1.RecordCount

MSFlexGrid1.Rows = filas + 1
MSFlexGrid1.Cols = 3

MSFlexGrid1.TextMatrix(0, 0) = "Fondo"
MSFlexGrid1.TextMatrix(0, 1) = "Concepto"
MSFlexGrid1.TextMatrix(0, 2) = "Monto"
Fila = 1
Do Until rcset1.EOF = True
    rcset.Close
    rcset.Open "SELECT fondo FROM flujo_fondos_fondos where codigo = " & rcset1.Fields("fondo"), CONN, adOpenStatic, adLockReadOnly
    MSFlexGrid1.TextMatrix(Fila, 0) = rcset.Fields("fondo")
    rcset.Close
    rcset.Open "SELECT concepto FROM flujo_fondos_concepto where codigo = " & rcset1.Fields("concepto"), CONN, adOpenStatic, adLockReadOnly
    MSFlexGrid1.TextMatrix(Fila, 1) = rcset.Fields("concepto")
    MSFlexGrid1.TextMatrix(Fila, 2) = rcset1.Fields("monto")
    rcset1.MoveNext
    Fila = Fila + 1
Loop
rcset1.Close
'rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
rcset1.Open "TRUNCATE TABLE `flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic

'Do Until rcset1.EOF = True
'    rcset1.Delete
'    rcset1.MoveNext
'Loop
'rcset1.Close

CONN.Close
MSFlexGrid1.ColWidth(0) = 5000
MSFlexGrid1.ColWidth(1) = 2500
MSFlexGrid1.ColWidth(2) = 2500
Command8.Enabled = True
Command7.Visible = False
End Sub

Private Sub Command2_Click()
Command8.Enabled = False
Form1.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click() ' por conceptos
inicial = InputBox("Ingrese fecha de búsqueda inicial", "Fecha Inicial")
final = InputBox("Ingrese fecha de búsqueda final", "Fecha Final")

Label2.Caption = inicial & " - " & final

inicial = Format(inicial, "YYYY/MM/DD")
final = Format(final, "YYYY/MM/DD")

a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
Dim rcset1 As ADODB.Recordset
Set rcset1 = New ADODB.Recordset
rcset1.CursorLocation = adUseClient
Dim rcset2 As ADODB.Recordset
Set rcset2 = New ADODB.Recordset
rcset2.CursorLocation = adUseClient
Dim rcset3 As ADODB.Recordset
Set rcset3 = New ADODB.Recordset
rcset3.CursorLocation = adUseClient

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close

rcset.Open "SELECT * FROM flujo_fondos_asiento where fecha between '" & inicial & "' and '" & final & "'", CONN, adOpenStatic, adLockReadOnly

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset.EOF = True
    rcset1.AddNew
    rcset1.Fields("fecha") = rcset.Fields("fecha")
    rcset1.Fields("fondo") = rcset.Fields("fondo")
    rcset1.Fields("concepto") = rcset.Fields("concepto")
    rcset1.Fields("monto") = rcset.Fields("monto")
    rcset1.Update
    rcset.MoveNext
Loop
rcset.Close
rcset1.Close
rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close
rcset3.Open "SELECT * FROM flujo_fondos_asiento_temp3", CONN, adOpenStatic, adLockOptimistic
Do Until rcset3.EOF = True
    rcset3.Delete
    rcset3.MoveNext
Loop
rcset3.Close


'group concepto
rcset.Open "SELECT concepto FROM flujo_fondos_asiento_temp1 group by concepto", CONN, adOpenStatic, adLockReadOnly
filas = rcset.RecordCount
MSFlexGrid1.Rows = filas + 1
MSFlexGrid1.Cols = 2
MSFlexGrid1.TextMatrix(0, 0) = "Concepto"
MSFlexGrid1.TextMatrix(0, 1) = "Monto"
Fila = 1
Do Until rcset.EOF = True
    rcset3.Open "SELECT SUM(monto) as suma FROM flujo_fondos_asiento_temp1 where concepto = '" & rcset.Fields("concepto") & "'", CONN, adOpenStatic, adLockReadOnly
    rcset1.Open "SELECT concepto FROM flujo_fondos_concepto where codigo = " & rcset.Fields("concepto"), CONN, adOpenStatic, adLockReadOnly
    MSFlexGrid1.TextMatrix(Fila, 0) = rcset1.Fields("concepto")
    MSFlexGrid1.TextMatrix(Fila, 1) = rcset3.Fields("suma")
    rcset.MoveNext
    Fila = Fila + 1
    rcset1.Close
    rcset3.Close
Loop

'rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
rcset1.Open "TRUNCATE TABLE `flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
'Do Until rcset1.EOF = True
'    rcset1.Delete
'    rcset1.MoveNext
'Loop
'rcset1.Close

CONN.Close
MSFlexGrid1.ColWidth(0) = 5000
MSFlexGrid1.ColWidth(1) = 2500
Command8.Enabled = True
Command7.Visible = False
End Sub

Private Sub Command4_Click()
inicial = InputBox("Ingrese fecha de búsqueda inicial", "Fecha Inicial")
final = InputBox("Ingrese fecha de búsqueda final", "Fecha Final")

Label2.Caption = inicial & " - " & final

inicial = Format(inicial, "YYYY/MM/DD")
final = Format(final, "YYYY/MM/DD")

a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
Dim rcset1 As ADODB.Recordset
Set rcset1 = New ADODB.Recordset
rcset1.CursorLocation = adUseClient
Dim rcset2 As ADODB.Recordset
Set rcset2 = New ADODB.Recordset
rcset2.CursorLocation = adUseClient
Dim rcset3 As ADODB.Recordset
Set rcset3 = New ADODB.Recordset
rcset3.CursorLocation = adUseClient

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close

rcset.Open "SELECT * FROM flujo_fondos_asiento where fecha between '" & inicial & "' and '" & final & "'", CONN, adOpenStatic, adLockReadOnly

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset.EOF = True
    rcset1.AddNew
    rcset1.Fields("fecha") = rcset.Fields("fecha")
    rcset1.Fields("fondo") = rcset.Fields("fondo")
    rcset1.Fields("concepto") = rcset.Fields("concepto")
    rcset1.Fields("monto") = rcset.Fields("monto")
    rcset1.Update
    rcset.MoveNext
Loop
rcset.Close
rcset1.Close
rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close
rcset3.Open "SELECT * FROM flujo_fondos_asiento_temp3", CONN, adOpenStatic, adLockOptimistic
Do Until rcset3.EOF = True
    rcset3.Delete
    rcset3.MoveNext
Loop
rcset3.Close


'group fondo
rcset.Open "SELECT fondo FROM flujo_fondos_asiento_temp1 group by fondo", CONN, adOpenStatic, adLockReadOnly
filas = rcset.RecordCount
MSFlexGrid1.Rows = filas + 1
MSFlexGrid1.Cols = 2
MSFlexGrid1.TextMatrix(0, 0) = "Fondo"
MSFlexGrid1.TextMatrix(0, 1) = "Monto"
Fila = 1
Do Until rcset.EOF = True
    rcset3.Open "SELECT SUM(monto) as suma FROM flujo_fondos_asiento_temp1 where fondo = '" & rcset.Fields("fondo") & "'", CONN, adOpenStatic, adLockReadOnly
    rcset1.Open "SELECT fondo FROM flujo_fondos_fondos where codigo = " & rcset.Fields("fondo"), CONN, adOpenStatic, adLockReadOnly
    MSFlexGrid1.TextMatrix(Fila, 0) = rcset1.Fields("fondo")
    MSFlexGrid1.TextMatrix(Fila, 1) = rcset3.Fields("suma")
    rcset.MoveNext
    Fila = Fila + 1
    rcset1.Close
    rcset3.Close
Loop

'rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
rcset1.Open "TRUNCATE TABLE `flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
'Do Until rcset1.EOF = True
'    rcset1.Delete
'    rcset1.MoveNext
'Loop
'rcset1.Close

CONN.Close
MSFlexGrid1.ColWidth(0) = 5000
MSFlexGrid1.ColWidth(1) = 2500
Command8.Enabled = True
Command7.Visible = False
End Sub

Private Sub Command5_Click() 'PERO DIA POR DIA



inicial = InputBox("Ingrese fecha de búsqueda inicial", "Fecha Inicial")
final = InputBox("Ingrese fecha de búsqueda final", "Fecha Final")

Label2.Caption = inicial & " - " & final

inicial = Format(inicial, "YYYY/MM/DD")
final = Format(final, "YYYY/MM/DD")

a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
Dim rcset1 As ADODB.Recordset
Set rcset1 = New ADODB.Recordset
rcset1.CursorLocation = adUseClient
Dim rcset2 As ADODB.Recordset
Set rcset2 = New ADODB.Recordset
rcset2.CursorLocation = adUseClient
Dim rcset3 As ADODB.Recordset
Set rcset3 = New ADODB.Recordset
rcset3.CursorLocation = adUseClient
Dim rcset4 As ADODB.Recordset
Set rcset4 = New ADODB.Recordset
rcset4.CursorLocation = adUseClient

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close

rcset.Open "SELECT * FROM flujo_fondos_asiento where fecha between '" & inicial & "' and '" & final & "'", CONN, adOpenStatic, adLockReadOnly
If rcset.RecordCount = 0 Then
    sdfsdf = MsgBox("No se han encontrado registros", vbInformation + vbOKOnly, "Sin registros")
    Exit Sub
End If
rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset.EOF = True
    rcset1.AddNew
    rcset1.Fields("fecha") = rcset.Fields("fecha")
    rcset1.Fields("fondo") = rcset.Fields("fondo")
    rcset1.Fields("concepto") = rcset.Fields("concepto")
    rcset1.Fields("monto") = rcset.Fields("monto")
    rcset1.Update
    rcset.MoveNext
Loop
rcset.Close
rcset1.Close
rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close
rcset3.Open "SELECT * FROM flujo_fondos_asiento_temp3", CONN, adOpenStatic, adLockOptimistic
Do Until rcset3.EOF = True
    rcset3.Delete
    rcset3.MoveNext
Loop
rcset3.Close
'group fecha
rcset4.Open "SELECT fecha FROM flujo_fondos_asiento_temp1 group by fecha order by fecha desc", CONN, adOpenStatic, adLockReadOnly
'group fondo
rcset.Open "SELECT fondo FROM flujo_fondos_asiento_temp1 group by fondo", CONN, adOpenStatic, adLockReadOnly
'group concepto
rcset2.Open "SELECT concepto FROM flujo_fondos_asiento_temp1 group by concepto", CONN, adOpenStatic, adLockReadOnly
'temporal 2

'seleccion en bloque
Do Until rcset4.EOF = True ' loop fecha
    rcset.MoveFirst
    Do Until rcset.EOF = True 'loop fondo
        rcset2.MoveFirst
        Do Until rcset2.EOF = True 'loop concepto
        
        rcset3.Open "SELECT SUM(monto) as suma FROM flujo_fondos_asiento_temp1 where fondo = '" & rcset.Fields("fondo") & "' AND concepto = '" & rcset2.Fields("concepto") & "' AND fecha = '" & Format(rcset4.Fields("fecha"), "YYYY/MM/DD") & "'", CONN, adOpenStatic, adLockReadOnly
        If IsNumeric(rcset3.Fields("suma")) Then
            'rcset1.Close
            rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockOptimistic
            rcset1.AddNew
            rcset1.Fields("fondo") = rcset.Fields("fondo")
            rcset1.Fields("concepto") = rcset2.Fields("concepto")
            rcset1.Fields("monto") = rcset3.Fields("suma")
            rcset1.Fields("fecha") = rcset4.Fields("fecha")
            rcset1.Update
            rcset1.Close
        End If
            rcset3.Close
        rcset2.MoveNext 'concepto
        Loop 'loop concepto
    rcset.MoveNext 'fondo
    Loop 'loop fondo
    'rcset1.Close
    rcset4.MoveNext
Loop 'fecha
    
    rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp2", CONN, adOpenStatic, adLockReadOnly
    filas = rcset1.RecordCount
    
    MSFlexGrid1.Rows = filas + 1
    MSFlexGrid1.Cols = 4
    
    MSFlexGrid1.TextMatrix(0, 0) = "Fecha"
    MSFlexGrid1.TextMatrix(0, 1) = "Fondo"
    MSFlexGrid1.TextMatrix(0, 2) = "Concepto"
    MSFlexGrid1.TextMatrix(0, 3) = "Monto"
    Fila = 1
    Do Until rcset1.EOF = True
        rcset.Close
        MSFlexGrid1.TextMatrix(Fila, 0) = rcset1.Fields("fecha")
        rcset.Open "SELECT fondo FROM flujo_fondos_fondos where codigo = " & rcset1.Fields("fondo"), CONN, adOpenStatic, adLockReadOnly
        MSFlexGrid1.TextMatrix(Fila, 1) = rcset.Fields("fondo")
        rcset.Close
        rcset.Open "SELECT concepto FROM flujo_fondos_concepto where codigo = " & rcset1.Fields("concepto"), CONN, adOpenStatic, adLockReadOnly
        MSFlexGrid1.TextMatrix(Fila, 2) = rcset.Fields("concepto")
        MSFlexGrid1.TextMatrix(Fila, 3) = rcset1.Fields("monto")
        rcset1.MoveNext
        Fila = Fila + 1
    Loop

rcset1.Close
'rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
rcset1.Open "TRUNCATE TABLE `flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic

'
'rcset1.MoveFirst
'Do Until rcset1.EOF = True
'    rcset1.Delete
'    rcset1.MoveNext
'Loop
'rcset1.Close

CONN.Close
MSFlexGrid1.ColWidth(0) = 2500
MSFlexGrid1.ColWidth(1) = 5000
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.ColWidth(3) = 2500
Command8.Enabled = True
Command7.Visible = False
End Sub

Private Sub Command6_Click() 'modificar
inicial = InputBox("Ingrese fecha de búsqueda inicial", "Fecha Inicial")
final = InputBox("Ingrese fecha de búsqueda final", "Fecha Final")

Label2.Caption = inicial & " - " & final

inicial = Format(inicial, "YYYY/MM/DD")
final = Format(final, "YYYY/MM/DD")

a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
Dim rcset1 As ADODB.Recordset
Set rcset1 = New ADODB.Recordset
rcset1.CursorLocation = adUseClient
Dim rcset2 As ADODB.Recordset
Set rcset2 = New ADODB.Recordset
rcset2.CursorLocation = adUseClient
Dim rcset3 As ADODB.Recordset
Set rcset3 = New ADODB.Recordset
rcset3.CursorLocation = adUseClient

rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
Do Until rcset1.EOF = True
    rcset1.Delete
    rcset1.MoveNext
Loop
rcset1.Close

rcset.Open "SELECT * FROM flujo_fondos_asiento where fecha between '" & inicial & "' and '" & final & "' order by fecha desc", CONN, adOpenStatic, adLockReadOnly
filas = rcset.RecordCount

MSFlexGrid1.Rows = filas + 1
MSFlexGrid1.Cols = 5
MSFlexGrid1.TextMatrix(0, 0) = "Fecha"
MSFlexGrid1.TextMatrix(0, 1) = "Fondo"
MSFlexGrid1.TextMatrix(0, 2) = "Concepto"
MSFlexGrid1.TextMatrix(0, 3) = "Monto"
MSFlexGrid1.TextMatrix(0, 4) = "Codigo"
Fila = 1
Do Until rcset.EOF = True
    MSFlexGrid1.TextMatrix(Fila, 0) = rcset.Fields("fecha")
    rcset1.Open "SELECT fondo FROM flujo_fondos_fondos where codigo = " & rcset.Fields("fondo"), CONN, adOpenStatic, adLockReadOnly
    MSFlexGrid1.TextMatrix(Fila, 1) = rcset1.Fields("fondo")
    rcset1.Close
    rcset1.Open "SELECT concepto FROM flujo_fondos_concepto where codigo = " & rcset.Fields("concepto"), CONN, adOpenStatic, adLockReadOnly
    MSFlexGrid1.TextMatrix(Fila, 2) = rcset1.Fields("concepto")
    rcset1.Close
    MSFlexGrid1.TextMatrix(Fila, 3) = rcset.Fields("monto")
    MSFlexGrid1.TextMatrix(Fila, 4) = rcset.Fields("codigo")
    rcset.MoveNext
    Fila = Fila + 1
Loop
rcset.Close
'rcset1.Open "SELECT * FROM flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic
rcset1.Open "TRUNCATE TABLE `flujo_fondos_asiento_temp1", CONN, adOpenStatic, adLockOptimistic

'Do Until rcset1.EOF = True
'    rcset1.Delete
'    rcset1.MoveNext
'Loop
'rcset1.Close

CONN.Close
MSFlexGrid1.ColWidth(0) = 2500
MSFlexGrid1.ColWidth(1) = 5000
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.ColWidth(3) = 2500
Command8.Enabled = True
Command7.Visible = True
End Sub

Private Sub Command7_Click() 'editar registro

a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
Dim rcset1 As ADODB.Recordset
Set rcset1 = New ADODB.Recordset
rcset1.CursorLocation = adUseClient
Dim rcset2 As ADODB.Recordset
Set rcset2 = New ADODB.Recordset
rcset2.CursorLocation = adUseClient
Dim rcset3 As ADODB.Recordset
Set rcset3 = New ADODB.Recordset
rcset3.CursorLocation = adUseClient

frmAsientoMod.Combo2.Clear
frmAsientoMod.Combo1.Clear
''''''llena los combos
rcset.Open "SELECT concepto FROM flujo_fondos_concepto order by concepto asc", CONN, adOpenStatic, adLockReadOnly
Do Until rcset.EOF = True
    frmAsientoMod.Combo1.AddItem (rcset.Fields("concepto"))
    rcset.MoveNext
Loop
rcset.Close
rcset.Open "SELECT fondo FROM flujo_fondos_fondos order by fondo asc", CONN, adOpenStatic, adLockReadOnly
Do Until rcset.EOF = True
    frmAsientoMod.Combo2.AddItem (rcset.Fields("fondo"))
    rcset.MoveNext
Loop
rcset.Close
''''''llena los combos

rcset.Open "SELECT * FROM flujo_fondos_asiento where codigo = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4), CONN, adOpenStatic, adLockReadOnly
frmAsientoMod.Text1.Text = rcset.Fields("fecha")
frmAsientoMod.Text2.Text = rcset.Fields("monto")
frmAsientoMod.Text3.Text = rcset.Fields("codigo")
rcset1.Open "SELECT concepto FROM flujo_fondos_concepto where codigo = " & rcset.Fields("concepto"), CONN, adOpenStatic, adLockReadOnly
frmAsientoMod.Combo1.Text = rcset1.Fields("concepto")
rcset1.Close
rcset1.Open "SELECT fondo FROM flujo_fondos_fondos where codigo = " & rcset.Fields("fondo"), CONN, adOpenStatic, adLockReadOnly
frmAsientoMod.Combo2.Text = rcset1.Fields("fondo")
rcset1.Close
Me.Enabled = False
frmAsientoMod.Show
frmAsientoMod.Enabled = True
frmAsientoMod.Text1.SetFocus
CONN.Close
End Sub

Private Sub Command8_Click() ' exportar a excel
CommonDialog1.ShowSave
ruta = CommonDialog1.FileName & CommonDialog1.Filter


Form2.Show
Form2.ProgressBar1.Visible = True
Form2.MousePointer = 11

Dim appp As New Excel.Application
Dim ws As New Excel.Worksheet
Dim wb As New Excel.Workbook
Dim r As Excel.Range
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset


Set wb = appp.Workbooks.Add
Set ws = wb.Worksheets.Add

Columnas = MSFlexGrid1.Cols
ultimo = MSFlexGrid1.Rows
ws.Activate



    For contadorcol = 1 To Columnas
        For contadorfil = 1 To ultimo
            ws.Cells(contadorfil, contadorcol) = MSFlexGrid1.TextMatrix(contadorfil - 1, contadorcol - 1)
            Form2.ProgressBar1.Value = contadorfila / ultimo * 100
        Next
    Next


Form2.Hide

ws.SaveAs ruta
appp.Quit
Form1.Enabled = True
rt = MsgBox("Se ha grabado el archivo como " & ruta, vbInformation + vbOKOnly, "Archivo Guardado")
Form2.ProgressBar1.Visible = False
End Sub
