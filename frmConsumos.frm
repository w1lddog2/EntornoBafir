VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConsumos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumos"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Envios a proveedores"
      Height          =   735
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Consumo individual de materias primas"
      Height          =   735
      Left            =   7680
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consumo individual de mezclas"
      Height          =   735
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exportar a excel"
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   735
      Left            =   9480
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Consumo masivo de Materias Primas"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consumo masivo de mezclas"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   11880
      _Version        =   393216
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmConsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dato As String
Private Sub Command1_Click()
periodoi = InputBox("Ingrese fecha 'desde' a analizar", "Desde")
If periodoi = "" Then
    Exit Sub
End If
If Not IsDate(periodoi) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If
periodof = InputBox("Ingrese fecha 'hasta' a analizar", "Hasta")
If Not IsDate(periodof) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If

Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\centralpesado.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With


sPathBase1 = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
                    
    With cnn1
         'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
         .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase1 & ";" & "Jet OLEDB:Database Password=flanflus;"
         .Open
    End With
    
    
    
    rst.Open "SELECT compuesto FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# group by compuesto order by compuesto asc", cnn, adOpenStatic, adLockReadOnly
    a = rst.RecordCount
    
    ReDim comPuesto(a - 1)
    i = 0
    Do Until rst.EOF
        comPuesto(i) = rst.Fields("compuesto")
        i = i + 1
        rst.MoveNext
    Loop
    rst.Close
    
    numeroregistros = a
    MSFlexGrid1.Rows = numeroregistros + 1
    MSFlexGrid1.Cols = 4
    MSFlexGrid1.TextMatrix(0, 0) = "Compuesto"
    MSFlexGrid1.TextMatrix(0, 1) = "Kilos"
    MSFlexGrid1.TextMatrix(0, 2) = "Costo por kilo $ar"
    MSFlexGrid1.TextMatrix(0, 3) = "Costo total $ar"
            
    Fila = 1
    i = 0
    For g = 0 To UBound(comPuesto)
        
        rst1.Open "SELECT SUM(pesado) as sumapesado FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# and compuesto = '" & comPuesto(i) & "'", cnn, adOpenStatic, adLockReadOnly
        
        p = InStr(1, comPuesto(i), "FD")
        If p <> 0 Then ' COstos para FDs
            costocomp = calculaprecioFD(comPuesto(i))
        Else
            rst2.Open "SELECT COSTO_TOTAL FROM formbase Where N_FORMULA = '" & comPuesto(i) & "'", cnn1, adOpenStatic, adLockReadOnly
            If rst2.RecordCount = 0 Then
                costocomp = 0
            Else
                costocomp = rst2.Fields("COSTO_TOTAL")
            End If
        End If
        MSFlexGrid1.TextMatrix(Fila, 0) = comPuesto(i)
        MSFlexGrid1.TextMatrix(Fila, 1) = Format(rst1.Fields("sumapesado"), "0.00")
        MSFlexGrid1.TextMatrix(Fila, 2) = Format(costocomp, "0.00")
        MSFlexGrid1.TextMatrix(Fila, 3) = Format(CDbl(rst1.Fields("sumapesado")) * costocomp, "0.00")
        
        Fila = Fila + 1
        i = i + 1
        rst1.Close
        If p = 0 Then
            rst2.Close
        End If
    Next
    AutoGrid frmConsumos.MSFlexGrid1
    cnn.Close
    cnn1.Close
End Sub

Private Sub Command2_Click()
periodoi = InputBox("Ingrese fecha 'desde' a analizar", "Desde")
If periodoi = "" Then
    Exit Sub
End If
If Not IsDate(periodoi) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If
periodof = InputBox("Ingrese fecha 'hasta' a analizar", "Hasta")
If Not IsDate(periodof) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If

Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\centralpesado.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With


sPathBase1 = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
                    
    With cnn1
         'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
         .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase1 & ";" & "Jet OLEDB:Database Password=flanflus;"
         .Open
    End With
    
    rst2.Open "SELECT precio FROM producto Where cod_prod = '460'", cnn1, adOpenStatic, adLockReadOnly
    dolar = rst2.Fields("precio")
    rst2.Close
    rst.Open "SELECT Materia_prima FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# group by materia_prima order by materia_prima asc", cnn, adOpenStatic, adLockReadOnly
    a = rst.RecordCount
    
    ReDim materia(a - 1)
    i = 0
    Do Until rst.EOF
        materia(i) = rst.Fields("Materia_prima")
        i = i + 1
        rst.MoveNext
    Loop
    rst.Close
    
    numeroregistros = a
    MSFlexGrid1.Rows = numeroregistros + 1
    MSFlexGrid1.Cols = 5
    MSFlexGrid1.TextMatrix(0, 0) = "Codigo Visual"
    MSFlexGrid1.TextMatrix(0, 1) = "Materia Prima"
    MSFlexGrid1.TextMatrix(0, 2) = "Kilos"
    MSFlexGrid1.TextMatrix(0, 3) = "Costo por kilo $ar"
    MSFlexGrid1.TextMatrix(0, 4) = "Costo total $ar"
            
    Fila = 1
    i = 0
    For g = 0 To UBound(materia)
        
        rst1.Open "SELECT SUM(pesado) as sumapesado FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# and materia_prima = '" & materia(i) & "'", cnn, adOpenStatic, adLockReadOnly
        
        p = InStr(1, materia(i), "FD")
        If p <> 0 Then ' COstos para FDs
            costocomp = 0
        Else
            rst2.Open "SELECT precio FROM producto Where cod_prod = '" & materia(i) & "'", cnn1, adOpenStatic, adLockReadOnly
            costocomp = rst2.Fields("precio")
        End If
        rst.Open "SELECT descrip,COD_VISUAL FROM producto where cod_prod = '" & materia(i) & "'", cnn1, adOpenStatic, adLockReadOnly
        MSFlexGrid1.TextMatrix(Fila, 0) = rst.Fields("cod_visual")
        MSFlexGrid1.TextMatrix(Fila, 1) = rst.Fields("descrip")
        MSFlexGrid1.TextMatrix(Fila, 2) = Format(rst1.Fields("sumapesado"), "0.00")
        MSFlexGrid1.TextMatrix(Fila, 3) = Format((costocomp * dolar), "0.00")
        MSFlexGrid1.TextMatrix(Fila, 4) = Format(CDbl((rst1.Fields("sumapesado")) * costocomp * dolar), "0.00")
        
        Fila = Fila + 1
        i = i + 1
        rst1.Close
        rst.Close
        If p = 0 Then
            rst2.Close
        End If
    Next
    AutoGrid frmConsumos.MSFlexGrid1
    cnn.Close
    cnn1.Close
End Sub

Private Sub Command3_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Command4_Click()
Dim appp As New Excel.Application
Dim ws As New Excel.Worksheet
Dim wb As New Excel.Workbook
Dim r As Excel.Range
Set wb = appp.Workbooks.Add
Set ws = wb.Worksheets.Add
ws.Activate

Fila = frmConsumos.MSFlexGrid1.Rows

columna = frmConsumos.MSFlexGrid1.Cols

For k = 1 To Fila
    For j = 1 To columna
    'MsgBox k
    If k <> 1 Then
        If j <> 1 Then
            If (MSFlexGrid1.TextMatrix(k - 1, j - 1)) = "" Then
                ws.Cells(k, j) = 0
            Else
                ws.Cells(k, j) = CDbl(MSFlexGrid1.TextMatrix(k - 1, j - 1))
            End If
        Else
            ws.Cells(k, j) = "'" & MSFlexGrid1.TextMatrix(k - 1, j - 1)
        End If
    Else
        ws.Cells(k, j) = "'" & MSFlexGrid1.TextMatrix(k - 1, j - 1)
    End If
    
    Next ' columna
Next 'fila


frmConsumos.CommonDialog1.ShowSave
ruta = frmConsumos.CommonDialog1.FileName & frmConsumos.CommonDialog1.Filter
ws.SaveAs ruta
appp.Quit
End Sub

Private Sub Command5_Click()





Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\centralpesado.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With


sPathBase1 = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
                    
    With cnn1
         'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
         .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase1 & ";" & "Jet OLEDB:Database Password=flanflus;"
         .Open
    End With
    
    periodoi = InputBox("Ingrese fecha 'desde' a analizar", "Desde")
If periodoi = "" Then
    Exit Sub
End If
If Not IsDate(periodoi) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If
periodof = InputBox("Ingrese fecha 'hasta' a analizar", "Hasta")
If Not IsDate(periodof) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If
    
    rst.Open "SELECT compuesto FROM pesado group by compuesto order by compuesto asc", cnn, adOpenStatic, adLockReadOnly
        frmconsumeseleccione.Combo1.Clear
    Do Until rst.EOF = True
        frmconsumeseleccione.Combo1.AddItem (rst.Fields("compuesto"))
        rst.MoveNext
    Loop
    rst.Close
    frmconsumeseleccione.Show (1)
    
    buscaa = frmconsumeseleccione.respuesta
        
    rst.Open "SELECT compuesto FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# and compuesto = '" & buscaa & "' group by compuesto", cnn, adOpenStatic, adLockReadOnly
    a = rst.RecordCount
    
    ReDim comPuesto(a - 1)
    i = 0
    Do Until rst.EOF
        comPuesto(i) = rst.Fields("compuesto")
        i = i + 1
        rst.MoveNext
    Loop
    rst.Close
    
    numeroregistros = a
    MSFlexGrid1.Rows = numeroregistros + 1
    MSFlexGrid1.Cols = 4
    MSFlexGrid1.TextMatrix(0, 0) = "Compuesto"
    MSFlexGrid1.TextMatrix(0, 1) = "Kilos"
    MSFlexGrid1.TextMatrix(0, 2) = "Costo por kilo $ar"
    MSFlexGrid1.TextMatrix(0, 3) = "Costo total $ar"
            
    Fila = 1
    i = 0
    For g = 0 To UBound(comPuesto)
        
        rst1.Open "SELECT SUM(pesado) as sumapesado FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# and compuesto = '" & comPuesto(i) & "'", cnn, adOpenStatic, adLockReadOnly
        
        p = InStr(1, comPuesto(i), "FD")
        If p <> 0 Then ' COstos para FDs
            costocomp = calculaprecioFD(comPuesto(i))
        Else
            rst2.Open "SELECT COSTO_TOTAL FROM formbase Where N_FORMULA = '" & comPuesto(i) & "'", cnn1, adOpenStatic, adLockReadOnly
            costocomp = rst2.Fields("COSTO_TOTAL")
        End If
        MSFlexGrid1.TextMatrix(Fila, 0) = comPuesto(i)
        MSFlexGrid1.TextMatrix(Fila, 1) = Format(rst1.Fields("sumapesado"), "0.00")
        MSFlexGrid1.TextMatrix(Fila, 2) = Format(costocomp, "0.00")
        MSFlexGrid1.TextMatrix(Fila, 3) = Format(CDbl(rst1.Fields("sumapesado")) * costocomp, "0.00")
        
        Fila = Fila + 1
        i = i + 1
        rst1.Close
        If p = 0 Then
            rst2.Close
        End If
    Next
    AutoGrid frmConsumos.MSFlexGrid1
    cnn.Close
    cnn1.Close
End Sub

Private Sub Command6_Click()


Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\centralpesado.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With


sPathBase1 = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
                    
    With cnn1
         'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
         .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase1 & ";" & "Jet OLEDB:Database Password=flanflus;"
         .Open
    End With
    
    
    periodoi = InputBox("Ingrese fecha 'desde' a analizar", "Desde")
If periodoi = "" Then
    Exit Sub
End If
If Not IsDate(periodoi) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If
periodof = InputBox("Ingrese fecha 'hasta' a analizar", "Hasta")
If Not IsDate(periodof) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If
    
    rst.Open "SELECT descrip FROM producto order by descrip asc", cnn, adOpenStatic, adLockReadOnly
    frmconsumeseleccione.Combo1.Clear
    
    Do Until rst.EOF = True
        frmconsumeseleccione.Combo1.AddItem (rst.Fields("descrip"))
        rst.MoveNext
    Loop
    rst.Close
    
    frmconsumeseleccione.Show (1)
    buscaa = frmconsumeseleccione.respuesta
    
    rst.Open "SELECT cod_prod, precio FROM producto where descrip = '" & buscaa & "'", cnn, adOpenStatic, adLockReadOnly
    buscaacod = rst.Fields("Cod_prod")
    buscaaprecio = rst.Fields("precio")
    rst.Close
    
    
    rst2.Open "SELECT precio FROM producto Where cod_prod = '460'", cnn1, adOpenStatic, adLockReadOnly
    dolar = rst2.Fields("precio")
    rst2.Close
    
    
    numeroregistros = 1
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 4
    MSFlexGrid1.TextMatrix(0, 0) = "Materia Prima"
    MSFlexGrid1.TextMatrix(0, 1) = "Kilos"
    MSFlexGrid1.TextMatrix(0, 2) = "Costo por kilo $ar"
    MSFlexGrid1.TextMatrix(0, 3) = "Costo total $ar"
       
        
        rst1.Open "SELECT SUM(pesado) as sumapesado FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# and materia_prima = '" & buscaacod & "'", cnn, adOpenStatic, adLockReadOnly
        Fila = 1
        rst.Open "SELECT descrip FROM producto where cod_prod = '" & buscaacod & "'", cnn1, adOpenStatic, adLockReadOnly
        MSFlexGrid1.TextMatrix(Fila, 0) = rst.Fields("descrip")
        MSFlexGrid1.TextMatrix(Fila, 1) = Format(rst1.Fields("sumapesado"), "0.00")
        MSFlexGrid1.TextMatrix(Fila, 2) = Format((buscaaprecio * dolar), "0.00")
        MSFlexGrid1.TextMatrix(Fila, 3) = Format(CDbl((rst1.Fields("sumapesado")) * buscaaprecio * dolar), "0.00")
        rst1.Close
        rst.Close
    
    AutoGrid frmConsumos.MSFlexGrid1
    cnn.Close
    cnn1.Close
End Sub

Private Sub Command7_Click()
periodoi = InputBox("Ingrese fecha 'desde' a analizar", "Desde")
If periodoi = "" Then
    Exit Sub
End If
If Not IsDate(periodoi) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If
periodof = InputBox("Ingrese fecha 'hasta' a analizar", "Hasta")
If Not IsDate(periodof) Then
    sdfsdfsdff = MsgBox("Debe ingresar una fecha válida", vbCritical + vbOKOnly)
    Exit Sub
End If

Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\centralpesado.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With


sPathBase1 = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
                    
    With cnn1
         'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
         .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase1 & ";" & "Jet OLEDB:Database Password=flanflus;"
         .Open
    End With
    rst2.Open "SELECT compuesto FROM pesado Where batch = 'ENVIO' group by compuesto", cnn, adOpenStatic, adLockReadOnly
    frmLista1.Combo1.Clear
    Do Until rst2.EOF = True
        frmLista1.Combo1.AddItem (rst2.Fields("compuesto"))
        rst2.MoveNext
    Loop
    rst2.Close
    frmLista1.Show (1)
    
    proveedor = frmConsumos.dato
    
    rst2.Open "SELECT precio FROM producto Where cod_prod = '460'", cnn1, adOpenStatic, adLockReadOnly
    dolar = rst2.Fields("precio")
    rst2.Close
    rst.Open "SELECT Materia_prima FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# and batch = 'ENVIO' and compuesto = '" & proveedor & "' group by materia_prima order by materia_prima asc", cnn, adOpenStatic, adLockReadOnly
    a = rst.RecordCount
    
    ReDim materia(a - 1)
    i = 0
    Do Until rst.EOF
        materia(i) = rst.Fields("Materia_prima")
        i = i + 1
        rst.MoveNext
    Loop
    rst.Close
    
    numeroregistros = a
    MSFlexGrid1.Rows = numeroregistros + 1
    MSFlexGrid1.Cols = 4
    MSFlexGrid1.TextMatrix(0, 0) = "Materia Prima"
    MSFlexGrid1.TextMatrix(0, 1) = "Kilos"
    MSFlexGrid1.TextMatrix(0, 2) = "Costo por kilo $ar"
    MSFlexGrid1.TextMatrix(0, 3) = "Costo total $ar"
    
            
    Fila = 1
    i = 0
    For g = 0 To UBound(materia)
        
        rst1.Open "SELECT SUM(pesado) as sumapesado FROM pesado Where Fecha >= #" & (Format(periodoi, "MM/DD/YY")) & "# and fecha <= #" & (Format(periodof, "MM/DD/YY")) & "# and materia_prima = '" & materia(i) & "' and batch = 'ENVIO' and compuesto = '" & proveedor & "'", cnn, adOpenStatic, adLockReadOnly
        
        p = InStr(1, materia(i), "FD")
        If p <> 0 Then ' COstos para FDs
            costocomp = 0
        Else
            rst2.Open "SELECT precio FROM producto Where cod_prod = '" & materia(i) & "'", cnn1, adOpenStatic, adLockReadOnly
            costocomp = rst2.Fields("precio")
        End If
        rst.Open "SELECT descrip FROM producto where cod_prod = '" & materia(i) & "'", cnn1, adOpenStatic, adLockReadOnly
        MSFlexGrid1.TextMatrix(Fila, 0) = rst.Fields("descrip")
        MSFlexGrid1.TextMatrix(Fila, 1) = rst1.Fields("sumapesado")
        MSFlexGrid1.TextMatrix(Fila, 2) = Format((costocomp * dolar), "0.00")
        MSFlexGrid1.TextMatrix(Fila, 3) = Format(CDbl((rst1.Fields("sumapesado")) * costocomp * dolar), "0.00")
        
        
        Fila = Fila + 1
        i = i + 1
        rst1.Close
        rst.Close
        If p = 0 Then
            rst2.Close
        End If
    Next
    AutoGrid frmConsumos.MSFlexGrid1
    cnn.Close
    cnn1.Close
End Sub
