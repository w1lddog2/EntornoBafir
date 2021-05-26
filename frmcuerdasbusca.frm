VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcuerdasbusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Cuerdas"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton Option2 
         Caption         =   "por cuerda"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "por Proveedor y Cuerda"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   3735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cuerda"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmcuerdasbusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
    If Combo1.Text = "" Or Combo2.Text = "" Then
        ffsff = MsgBox("Debe ingresar un valor", vbCritical + vbOKOnly, "Error")
        Combo1.SetFocus
        Exit Sub
    Else
        Dim db As Database
        Dim rs As Recordset
        
        Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
        Set rs = db.OpenRecordset("Select * From cuerdas where cuerda = '" & Combo1.Text & "' And proveedor = '" & Combo2.Text & "'")
        
        If rs.RecordCount = 0 Then
        mfss = MsgBox("No se han encontrado items correspondientes a su busqueda", vbCritical + vbOKOnly, "Busqueda")
        Combo1.SetFocus
        Exit Sub
        Else
            
            rs.MoveFirst
            rs.MoveLast
            numero = rs.RecordCount
            MSFlexGrid1.Rows = numero + 1
            
            MSFlexGrid1.TextMatrix(0, 0) = "Fecha"
            MSFlexGrid1.TextMatrix(0, 1) = "Proveedor"
            MSFlexGrid1.TextMatrix(0, 2) = "Remito"
            MSFlexGrid1.TextMatrix(0, 3) = "Lote"
            MSFlexGrid1.TextMatrix(0, 4) = "Cuerda"
            MSFlexGrid1.TextMatrix(0, 5) = "Material"
            MSFlexGrid1.TextMatrix(0, 6) = "Informe"
            MSFlexGrid1.TextMatrix(0, 7) = "Aprobado"
            MSFlexGrid1.TextMatrix(0, 8) = "Responsable"
            MSFlexGrid1.TextMatrix(0, 9) = "Metros"
            
            rs.MoveFirst
            For asd = 1 To numero
            MSFlexGrid1.TextMatrix(asd, 0) = rs.Fields("Fecha")
            MSFlexGrid1.TextMatrix(asd, 1) = rs.Fields("Proveedor")
            MSFlexGrid1.TextMatrix(asd, 2) = rs.Fields("Remito")
            MSFlexGrid1.TextMatrix(asd, 3) = rs.Fields("Lote")
            MSFlexGrid1.TextMatrix(asd, 4) = rs.Fields("Cuerda")
            MSFlexGrid1.TextMatrix(asd, 5) = rs.Fields("Material")
            MSFlexGrid1.TextMatrix(asd, 6) = rs.Fields("Informe_realizado")
            If rs.Fields("aprovado").Value = True Then
            MSFlexGrid1.TextMatrix(asd, 7) = "Si"
            Else
            MSFlexGrid1.TextMatrix(asd, 7) = "No"
            End If
            MSFlexGrid1.TextMatrix(asd, 8) = rs.Fields("Responsable")
            MSFlexGrid1.TextMatrix(asd, 9) = rs.Fields("Metros")
            rs.MoveNext
            Next
        End If
    
    
    
    
    End If
End If
If Option2.Value = True Then
    If Combo1.Text = "" Then
        sdfsaf = MsgBox("Debe ingresar un valor", vbCritical + vbOKOnly, "Error")
        Combo1.SetFocus
        Exit Sub
    Else
        'Dim db As Database
        'Dim rs As Recordset
        
        Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
        Set rs = db.OpenRecordset("Select * From cuerdas where cuerda = '" & Combo1.Text & "'")
            
            If rs.RecordCount = 0 Then
                
                sdfsfwd = MsgBox("No se han encontrado resultados", vbCritical + vbOKOnly, "Busqueda")
                Combo1.SetFocus
                Exit Sub
            Else
          
            MSFlexGrid1.TextMatrix(0, 0) = "Fecha"
            MSFlexGrid1.TextMatrix(0, 1) = "Proveedor"
            MSFlexGrid1.TextMatrix(0, 2) = "Remito"
            MSFlexGrid1.TextMatrix(0, 3) = "Lote"
            MSFlexGrid1.TextMatrix(0, 4) = "Cuerda"
            MSFlexGrid1.TextMatrix(0, 5) = "Material"
            MSFlexGrid1.TextMatrix(0, 6) = "Informe"
            MSFlexGrid1.TextMatrix(0, 7) = "Aprobado"
            MSFlexGrid1.TextMatrix(0, 8) = "Responsable"
            MSFlexGrid1.TextMatrix(0, 9) = "Metros"
            rs.MoveLast
            numero = rs.RecordCount
            MSFlexGrid1.Rows = numero + 1
            rs.MoveFirst
            For asd = 1 To numero
            MSFlexGrid1.TextMatrix(asd, 0) = rs.Fields("Fecha")
            MSFlexGrid1.TextMatrix(asd, 1) = rs.Fields("Proveedor")
            MSFlexGrid1.TextMatrix(asd, 2) = rs.Fields("Remito")
            MSFlexGrid1.TextMatrix(asd, 3) = rs.Fields("Lote")
            MSFlexGrid1.TextMatrix(asd, 4) = rs.Fields("Cuerda")
            MSFlexGrid1.TextMatrix(asd, 5) = rs.Fields("Material")
            MSFlexGrid1.TextMatrix(asd, 6) = rs.Fields("Informe_realizado")
            If rs.Fields("aprovado").Value = True Then
            MSFlexGrid1.TextMatrix(asd, 7) = "Si"
            Else
            MSFlexGrid1.TextMatrix(asd, 7) = "No"
            End If
            MSFlexGrid1.TextMatrix(asd, 8) = rs.Fields("Responsable")
            MSFlexGrid1.TextMatrix(asd, 9) = rs.Fields("Metros")
            rs.MoveNext
            Next
            End If
            
    
    End If
End If
AutoGrid MSFlexGrid1
db.Close
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Visible = True
frmcuerdasbusca.Hide
End Sub

Private Sub Option1_Click()
Combo1.Enabled = True
Combo2.Enabled = True
End Sub

Private Sub Option2_Click()
Combo2.Enabled = False
End Sub
