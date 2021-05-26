VERSION 5.00
Begin VB.Form frmAEDDOC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observaciones"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmAEDDOC.frx":0000
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmAEDDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public formUlario
Public AEDstr As String

Private Sub Command1_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
    sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"
        With cnn
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
            .Open
        End With
        rst.Open AEDstr, cnn, adOpenStatic, adLockOptimistic
        
        
        rst.Fields("observaciones") = Text1.Text
        rst.Update
        cnn.Close
        
Command1.Enabled = False
Command2.Enabled = True
Text1.Enabled = False
End Sub

Private Sub Command2_Click()
Text1.Enabled = True
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Dim tmpform As Form
Dim tform As Form
        
For Each tmpform In Forms
    If tmpform.Name = formUlario Then
        Set tform = tmpform
    End If
Next
If tform.Enabled = False Then
    tform.Enabled = True
End If
If tform.Visible = False Then
    tform.Visible = True
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
    sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"
        With cnn
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
            .Open
        End With
        rst.Open AEDstr, cnn, adOpenStatic, adLockReadOnly
        Text1.Text = rst.Fields("observaciones") & ""
        cnn.Close
End Sub
