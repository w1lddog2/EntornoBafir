VERSION 5.00
Begin VB.Form frmbackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicacion de Backup"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hacer!!!!"
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Area"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.OptionButton Option2 
         Caption         =   "Ingenieria"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Laboratorio"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Por ejemplo: C:\Backup.rar"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre de archivo"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Formato AAAAMMDDHHMMSS"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Ultimo Backup"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text1.Enabled = True
Else
    Text1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset



If Option1.Value = True Then
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset

    sPathBase = "\\Servidor2\e\entornobafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    
    rst.Open "SELECT fechalab FROM backupbafir", cnn, adOpenStatic, adLockOptimistic
    
       fechaa = Format(Date, "YYYYMMDD")
    horaa = Format(Time, "hhmmss")
    fechaconvertida = fechaa & horaa
    
    rst.Fields("fechalab") = fechaconvertida
    rst.Update
    rst.Close
    rst.Open "SELECT listalab FROM backupbafirlistalab", cnn, adOpenStatic, adLockReadOnly
    
    Do Until rst.EOF = True
        archivos = archivos & " " & rst.Fields("listalab")
        rst.MoveNext
    Loop
    
    
    
    
    
    
    
    
    cnn.Close
Else
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset

    sPathBase = "\\Servidor2\e\entornobafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    
    rst.Open "SELECT fechaing FROM backupbafir", cnn, adOpenStatic, adLockOptimistic
       fechaa = Format(Date, "YYYYMMDD")
    horaa = Format(Time, "hhmmss")
    fechaconvertida = fechaa & horaa
    
    rst.Fields("fechaing") = fechaconvertida
    rst.Update
    rst.Close
    rst.Open "SELECT listaing FROM backupbafirlistaing", cnn, adOpenStatic, adLockReadOnly
    
    Do Until rst.EOF = True
        archivos = archivos & " " & rst.Fields("listaing")
        rst.MoveNext
    Loop
    
     
    
    cnn.Close
End If
    
    destino = Text2.Text
    If Check1.Value = 1 Then
        If Option1.Value = True Then
            comando = "\\Servidor2\E\EntornoBafir\BackUp\rar a -x -m5 -r -ilog\\Servidor2\E\Backups_del_area_industrial\" & fechaconvertida & "_Backuplaboratorio.log -t -ta" & Text1.Text & "-tk -v640000 -y "
        Else
            comando = "\\Servidor2\E\EntornoBafir\BackUp\rar a -x -m5 -r -ilog\\Servidor2\E\Backups_del_area_industrial\" & fechaconvertida & "_Backupingenieria.log -t -ta" & Text1.Text & "-tk -v640000 -y "
        End If
    Else
        If Option1.Value = True Then
            comando = "\\Servidor2\E\EntornoBafir\BackUp\rar a -x -m5 -r -ilog\\Servidor2\E\Backups_del_area_industrial\" & fechaconvertida & "_Backuplaboratorio.log -t -tk -v640000 -y "
        Else
            comando = "\\Servidor2\E\EntornoBafir\BackUp\rar a -x -m5 -r -ilog\\Servidor2\E\Backups_del_area_industrial\" & fechaconvertida & "_BackupIngenieria.log -t -tk -v640000 -y "
        End If
    End If
    Shell (comando & destino & " " & archivos), vbNormalFocus
    
    
    ReDim destinatarios(1 To 5)
    indicedestinatarios = 5
    asunto = "Entorno Bafir: Actualización de Backup"
    mail = ": Se ha realizado el backup correspondiente a la fecha " & Date & ". Los siguientes son archivos que por alguna razón no han sido accesados para su resguardo. Por favor, checkee cual de ellos pueden ser críticos para su función, ya que hasta tanto no esté terminado el algoritmo de backup, los mismos no serán resguardados"
    
    frmbackup.Caption = "Aplicación de Backup - Trabajando. Espere Por favor"
    Form2.Show
    Form2.Hide
    Form2.Show
    Form2.Hide
    Form2.Show
    Form2.Hide
    Form2.Show
    Form2.Hide
    Form2.Show
    Form2.Hide
    Form2.Show
    Form2.Hide
    Form2.Show
    Form2.Hide
    Form2.Show
    Form2.Hide
    
    Do
    
    Loop While IsTaskRunning("\\servidor2\e\EntornoBafir\BackUp\rar.exe") = True
    
    'Call EndTask("\\servidor2\e\EntornoBafir\BackUp\rar.exe")
    Open "\\Servidor2\E\Backups_del_area_industrial\" & fechaconvertida & "_Backuplaboratorio.log" For Input As #1
    Dim Linea As String, total As String
    Do Until EOF(1)
    Line Input #1, Linea
    total = total + Linea + vbCrLf
    Loop
    Close #1

    mail = mail & " " & total
    
    
    
    destinatarios(1) = "pablopirri@bafir.com.ar"
    destinatarios(2) = "gpaludi@bafir.com.ar"
    destinatarios(3) = "calvarez@bafir.com.ar"
    destinatarios(4) = "mmonzani@bafir.com.ar"
    destinatarios(5) = "ingenieria@bafir.com.ar"
    
    
    
    frmSendinfo.Show
    Call Moduloenvio
    frmSendinfo.Hide
    frmbackup.Caption = "Aplicación de Backup"
    Unload Form2
    
    
    
    
    
    
    
    
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Option1_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    
    rst.Open "SELECT fechalab FROM backupbafir", cnn, adOpenStatic, adLockReadOnly
Text1.Text = rst.Fields("fechalab")
cnn.Close
fechat = Format(Date, "YYYYMMDD")
horat = Format(Time, "hhmmss")
fechaconv = fechat & horat
Text2.Text = ("\\Servidor2\E\Backups_del_area_industrial\" & fechaconv & "_BackupLaboratorio.rar")
If Text1.Text = "" Then
    Text1.Enabled = False
    Check1.Value = 0
End If
End Sub

Private Sub Option2_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    
    rst.Open "SELECT fechaing FROM backupbafir", cnn, adOpenStatic, adLockReadOnly
Text1.Text = rst.Fields("fechaing")
cnn.Close

fechat = Format(Date, "YYYYMMDD")
horat = Format(Time, "hhmmss")
fechaconv = fechat & horat
Text2.Text = ("\\Servidor2\E\Backups_del_area_industrial\" & fechaconv & "_BackupIngenieria.rar")
If Text1.Text = "" Then
    Text1.Enabled = False
    Check1.Value = 0
End If
End Sub
