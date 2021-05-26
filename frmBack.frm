VERSION 5.00
Begin VB.Form frmBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Back Up del Sistema"
   ClientHeight    =   6165
   ClientLeft      =   1665
   ClientTop       =   330
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   11865
   Begin VB.CommandButton Command9 
      Caption         =   "Seleccionar Destino"
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Agregar directorio"
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   4080
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Salir"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Realizar Back Up!!"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stand by"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Programar Fechas y hora"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar o quitar archivos a la lista"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label7 
      Height          =   735
      Left            =   6600
      TabIndex        =   20
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Listado:"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Tamaño"
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Archivos"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "frmBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private sAppName As String, sAppPath As String
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Function ShortFileName(ByVal long_name As String) As String
Dim length As Long
Dim short_name As String

    short_name = Space$(1024)
    length = GetShortPathName(long_name, short_name, Len(short_name))
    ShortFileName = Left$(short_name, length)
End Function



'Public Function Shell(ByVal Pathname As String, Optional ByVal Wait As Boolean = False, Optional ByVal Timeout As Integer = -1) As Integer

'End Function



Public Function Hacer_Backup()
tiempoinicial = Time()
tiempoinicial = ((Hour(tiempoinicial)) * 60) + Minute(tiempoinicial) + ((Second(tiempoinicial)) / 60)






Dim db As Database
Dim rs As Recordset
Dim rs2 As Recordset
Dim fso As FileSystemObject
Dim foldersys As Folder
On Error Resume Next
FileCopy "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", Label7.Caption & "\" & Format(Date, "yymmdd") & "-Partidas de compuesto.mdb"
If Err.Number <> 0 Then
fechabase = "Fecha: " & Date
horabase = "Hora: " & Time
errobase = "Error: " & Err.Number
Descbase = "Descripción: " & Err.Description
archivobase = "Archivo: \\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
End If

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs2 = db.OpenRecordset("Select fecha, hora, Error, descripcion, archivo from back_err")

'aca va el codigo de realizar backup
    Set rs = db.OpenRecordset("Select ruta, arch from " & Label6.Caption)
    destino = Label7.Caption & "\"
    
    Form2.Show
    'asdasdasdasd = MsgBox("Puede que el programa no responda por un lapso de tiempo prolongado. Por favor, antes de continuar cierre todas las aplicaciones.", vbInformation + vbOKOnly, "Advertencia")
    frmBack.Enabled = False
    frmBack.Visible = False
    Form2.Visible = True
    Form2.Enabled = True
    Form2.Label1.Visible = True
    Form2.ProgressBar1.Visible = True
    rs.MoveLast
    a = rs.RecordCount
    rs.MoveFirst
    Form2.ProgressBar1.Value = 0
    
    Do
    Form2.ProgressBar1.Value = rs.AbsolutePosition * 100 / a
    'a = destino & Format(Date, "yymmdd") & rs.Fields("arch")
    On Error Resume Next
    
    'FileCopy rs.Fields("ruta"), destino & Format(Date, "yymmdd") & "-" & rs.Fields("arch")
    
    If rs.Fields("arch") = "FolderCarpeta" Then
    ac = 1
    b = 1
    Do Until b = 0
    b = InStr(ac, rs.Fields("ruta"), "\")
    If b <> 0 Then
        C = b
    End If
    ac = b + 1
    Loop
    totalde = Len(rs.Fields("ruta"))
    algo = totalde - C
    folderva = Right(rs.Fields("ruta"), algo)
    
    
    Set fso = New FileSystemObject
    fso.CopyFolder rs.Fields("ruta"), destino & folderva
    Else
    FileCopy rs.Fields("ruta"), destino & rs.Fields("arch")
    End If
    
    
    If Err.Number <> 0 Then
        rs2.AddNew
        rs2.Fields("fecha") = Date
        rs2.Fields("hora") = Time
        rs2.Fields("error") = Err.Number
        rs2.Fields("descripcion") = Err.Description
        rs2.Fields("archivo") = rs.Fields("ruta")
        rs2.Update
    End If
    
    
    
    rs.MoveNext
    Loop Until rs.EOF = True
    
    
    Form2.ProgressBar1.Visible = False
    Form2.Hide
    

'Dim ActiveXZip As New ActiveXZip
'ActiveXZip.Create destino & Format(Date, "yymmdd") & ".zip"
'ActiveXZip.addFile fileSpec:=destino & Format(Date, "yymmdd") & "*.*", recursive:=True, storePaths:=True, Password:=""
'ActiveXZip.Close
    
    ''''''aca hay que editar el back.bat
    
    'ShellExecute 0&, vbNullString, "\\Servidor2\e\EntornoBafir\Documentos\Manual Entorno Bafir.PDF", vbNullString, vbNullString, vbMaximizedFocus
    
    orig = ShortFileName(Label7.Caption)
    
    
    Dim intFileHandle As Integer
    intFileHandle = FreeFile
    Open "\\Servidor2\e\EntornoBafir\BackUp\back.bat" For Output As #intFileHandle
    Print #intFileHandle, "cd\"
    Print #intFileHandle, "f:"
    Print #intFileHandle, "cd\"
    Print #intFileHandle, "cd entorn~1"
    Print #intFileHandle, "cd backup"
    Print #intFileHandle, "rar  a -r -agYYMMDD -ap -ep2 -m5 -v600m " & orig & "\" & " " & orig & "\*.*"
    Close #intFileHandle
    
    
    
    
    sAppName = "Finalizado - back"
    sAppPath = "f:\entorn~1\backup\back.bat"
    
    Shell sAppPath, vbNormalFocus
    
    
    Do Until IsTaskRunning(sAppName) = True
    
    Loop
    
    Call EndTask(sAppName)
       
    
    frmBack.Enabled = True
    
    

    rs2.MoveFirst
    
    
    intFileHandle = FreeFile
    Open destino & Format(Date, "yymmdd") & "-Error log.txt" For Output As #intFileHandle
    rs2.MoveFirst
    Print #intFileHandle, "**********************************************************"
    Print #intFileHandle, "**********************************************************"
    Print #intFileHandle, "                  Entorno Bafir Error Log"
    Print #intFileHandle, "**********************************************************"
    Print #intFileHandle, "**********************************************************"
    
    If rs2.RecordCount = 0 Then
    Print #intFileHandle, Date, Time, "Back Up Sin errores"
    Close #intFileHandle
    Else
        Do Until rs2.EOF = True
            fecha = "Fecha: " & rs2.Fields("Fecha")
            hora = "Hora: " & rs2.Fields("Hora")
            erro = "Error: " & rs2.Fields("Error")
            Desc = "Descripción: " & rs2.Fields("Descripcion")
            arcHivo = "Archivo: " & rs2.Fields("archivo")
            Print #intFileHandle, fecha, hora, erro, Desc, arcHivo
            rs2.MoveNext
        Loop
            Print #intFileHandle, fechabase, horabase, errobase, Descbase, archivobase
        Close #intFileHandle
    rs2.MoveFirst
    Do
    rs2.MoveFirst
    If rs2.EOF = True Then
    Exit Do
    End If
    rs2.Delete
    rs2.MoveNext
    Loop
    
    End If

    ShellExecute 0&, vbNullString, destino & Format(Date, "yymmdd") & "-Error log.txt", vbNullString, vbNullString, vbMaximizedFocus
    rs.MoveFirst
    Do Until rs.EOF = True
    If rs.Fields("arch") = "FolderCarpeta" Then
    
    a = 1
    b = 1
    Do Until b = 0
    b = InStr(a, rs.Fields("ruta"), "\")
    If b <> 0 Then
        C = b
    End If
    a = b + 1
    Loop
    totalde = Len(rs.Fields("ruta"))
    algo = totalde - C
    folderva = Right(rs.Fields("ruta"), algo)
    
    fso.DeleteFolder (destino & folderva), True
    
    
    Else
    Kill destino & rs.Fields("arch")
    End If
    rs.MoveNext
    Loop
    
    
    db.Close
    
    'Kill "\\Servidor2\e\EntornoBafir\backup\" & Format(Date, "yymmdd") & "-" & "partidas de compuesto.mdb"
    asdadssdasd = Format(Date, "yymmdd") & "-Partidas de compuesto.mdb"
    Kill Label7.Caption & "\" & Format(Date, "yymmdd") & "-Partidas de compuesto.mdb"
    frmBack.Visible = True
    frmBack.Enabled = True
    frmBack.SetFocus
    tiempofinal = Time()
    tiempofinal = ((Hour(tiempofinal)) * 60) + Minute(tiempofinal) + ((Second(tiempofinal)) / 60)
    dfg = MsgBox("Se ha terminado de comprimir los archivos incluidos en el backup. Se ha tardado " & tiempofinal - tiempoinicial & " minutos en la realización", vbInformation + vbOKOnly, "Proceso Terminado")
    
    
End Function


Private Sub Command1_Click()
frmPassword.Show (1)
sdf = frmPassword.Password
'sdf = InputBox("Ingrese la contraseña", "Contraseña")
If sdf = vbCancel Then
Exit Sub
Else
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs1 = db.OpenRecordset("Select listados, propietario from backlistados where listados = '" & Label6.Caption & "'")
Set rs = db.OpenRecordset("Select Funcion, Dato from Reg Where Funcion = '" & rs1.Fields("propietario") & "'")

    If sdf <> rs.Fields("dato") Then
        fd = MsgBox("Contraseña Incorrecta", vbCritical + vbOKOnly, "Contraseña Incorrecta")
        Exit Sub
    End If






sdfsdf = MsgBox("Selecciones los archivos y modifique la ruta para verlo a través de Entorno de Red. Si los ingresa a través de rutas tales como 'F:\Producción' es posible que no se realice el back up del archivo. Aseguresé de tener compartida la carpeta en caso de ser archivos locales.", vbCritical + vbOKOnly, "ATENCION!!!")
frmBack.Height = 6540
End If
End Sub

Private Sub Command2_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Fecha, Hora from BackUpFecha Order by Fecha, Hora")
frmCalendar.List1.Clear
If rs.RecordCount <> 0 Then
Do
If CDate(rs.Fields("fecha")) < Date Then
    rs.Delete
Else
frmCalendar.List1.AddItem (rs.Fields("Fecha"))
End If
rs.MoveNext
Loop Until rs.EOF = True
End If
db.Close
frmBack.Enabled = False
frmBack.Visible = False
frmCalendar.Show
End Sub

Private Sub Command3_Click()
If Label7.Caption = "" Then
    jg = MsgBox("Debe seleccionar un destino para el Backup", vbCritical + vbOKOnly, "Error")
    Command9.SetFocus
    Exit Sub
End If
frmPassword.Show (1)
sdf = frmPassword.Password
'sdf = InputBox("Ingrese la contraseña", "Contraseña")
If sdf = vbCancel Then
Exit Sub
Else
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs1 = db.OpenRecordset("Select listados, propietario from backlistados where listados = '" & Label6.Caption & "'")
Set rs = db.OpenRecordset("Select Funcion, Dato from Reg Where Funcion = '" & rs1.Fields("propietario") & "'")

    If sdf <> rs.Fields("dato") Then
        fd = MsgBox("Contraseña Incorrecta", vbCritical + vbOKOnly, "Contraseña Incorrecta")
        Exit Sub
    End If






frmStand.Show
frmBack.Enabled = False
frmBack.Visible = False
End If
End Sub

Private Sub Command4_Click()
If Label7.Caption = "" Then
    jg = MsgBox("Debe seleccionar un destino para el Backup", vbCritical + vbOKOnly, "Error")
    Command9.SetFocus
    Exit Sub
End If
frmPassword.Show (1)
sdf = frmPassword.Password
'sdf = InputBox("Ingrese la contraseña", "Contraseña")
If sdf = vbCancel Then
Exit Sub
Else
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs1 = db.OpenRecordset("Select listados, propietario from backlistados where listados = '" & Label6.Caption & "'")

Set rs = db.OpenRecordset("Select Funcion, Dato from Reg Where Funcion = '" & rs1.Fields("propietario") & "'")




    If sdf <> rs.Fields("dato") Then
        fd = MsgBox("Contraseña Incorrecta", vbCritical + vbOKOnly, "Contraseña Incorrecta")
        db.Close
        Exit Sub
    End If
db.Close
Hacer_Backup
End If

End Sub

Private Sub Command5_Click()
Form1.Enabled = True
Form1.Visible = True
frmBack.Hide
End Sub

Sub agregar_Archivo()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select ruta,arch,size from " & Label6.Caption)

rs.AddNew
rs.Fields("ruta") = Text1.Text
rs.Fields("Arch") = File1.FileName
On Error Resume Next
tamaño = Format(FileLen(Text1.Text), "#,###")
rs.Fields("size") = tamaño
rs.Update
db.Close
List1.AddItem (Text1.Text)
Label2.Caption = CInt(Label2.Caption) + 1
Label4.Caption = CDbl(Label4.Caption) + CDbl(tamaño)
End Sub

Private Sub Command6_Click()
agregar_Archivo
End Sub

Private Sub Command7_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select ruta, size from " & Label6.Caption & " where ruta = '" & List1.Text & "'")

Label4.Caption = CDbl(Label4.Caption) - CDbl(rs.Fields("size"))
rs.Delete
db.Close

List1.RemoveItem List1.ListIndex
Label2.Caption = CInt(Label2.Caption) - 1
End Sub

Private Sub Command8_Click()
Dim db As Database
Dim rs As Recordset
Dim fso As FileSystemObject

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select ruta,arch,size from " & Label6.Caption)

Set fso = New FileSystemObject

rs.AddNew
rs.Fields("ruta") = Dir1.path
rs.Fields("Arch") = "FolderCarpeta"
rs.Fields("size") = TamañoCarpeta(rs.Fields("ruta"))
Label2.Caption = CInt(Label2.Caption) + 1
Label4.Caption = CDbl(Label4.Caption) + TamañoCarpeta(rs.Fields("ruta"))
rs.Update
db.Close
List1.AddItem (Dir1.path)
End Sub

Private Sub Command9_Click()
frmBack.Enabled = False
frmBackSelDir.Show
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub File1_Click()
Text1.Text = Dir1.path & "\" & File1.FileName
End Sub

Private Sub File1_DblClick()
agregar_Archivo
End Sub

Function TamañoCarpeta(strCarpeta)
Dim fso, Carpeta
Set fso = CreateObject("Scripting.FileSystemObject")
Set Carpeta = fso.GetFolder(strCarpeta)
TamañoCarpeta = Format(Carpeta.Size, "#,###")
End Function ' TamañoCarpeta

