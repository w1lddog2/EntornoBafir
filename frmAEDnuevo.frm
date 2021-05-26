VERSION 5.00
Begin VB.Form frmAEDnuevo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ensayo de AED nº"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command23 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   9480
      TabIndex        =   121
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   9480
      TabIndex        =   120
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   9480
      TabIndex        =   119
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   9480
      TabIndex        =   118
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   7680
      TabIndex        =   117
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   9480
      TabIndex        =   116
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   5280
      TabIndex        =   113
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Volver"
      Height          =   495
      Left            =   9240
      TabIndex        =   112
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text28 
      Enabled         =   0   'False
      Height          =   405
      Left            =   9480
      TabIndex        =   111
      Text            =   "Text1"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   7680
      TabIndex        =   109
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text27 
      Enabled         =   0   'False
      Height          =   405
      Left            =   7680
      TabIndex        =   108
      Text            =   "Text1"
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox Text26 
      Height          =   405
      Left            =   5520
      TabIndex        =   106
      Text            =   "Text1"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   3960
      TabIndex        =   104
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text25 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      TabIndex        =   103
      Text            =   "Text1"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   2040
      TabIndex        =   101
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text24 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2040
      TabIndex        =   100
      Text            =   "Text1"
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox Text23 
      Enabled         =   0   'False
      Height          =   405
      Left            =   9480
      TabIndex        =   97
      Text            =   "Text1"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   7680
      TabIndex        =   95
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text22 
      Enabled         =   0   'False
      Height          =   405
      Left            =   7680
      TabIndex        =   94
      Text            =   "Text1"
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox Text21 
      Height          =   405
      Left            =   5520
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   3960
      TabIndex        =   90
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      TabIndex        =   89
      Text            =   "Text1"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   2040
      TabIndex        =   87
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text19 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2040
      TabIndex        =   86
      Text            =   "Text1"
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      Height          =   405
      Left            =   9480
      TabIndex        =   83
      Text            =   "Text1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   7680
      TabIndex        =   81
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text17 
      Enabled         =   0   'False
      Height          =   405
      Left            =   7680
      TabIndex        =   80
      Text            =   "Text1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   405
      Left            =   5520
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   3960
      TabIndex        =   76
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   2040
      TabIndex        =   73
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2040
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   405
      Left            =   9480
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   7680
      TabIndex        =   67
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   405
      Left            =   7680
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   5520
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   3960
      TabIndex        =   62
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   2040
      TabIndex        =   59
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2040
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   405
      Left            =   9480
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   5520
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   2040
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   405
      Left            =   9480
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   405
      Left            =   7680
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   5520
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   3960
      TabIndex        =   43
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ingresar"
      Height          =   315
      Left            =   2040
      TabIndex        =   40
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2040
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar datos"
      Height          =   495
      Left            =   9240
      TabIndex        =   35
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "Permitido 5 - 15 %"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9120
      TabIndex        =   115
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   2160
      TabIndex        =   114
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   9960
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label71 
      Caption         =   "Dureza"
      Height          =   255
      Left            =   8760
      TabIndex        =   110
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label70 
      Caption         =   "Peso en agua"
      Height          =   255
      Left            =   6600
      TabIndex        =   107
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label69 
      Caption         =   "Peso"
      Height          =   255
      Left            =   5040
      TabIndex        =   105
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label68 
      Caption         =   "Perímetro"
      Height          =   255
      Left            =   3120
      TabIndex        =   102
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label67 
      Caption         =   "Espesor"
      Height          =   255
      Left            =   1200
      TabIndex        =   99
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label66 
      Caption         =   "Oring 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   98
      Top             =   7080
      Width           =   735
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   9960
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label65 
      Caption         =   "Dureza"
      Height          =   255
      Left            =   8760
      TabIndex        =   96
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label64 
      Caption         =   "Peso en agua"
      Height          =   255
      Left            =   6600
      TabIndex        =   93
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label63 
      Caption         =   "Peso"
      Height          =   255
      Left            =   5040
      TabIndex        =   91
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label62 
      Caption         =   "Perímetro"
      Height          =   255
      Left            =   3120
      TabIndex        =   88
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label61 
      Caption         =   "Espesor"
      Height          =   255
      Left            =   1200
      TabIndex        =   85
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label60 
      Caption         =   "Oring 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   84
      Top             =   6120
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   9960
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label59 
      Caption         =   "Dureza"
      Height          =   255
      Left            =   8760
      TabIndex        =   82
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label58 
      Caption         =   "Peso en agua"
      Height          =   255
      Left            =   6600
      TabIndex        =   79
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label57 
      Caption         =   "Peso"
      Height          =   255
      Left            =   5040
      TabIndex        =   77
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label56 
      Caption         =   "Perímetro"
      Height          =   255
      Left            =   3120
      TabIndex        =   74
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label55 
      Caption         =   "Espesor"
      Height          =   255
      Left            =   1200
      TabIndex        =   71
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label54 
      Caption         =   "Oring 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   70
      Top             =   5160
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   9960
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label53 
      Caption         =   "Dureza"
      Height          =   255
      Left            =   8760
      TabIndex        =   68
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label52 
      Caption         =   "Peso en agua"
      Height          =   255
      Left            =   6600
      TabIndex        =   65
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label51 
      Caption         =   "Peso"
      Height          =   255
      Left            =   5040
      TabIndex        =   63
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label50 
      Caption         =   "Perímetro"
      Height          =   255
      Left            =   3120
      TabIndex        =   60
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label49 
      Caption         =   "Espesor"
      Height          =   255
      Left            =   1200
      TabIndex        =   57
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label48 
      Caption         =   "Oring 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   56
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label47 
      Caption         =   "Elongación (mm)"
      Height          =   255
      Left            =   8160
      TabIndex        =   54
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label46 
      Caption         =   "Carga de rotura (Kg)"
      Height          =   255
      Left            =   3960
      TabIndex        =   52
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label45 
      Caption         =   "Modulo al 50% (kg)"
      Height          =   255
      Left            =   600
      TabIndex        =   50
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label44 
      Caption         =   "Dureza"
      Height          =   255
      Left            =   8760
      TabIndex        =   48
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label43 
      Caption         =   "Peso en agua"
      Height          =   255
      Left            =   6600
      TabIndex        =   46
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label42 
      Caption         =   "Peso"
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label41 
      Caption         =   "Perímetro"
      Height          =   255
      Left            =   3120
      TabIndex        =   41
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label40 
      Caption         =   "Espesor"
      Height          =   255
      Left            =   1200
      TabIndex        =   38
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label39 
      Caption         =   "Oring 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label38 
      Caption         =   "Valores de los orings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   9960
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label37 
      Caption         =   "Datos del ensayo nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label36 
      BackColor       =   &H80000009&
      Caption         =   "36"
      Height          =   255
      Left            =   9120
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label35 
      BackColor       =   &H80000009&
      Caption         =   "35"
      Height          =   255
      Left            =   9120
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label34 
      BackColor       =   &H80000009&
      Caption         =   "34"
      Height          =   255
      Left            =   9120
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label33 
      BackColor       =   &H80000009&
      Caption         =   "33"
      Height          =   495
      Left            =   5520
      TabIndex        =   30
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label32 
      BackColor       =   &H80000009&
      Caption         =   "32"
      Height          =   255
      Left            =   5520
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label31 
      BackColor       =   &H80000009&
      Caption         =   "31"
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label30 
      BackColor       =   &H80000009&
      Caption         =   "30"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label29 
      BackColor       =   &H80000009&
      Caption         =   "29"
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label28 
      BackColor       =   &H80000009&
      Caption         =   "28"
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label27 
      BackColor       =   &H80000009&
      Caption         =   "27"
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackColor       =   &H80000009&
      Caption         =   "26"
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label25 
      BackColor       =   &H80000009&
      Caption         =   "25"
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000009&
      Caption         =   "24"
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label23 
      BackColor       =   &H80000009&
      Caption         =   "23"
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000009&
      Caption         =   "22"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000009&
      Caption         =   "20"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000009&
      Caption         =   "19"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label18 
      Caption         =   "Volumen interno"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "Dimensiones del recipiente de ensayo ID x H (mm)"
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "Compresion del sello"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Dimensiones del alojamiento ID x OD x H (mm)"
      Height          =   735
      Left            =   3480
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "Pieza"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Medio"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Presión"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Temperatura"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Lote"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Titulo"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Material"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Norma"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Especificación"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Doc."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Proyecto"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmAEDnuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public responsablecon As String
Public h

Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset
Dim dime As String


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDconst where cod = 0")

titulo = InputBox("Ingrese título", "Título", rs.Fields("Titulo"))
    If titulo = "" Then
        Exit Sub
    End If
proyecto = InputBox("Ingrese proyecto", "Proyecto", rs.Fields("Proyecto"))
    If proyecto = "" Then
        Exit Sub
    End If
doc = InputBox("Ingrese Documento", "Documento", rs.Fields("Doc"))
    If doc = "" Then
        Exit Sub
    End If
especificacion = InputBox("Ingrese especificación", "Especificación", rs.Fields("Especificacion"))
    If especificacion = "" Then
        Exit Sub
    End If
Norma = InputBox("Ingrese norma", "Norma", rs.Fields("Norma"))
    If Norma = "" Then
        Exit Sub
    End If
material = InputBox("Ingrese material", "Material", rs.Fields("Material"))
    If material = "" Then
        Exit Sub
    End If
Codigo = InputBox("Ingrese código", "Código", rs.Fields("Codigo"))
    If Codigo = "" Then
        Exit Sub
    End If
lote = InputBox("Ingrese lote", "Lote", rs.Fields("Lote"))
    If lote = "" Then
        Exit Sub
    End If
temperatura = InputBox("Ingrese temperatura", "Temperatura", rs.Fields("Temperatura"))
    If temperatura = "" Then
        Exit Sub
    End If
presion = InputBox("Ingrese presión", "Presión", rs.Fields("Presion"))
    If presion = "" Then
        Exit Sub
    End If
medio = InputBox("Ingrese Medio", "Medio", rs.Fields("Medio"))
    If medio = "" Then
        Exit Sub
    End If

ciclos = InputBox("Ingrese Cantidad de Ciclos", "ciclos", rs.Fields("ciclos"))
    If ciclos = "" Then
        Exit Sub
    End If

tiempodeciclo = InputBox("Ingrese Tiempo de ciclo", "Tiempo de ciclo", rs.Fields("tiempodeciclo"))
    If tiempodeciclo = "" Then
        Exit Sub
    End If

tiempodedescompresion = InputBox("Ingrese tiempo de descompresión", "Tiempo de descompresión", rs.Fields("tiempodedescompresion"))
    If tiempodedescompresion = "" Then
        Exit Sub
    End If


instrumento = InputBox("Ingrese Cliente", "Cliente", rs.Fields("Instrumento"))
    If instrumento = "" Then
        Exit Sub
    End If
loteinst = InputBox("Ingrese Pieza", "Pieza", rs.Fields("Loteinst"))
    If loteinst = "" Then
        Exit Sub
    End If
dime = InputBox("Ingrese dimensiones del alojamiento IDxODxH (mm)", "Dimensiones", rs.Fields("Alojam"))
    If dime = "" Then
        Exit Sub
    End If
a = InStr(1, (dime), ".")
    If a <> 0 Then
    Mid(dime, a) = ","
    End If
    
arradime = Explode("x", dime)

h = InputBox("La altura de alojamiento es (mm)", "Corrección", arradime(2))
    If h = "" Then
        Exit Sub
    End If

dime1 = InputBox("Ingrese dimensiones del recipiente IDxH (mm)", "Dimensiones", rs.Fields("dimen"))
    If dime1 = "" Then
        Exit Sub
    End If
volumen = InputBox("Ingrese volumen interno", "Volumen interno", rs.Fields("volumen"))
    If volumen = "" Then
        Exit Sub
    End If

frmAEDnuevo.Label19.Caption = titulo
frmAEDnuevo.Label20.Caption = proyecto
frmAEDnuevo.Label22.Caption = doc
frmAEDnuevo.Label23.Caption = especificacion
frmAEDnuevo.Label24.Caption = Norma
frmAEDnuevo.Label25.Caption = material
frmAEDnuevo.Label26.Caption = Codigo
frmAEDnuevo.Label27.Caption = lote

frmAEDnuevo.Label28.Caption = temperatura
frmAEDnuevo.Label29.Caption = presion
frmAEDnuevo.Label30.Caption = medio
frmAEDnuevo.Label31.Caption = instrumento
frmAEDnuevo.Label32.Caption = loteinst

frmAEDnuevo.Label33.Caption = dime
frmAEDnuevo.Label35.Caption = dime1
frmAEDnuevo.Label36.Caption = volumen
    rs.Edit
    rs.Fields("Titulo") = titulo
    rs.Fields("Proyecto") = proyecto
    rs.Fields("Doc") = doc
    rs.Fields("Especificacion") = especificacion
    rs.Fields("Norma") = Norma
    rs.Fields("Material") = material
    rs.Fields("Codigo") = Codigo
    rs.Fields("Lote") = lote
    rs.Fields("Temperatura") = temperatura
    rs.Fields("Presion") = presion
    rs.Fields("Medio") = medio
    rs.Fields("Instrumento") = instrumento
    rs.Fields("Loteinst") = loteinst
    rs.Fields("Alojam") = dime
    rs.Fields("dimen") = dime1
    rs.Fields("volumen") = volumen
    rs.Fields("ciclos") = ciclos
    rs.Fields("tiempodeciclo") = tiempodeciclo
    rs.Fields("tiempodedescompresion") = tiempodedescompresion
    rs.Update
db.Close
End Sub

Private Sub Command10_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text17.Text = este / 3
End Sub

Private Sub Command11_Click()
este = 0
For hacer = 1 To 4
Valor = InputBox("Ingrese el valor " & hacer & " de 4", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text19.Text = este / 4
End Sub

Private Sub Command12_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text20.Text = este / 3
End Sub

Private Sub Command13_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text22.Text = este / 3
End Sub

Private Sub Command14_Click()
este = 0
For hacer = 1 To 4
Valor = InputBox("Ingrese el valor " & hacer & " de 4", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text24.Text = este / 4
End Sub

Private Sub Command15_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text25.Text = este / 3
End Sub

Private Sub Command16_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text27.Text = este / 3
End Sub

Private Sub Command17_Click()
Form1.Enabled = True
Form1.Visible = True
frmAEDnuevo.Hide
End Sub

Private Sub Command18_Click()
If Label19.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label20.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label22.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label23.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label24.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label25.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label26.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label27.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label28.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label29.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label30.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label31.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label32.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label33.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label34.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label35.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Label36.Caption = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

If Text1.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text2.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text3.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text4.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text5.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text6.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text7.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text8.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text9.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text10.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text11.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text12.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text13.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text14.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text15.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text16.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text17.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text18.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text19.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text20.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text21.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text22.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text23.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text24.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text25.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text26.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text27.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If Text28.Text = "" Then
    dfsdf = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If




Dim db As Database
Dim rs As Recordset


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDorig")
rs.AddNew
rs.Fields("ensayo") = Label3.Caption
rs.Fields("espesor1") = Text1.Text
rs.Fields("espesor2") = Text9.Text
rs.Fields("espesor3") = Text14.Text
rs.Fields("espesor4") = Text19.Text
rs.Fields("espesor5") = Text24.Text
rs.Fields("perimetro1") = Text2.Text
rs.Fields("perimetro2") = Text10.Text
rs.Fields("perimetro3") = Text15.Text
rs.Fields("perimetro4") = Text20.Text
rs.Fields("perimetro5") = Text25.Text
rs.Fields("fecha") = Date
Valor = Text3.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("peso1") = Valor
Valor = Text11.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("peso2") = Valor
Valor = Text16.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("peso3") = Valor
Valor = Text21.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("peso4") = Valor
Valor = Text26.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("peso5") = Valor
rs.Fields("pagua1") = Text4.Text
rs.Fields("pagua2") = Text12.Text
rs.Fields("pagua3") = Text17.Text
rs.Fields("pagua4") = Text22.Text
rs.Fields("pagua5") = Text27.Text

Valor = Text5.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("dureza1") = Valor
Valor = Text13.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("dureza2") = Valor
Valor = Text18.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("dureza3") = Valor
Valor = Text23.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("dureza4") = Valor
Valor = Text28.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("dureza5") = Valor
Valor = Text3.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("densidad1") = (Valor * 0.9971) / (Valor - Text4.Text)
Valor = Text11.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("densidad2") = (Valor * 0.9971) / (Valor - Text12.Text)
Valor = Text16.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("densidad3") = (Valor * 0.9971) / (Valor - Text17.Text)
Valor = Text21.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("densidad4") = (Valor * 0.9971) / (Valor - Text22.Text)
Valor = Text26.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("densidad5") = (Valor * 0.9971) / (Valor - Text27.Text)
rs.Fields("diamint1") = (Text2.Text / 3.14) - (2 * Text1.Text)
rs.Fields("diamint2") = (Text10.Text / 3.14) - (2 * Text9.Text)
rs.Fields("diamint3") = (Text15.Text / 3.14) - (2 * Text14.Text)
rs.Fields("diamint4") = (Text20.Text / 3.14) - (2 * Text19.Text)
rs.Fields("diamint5") = (Text25.Text / 3.14) - (2 * Text24.Text)
Valor = Text7.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("mpa1") = (Valor / (Text1.Text * Text1.Text * 1.57)) * 10
Valor = Text8.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
di = (Text2.Text / 3.14) - (2 * Text1.Text)
rs.Fields("elong1") = (Valor * 2 + (13.07 * 3.14) - (di * 3.14)) / (di * 3.14) * 100
a = (Valor * 2 + 12.9 - (Text2.Text - (2 * Text1.Text))) / (Text2.Text - (2 * Text1.Text))
Valor = Text6.Text
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
rs.Fields("modulo501") = (Valor / (Text1.Text * Text1.Text * 1.57)) * 10
rs.Fields("responsable") = responsablecon
rs.Update


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDconst")


rs.AddNew
rs.Fields("cod") = Label3.Caption
rs.Fields("Titulo") = Label19.Caption
rs.Fields("Proyecto") = Label20.Caption
rs.Fields("Doc") = Label22.Caption
rs.Fields("Especificacion") = Label23.Caption
rs.Fields("Norma") = Label24.Caption
rs.Fields("Material") = Label25.Caption
rs.Fields("Codigo") = Label26.Caption
rs.Fields("Lote") = Label27.Caption
rs.Fields("Temperatura") = Label28.Caption
rs.Fields("Presion") = Label29.Caption
rs.Fields("Medio") = Label30.Caption
rs.Fields("Instrumento") = Label31.Caption
rs.Fields("Loteinst") = Label32.Caption
rs.Fields("Alojam") = Label33.Caption
rs.Fields("dimen") = Label35.Caption
rs.Fields("volumen") = Label36.Caption
rs.Fields("comp") = Label34.Caption
rs.Update

dfsdfsf = MsgBox("Los datos se han guardado exitosamente con el nº de ensayo " & Label3.Caption, vbInformation + vbOKOnly, "Datos Guardados")

Form1.Enabled = True
Form1.Visible = True
frmAEDnuevo.Hide
db.Close
End Sub

Private Sub Command19_Click()
dureza = 0
For tomar = 1 To 5
dureza = dureza + CDbl(punto_por_coma(InputBox("Ingrese la dureza", "Ingrese la dureza " & tomar & "/5")))
Next
Text5.Text = dureza / 5
End Sub

Private Sub Command2_Click()
If h = "" Then
    sdfsdf = MsgBox("Debe ingresar primero las propiedades del ensayo", vbCritical + vbOKOnly, "Error")
    Command1.SetFocus
    Exit Sub
End If

este = 0
For hacer = 1 To 4
Valor = InputBox("Ingrese el valor " & hacer & " de 4", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text1.Text = este / 4

compse = ((-(h / Text1.Text)) + 1) * 100
'h es la altura del alojamiento

Label34.Caption = Format(compse, "0.00") & "%"
End Sub

Private Sub Command20_Click()
dureza = 0
For tomar = 1 To 5
dureza = dureza + CDbl(punto_por_coma(InputBox("Ingrese la dureza", "Ingrese la dureza " & tomar & "/5")))
Next
Text13.Text = dureza / 5
End Sub

Private Sub Command21_Click()
dureza = 0
For tomar = 1 To 5
dureza = dureza + CDbl(punto_por_coma(InputBox("Ingrese la dureza", "Ingrese la dureza " & tomar & "/5")))
Next
Text18.Text = dureza / 5
End Sub

Private Sub Command22_Click()
dureza = 0
For tomar = 1 To 5
dureza = dureza + CDbl(punto_por_coma(InputBox("Ingrese la dureza", "Ingrese la dureza " & tomar & "/5")))
Next
Text23.Text = dureza / 5
End Sub

Private Sub Command23_Click()
dureza = 0
For tomar = 1 To 5
dureza = dureza + CDbl(punto_por_coma(InputBox("Ingrese la dureza", "Ingrese la dureza " & tomar & "/5")))
Next
Text28.Text = dureza / 5
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
    sdfsdfdf = MsgBox("Debe completar el espesor primero", vbCritical + vbOKOnly, "Error")
    Command2.SetFocus
    Exit Sub
End If

este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text2.Text = este / 3

'''''''para modulo al 50
di = (Text2.Text / 3.14) - (2 * Text1.Text)

estirar = ((50 / 100) + 1 - (41.04 / (di * 3.14))) * ((di * 3.14) / 2)
' 41.04 es el perimetro de la mordaza; ((Text2.Text - (2 * Text1.Text)) es el perimetro interno del oring

MsgBox ("Para módulo al 50% debe estirar el oring " & estirar & " mm.")

'''''''para modulo al 50


End Sub

Private Sub Command4_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text4.Text = este / 3
End Sub

Private Sub Command5_Click()
este = 0
For hacer = 1 To 4
Valor = InputBox("Ingrese el valor " & hacer & " de 4", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text9.Text = este / 4
End Sub

Private Sub Command6_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text10.Text = este / 3
End Sub

Private Sub Command7_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text12.Text = este / 3
End Sub

Private Sub Command8_Click()
este = 0
For hacer = 1 To 4
Valor = InputBox("Ingrese el valor " & hacer & " de 4", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text14.Text = este / 4
End Sub

Private Sub Command9_Click()
este = 0
For hacer = 1 To 3
Valor = InputBox("Ingrese el valor " & hacer & " de 3", "Ingreso de datos")
If Valor = "" Then
    Exit Sub
End If
a = InStr(1, (Valor), ".")
    If a <> 0 Then
    Mid(Valor, a) = ","
    End If
este = este + CDbl(Valor)
Next
Text15.Text = este / 3
End Sub

