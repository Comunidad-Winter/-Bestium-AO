VERSION 5.00
Begin VB.Form frmMacro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de macros"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Guardar 
      Caption         =   "Grabar Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tecla Numerica / Comando"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox F10Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox F9Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox F8Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox F7Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox F3Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox F4Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox F5Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox F6Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox F2Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox F1Macro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Text            =   "/SANAR"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "1"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "2"
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "3"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "4"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "5"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "6"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "7"
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
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "8"
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
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "9"
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
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "0"
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
         Left            =   120
         TabIndex        =   1
         Top             =   4680
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

 
Private Sub Guardar_Click()
SaveSetting App.EXEName, "textos", "F1Macro", F1Macro.Text
SaveSetting App.EXEName, "textos", "F2Macro", F2Macro.Text
SaveSetting App.EXEName, "textos", "F3Macro", F3Macro.Text
SaveSetting App.EXEName, "textos", "F4Macro", F4Macro.Text
SaveSetting App.EXEName, "textos", "F5Macro", F5Macro.Text
SaveSetting App.EXEName, "textos", "F6Macro", F6Macro.Text
SaveSetting App.EXEName, "textos", "F7Macro", F7Macro.Text
SaveSetting App.EXEName, "textos", "F7Macro", F8Macro.Text
SaveSetting App.EXEName, "textos", "F7Macro", F9Macro.Text
SaveSetting App.EXEName, "textos", "F7Macro", F10Macro.Text
 
frmMensaje.Show
frmMensaje.lblmensaje.Caption = "Macros Guardados."
Me.Hide
End Sub
 
Private Sub form_load()
F1Macro.Text = GetSetting(App.EXEName, "textos", "F1Macro", "")
F2Macro.Text = GetSetting(App.EXEName, "textos", "F2Macro", "")
F3Macro.Text = GetSetting(App.EXEName, "textos", "F3Macro", "")
F4Macro.Text = GetSetting(App.EXEName, "textos", "F4Macro", "")
F5Macro.Text = GetSetting(App.EXEName, "textos", "F5Macro", "")
F6Macro.Text = GetSetting(App.EXEName, "textos", "F6Macro", "")
F7Macro.Text = GetSetting(App.EXEName, "textos", "F7Macro", "")
F7Macro.Text = GetSetting(App.EXEName, "textos", "F8Macro", "")
F7Macro.Text = GetSetting(App.EXEName, "textos", "F9Macro", "")
F7Macro.Text = GetSetting(App.EXEName, "textos", "F10Macro", "")

If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))

End Sub
 
 
