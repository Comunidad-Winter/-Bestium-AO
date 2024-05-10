VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmLauncher 
   BorderStyle     =   0  'None
   Caption         =   "Tierras Perdidas"
   ClientHeight    =   8520
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7140
   ControlBox      =   0   'False
   Icon            =   "frmLauncher.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   10610
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image Configurar 
      Height          =   690
      Left            =   3620
      Top             =   7680
      Width           =   1560
   End
   Begin VB.Image Manual 
      Height          =   690
      Left            =   1920
      Top             =   7680
      Width           =   1560
   End
   Begin VB.Image Foro 
      Height          =   690
      Left            =   240
      Top             =   7680
      Width           =   1560
   End
   Begin VB.Image Actualizar 
      Height          =   420
      Left            =   5520
      Top             =   6360
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "¡Se encontraron actualizaciones disponibles!"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6440
      Width           =   5055
   End
   Begin VB.Image Salir 
      Height          =   690
      Left            =   5340
      Top             =   7680
      Width           =   1560
   End
   Begin VB.Image Jugar 
      Height          =   615
      Left            =   240
      Top             =   6840
      Width           =   6630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   840
      TabIndex        =   0
      Top             =   4200
      Width           =   2655
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************configurar
Private Sub Configurar_Click()
frmConfiguracion.Show
End Sub

Private Sub configurar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Configurar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_config_a.jpg")
End Sub

Private Sub configurar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Configurar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_config_I.jpg")
End Sub
'************configurar
'*************jugar
Private Sub Jugar_Click()
Call Main
Unload Me
End Sub

Private Sub jugar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Jugar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_Jugar_A.jpg")
End Sub

Private Sub jugar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Jugar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_Jugar_I.jpg")
End Sub
'************jugar
'*************salir
Private Sub Salir_Click()
End
End Sub

Private Sub Salir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Salir.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_salir_A.jpg")
End Sub

Private Sub salir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Salir.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_salir_I.jpg")
End Sub
'************salir
'*************foro
Private Sub foro_Click()
Call ShellExecute(0, "Open", "http://benja-studios.foroargentina.net/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub foro_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Foro.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_foro_A.jpg")
End Sub

Private Sub foro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Foro.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_foro_I.jpg")
End Sub
'************foro
'*************manual
Private Sub manual_Click()
Call ShellExecute(0, "Open", "http://www.tierras-perdidas.com/ao/manual.php?id=0", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub manual_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Manual.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_manual_A.jpg")
End Sub

Private Sub manual_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Manual.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_manual_I.jpg")
End Sub
'************manual
'*************actualizar
Private Sub actualizar_Click()
Shell (DirInterfaces & "Launcher\au.exe")
End Sub

Private Sub actualizar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Actualizar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_main_actualizar_A.jpg")
End Sub

Private Sub actualizar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Actualizar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_main_actualizar_I.jpg")
End Sub
'************actualizar

Private Sub Form_Load()
Me.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main.jpg")
Jugar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_Jugar_N.jpg")
Configurar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_config_N.jpg")
Foro.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_foro_N.jpg")
Salir.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_salir_N.jpg")
Manual.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_manual_N.jpg")
Actualizar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_actualizar_N.jpg")
Call WebBrowser1.Navigate("http://benja-studios.foroargentina.net/h1-noticias")
 
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If
Winsock1.Connect "127.0.0.1", "7666"
End Sub

Private Sub Winsock1_Connect()
Label1.ForeColor = vbGreen
Label1.Caption = "Online"
Winsock1.Close
End Sub
 
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label1.ForeColor = vbRed
Label1.Caption = "Offline"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Jugar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_Jugar_N.jpg")
Configurar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_config_N.jpg")
Foro.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_foro_N.jpg")
Salir.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_salir_N.jpg")
Manual.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_manual_N.jpg")
Actualizar.Picture = LoadPicture(DirInterfaces & "Launcher\Launcher_Main_actualizar_N.jpg")
End Sub
