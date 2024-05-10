VERSION 5.00
Begin VB.Form frmSubastar 
   BorderStyle     =   0  'None
   Caption         =   "Subasta"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StartBid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Text            =   "1000"
      Top             =   5000
      Width           =   1410
   End
   Begin VB.TextBox Amount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      TabIndex        =   1
      Text            =   "1"
      Top             =   4520
      Width           =   1095
   End
   Begin VB.ListBox ItemList 
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   3960
      ItemData        =   "frmSubastar.frx":0000
      Left            =   240
      List            =   "frmSubastar.frx":0040
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image cmdCancelar 
      Height          =   825
      Left            =   170
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Image cmdSubastar 
      Height          =   840
      Left            =   1200
      Top             =   5400
      Width           =   1920
   End
End
Attribute VB_Name = "frmSubastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSubastar_Click()
SendData ("/TUBASTAR " & frmSubastar.ItemList.ListIndex + 1 & "@" & Amount & "@" & StartBid)
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub form_load()
Me.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Main.jpg")
cmdSubastar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Iniciar_N.jpg")
cmdCancelar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Cancelar_N.jpg")
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
End Sub

Private Sub cmdSubastar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSubastar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Iniciar_A.jpg")
End Sub

Private Sub cmdSubastar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSubastar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Iniciar_I.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSubastar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Iniciar_n.jpg")
cmdCancelar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Cancelar_n.jpg")
End Sub

Private Sub cmdCancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancelar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Cancelar_A.jpg")
End Sub

Private Sub cmdCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancelar.Picture = LoadPicture(DirInterfaces & "Principal\Subasta_Cancelar_I.jpg")
End Sub
