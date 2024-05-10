VERSION 5.00
Begin VB.Form frmTiraItems 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmTiraItem.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1950
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   310
      Left            =   755
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "0"
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image Command7 
      Height          =   390
      Left            =   200
      Top             =   420
      Width           =   495
   End
   Begin VB.Image Command1 
      Height          =   450
      Left            =   240
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Image Command4 
      Height          =   420
      Left            =   2160
      Top             =   840
      Width           =   930
   End
   Begin VB.Image Command5 
      Height          =   420
      Left            =   1200
      Top             =   840
      Width           =   930
   End
   Begin VB.Image Command6 
      Height          =   420
      Left            =   240
      Top             =   840
      Width           =   930
   End
   Begin VB.Image command3 
      Height          =   390
      Left            =   2625
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "frmTiraItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
frmTiraItems.Visible = False
SendData "TI" & Inventario.SelectedItem & "," & frmTiraItems.Text1.Text
frmTiraItems.Text1.Text = "0"
End Sub
Private Sub Command4_Click()
frmTiraItems.Visible = False
If Inventario.SelectedItem <> FLAGORO Then
    SendData "TI" & Inventario.SelectedItem & "," & Inventario.Amount(Inventario.SelectedItem)
Else
    SendData "TI" & Inventario.SelectedItem & "," & UserGLD
End If

frmTiraItems.Text1.Text = "0"
End Sub

Private Sub Command7_Click()
If Text1.Text = "" Then
Text1.Text = 0
Exit Sub
End If
Text1.Text = Text1.Text - 1
End Sub

Private Sub Form_Deactivate()
'Unload Me
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
Text1.Text = 1
Exit Sub
End If
Text1.Text = Text1.Text + 1
End Sub

Private Sub Command6_Click()
If Text1.Text = "" Then
Text1.Text = 100
Exit Sub
End If
Text1.Text = Text1.Text + 1000
End Sub

Private Sub Command5_Click()
If Text1.Text = "" Then
Text1.Text = 1000
Exit Sub
End If
Text1.Text = Text1.Text + 1000
End Sub

Private Sub Form_Load()
    If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
Me.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_Main.jpg")
Command7.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMenosN.jpg")
Command1.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BaceptarN.jpg")
command3.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMasN.jpg")
Command4.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BtodoN.jpg")
Command5.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+1000N.jpg")
Command6.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+100N.jpg")
End Sub

Private Sub text1_Change()
On Error GoTo ErrHandler
    If Val(Text1.Text) < 0 Then
        Text1.Text = MAX_INVENTORY_OBJS
    End If
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        If Inventario.SelectedItem <> FLAGORO Or Val(Text1.Text) > UserGLD Then
            Text1.Text = "1"
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub


Private Sub command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMenosa.jpg")
End Sub

Private Sub command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMenosi.jpg")
End Sub

Private Sub command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_Baceptara.jpg")
End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_Baceptari.jpg")
End Sub

Private Sub command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
command3.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMasa.jpg")
End Sub

Private Sub command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
command3.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMasi.jpg")
End Sub

Private Sub command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_Btodoa.jpg")
End Sub

Private Sub command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_Btodoi.jpg")
End Sub

Private Sub command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+1000a.jpg")
End Sub

Private Sub command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+1000i.jpg")
End Sub

Private Sub command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+100a.jpg")
End Sub

Private Sub command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+100i.jpg")
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMenosN.jpg")
Command1.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BaceptarN.jpg")
command3.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BMasN.jpg")
Command4.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_BtodoN.jpg")
Command5.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+1000N.jpg")
Command6.Picture = LoadPicture(DirInterfaces & "Principal\TirarObj_B+100N.jpg")
End Sub


