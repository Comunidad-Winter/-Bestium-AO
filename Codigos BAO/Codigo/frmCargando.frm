VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   803
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      Picture         =   "frmCargando.frx":6E0BE
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      Begin VB.Image Barra 
         Height          =   615
         Left            =   2160
         Picture         =   "frmCargando.frx":DC17C
         Top             =   8040
         Width           =   7620
      End
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   2400
      Left            =   2340
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   2640
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   4233
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCargando.frx":E0181
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
LOGO.Picture = LoadPicture(App.Path & "\Graficos\Cargando1.jpg")
Barra.Width = 0
End Sub

