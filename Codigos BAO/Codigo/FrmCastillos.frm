VERSION 5.00
Begin VB.Form FrmCastillos 
   BackColor       =   &H00808080&
   Caption         =   "Viajar a Castillos"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4605
   FillColor       =   &H00008080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "NORTE"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SUR"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000008&
      Caption         =   "ESTE"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OESTE"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde acá podras viajar a los diferentes tipos de Castillos, clickeando en los BOTONES, Pueden viajar: Lider/Sublider/Miembro"
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "FrmCastillos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Oeste?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO OESTE")
End If
End Sub
Private Sub Command2_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Este?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO ESTE")
End If
End Sub

Private Sub Command3_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Sur?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO SUR")
End If
End Sub

Private Sub Command4_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Norte?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO NORTE")
End If
End Sub


