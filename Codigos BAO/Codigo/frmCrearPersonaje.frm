VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Crear Personaje"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   591
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cabeza 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   8760
      TabIndex        =   17
      Top             =   4800
      Width           =   2220
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5160
      TabIndex        =   16
      Top             =   480
      Width           =   3660
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":8AE86
      Left            =   9480
      List            =   "frmCrearPersonaje.frx":8AE99
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2280
      Width           =   1815
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":8AEC6
      Left            =   9480
      List            =   "frmCrearPersonaje.frx":8AED0
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
   End
   Begin VB.PictureBox PlayerView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   975
      Left            =   1440
      ScaleHeight     =   915
      ScaleWidth      =   780
      TabIndex        =   13
      Top             =   5640
      Width           =   840
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   39
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   38
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   37
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   36
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   35
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   34
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   32
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   31
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   30
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   29
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   26
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label modConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   12
      Top             =   4440
      Width           =   225
   End
   Begin VB.Label modCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   11
      Top             =   4095
      Width           =   225
   End
   Begin VB.Label modInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   10
      Top             =   3435
      Width           =   210
   End
   Begin VB.Label modAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   9
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label modFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   8
      Top             =   3120
      Width           =   210
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6675
      TabIndex        =   7
      Top             =   7275
      Width           =   270
   End
   Begin VB.Image boton 
      Height          =   1245
      Index           =   2
      Left            =   840
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Image boton 
      Height          =   735
      Index           =   1
      Left            =   8160
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   1485
   End
   Begin VB.Image boton 
      Height          =   690
      Index           =   0
      Left            =   9840
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   1440
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   225
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   3840
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If


If cabeza.listIndex < 0 Then
MsgBox "Seleccione su rostro."
Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)

Call Audio.PlayWave("click.wav")

Select Case Index
    Case 0
        
   
        UserName = txtNombre.Text
        
                If Len(txtNombre.Text) < 2 Then
    MsgBox "El nombre debe de tener entre 2 y 15 caracteres."
    Exit Sub
End If
 
If Len(txtNombre.Text) >= 16 Then
    MsgBox "El nombre debe de tener entre 2 y 15 caracteres."
    Exit Sub
End If
        
        Dim AllCr As Long
Dim CantidadEsp As Byte
Dim thiscr As String
 
Do
    AllCr = AllCr + 1
    If AllCr > Len(UserName) Then Exit Do
    thiscr = mid(UserName, AllCr, 1)
    If InStr(1, " ", UCase(thiscr)) = 1 Then
           CantidadEsp = CantidadEsp + 1
    End If
Loop
If CantidadEsp > 1 Then
     MsgBox "El nombre no puede tener mas de 1 espacio."
     Exit Sub
End If
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            frmPasswdSinPadrinos.Show vbModal, Me
        End If
        
    Case 1
    
        Me.Visible = False
        Call SendData("/SALIR")
        
    Case 2
        Call Audio.PlayWave("cupdice.Wav")
        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call SendData("HJHQSC")
    End If

End Sub


Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\CP-Interface.jpg")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

Call TirarDados
End Sub

Private Sub lstRaza_Click()

Call DameOpciones

Select Case (lstRaza.List(lstRaza.listIndex))
    Case Is = "Humano"
        modFuerza.Caption = "+1"
        modConstitucion.Caption = "+2"
        modAgilidad.Caption = "+1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modFuerza.Caption = ""
        modConstitucion.Caption = "+1"
        modAgilidad.Caption = "+4"
        modInteligencia.Caption = "+2"
        modCarisma.Caption = "+2"
    Case Is = "Elfo Oscuro"
        modFuerza.Caption = "+1"
        modConstitucion.Caption = "+1"
        modAgilidad.Caption = "+1"
        modInteligencia.Caption = "+1"
        modCarisma.Caption = "-3"
    Case Is = "Enano"
        modFuerza.Caption = "+3"
        modConstitucion.Caption = "+3"
        modAgilidad.Caption = "-1"
        modInteligencia.Caption = "-3"
        modCarisma.Caption = "-2"
    Case Is = "Gnomo"
        modFuerza.Caption = "-2"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+3"
        modInteligencia.Caption = "+3"
        modCarisma.Caption = "+1"
End Select


End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cabeza_Click()
MiCabeza = Val(cabeza.List(cabeza.listIndex))
Call DibujarCPJ(MiCuerpo, MiCabeza)
End Sub
 
Private Sub lstGenero_Click()
Call DameOpciones
End Sub
