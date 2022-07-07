VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1920
   ClientLeft      =   9360
   ClientTop       =   3240
   ClientWidth     =   2670
   ClipControls    =   0   'False
   FillColor       =   &H00808080&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Timer tmrAutoRemo 
      Interval        =   1
      Left            =   7560
      Top             =   5520
   End
   Begin VB.Timer AutoPotAzules 
      Enabled         =   0   'False
      Interval        =   140
      Left            =   4800
      Top             =   4680
   End
   Begin VB.Timer ShapeInfo 
      Interval        =   750
      Left            =   3120
      Top             =   4800
   End
   Begin VB.Timer check 
      Interval        =   60000
      Left            =   2520
      Top             =   4800
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00151515&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   10320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   7440
      Width           =   4890
   End
   Begin VB.TextBox txtDat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   11280
      TabIndex        =   0
      Top             =   6960
      Width           =   3855
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3600
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrRandom 
      Interval        =   10000
      Left            =   600
      Top             =   4800
   End
   Begin VB.Timer tmrReadMemory 
      Interval        =   1
      Left            =   1080
      Top             =   4800
   End
   Begin VB.Timer tmrEstado 
      Interval        =   1
      Left            =   120
      Top             =   4800
   End
   Begin VB.Timer TimerHwnd 
      Interval        =   1
      Left            =   1560
      Top             =   4800
   End
   Begin VB.Timer AutoPotRojas 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   4320
      Top             =   4680
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   162
      TabIndex        =   2
      Top             =   120
      Width           =   2430
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   195
         Left            =   2040
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer tmrTOP 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1920
         Top             =   0
      End
      Begin VB.Label lblEstadoPotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Desactivado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "AutoPotas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ProloGyX Cheats"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Cmb2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1980
         TabIndex        =   9
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Cmb1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label cmdDat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Como usar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label HpBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "396/396"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   525
         TabIndex        =   4
         Top             =   405
         Width           =   1380
      End
      Begin VB.Shape Hpshp 
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   525
         Top             =   405
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000010&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label ManaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1945/1945"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   525
         TabIndex        =   3
         Top             =   780
         Width           =   1380
      End
      Begin VB.Shape MANShp 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   225
         Left            =   525
         Top             =   765
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000010&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Begin Code

Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function GetDeviceCaps Lib "gdi32" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long


Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private hwnd_Ventana As Long
Private Caption_Ventana As String
Private Length As Long
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" _
               (ByVal hwnd As Long, _
                ByVal lpString As String, _
                ByVal cch As Long) As Long

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '

Dim separa() As String
Dim mManaMax As Long
Dim mVidaMin As Long
Dim mVidaMax As Long
Dim mManaMin As Long
Dim mParalizado As Long
Sub PotearAzules()
            If UsandoPocion = "NO" Then
                Macros ("BuscarCor")
                Call Inicial.setInventario
                Call Inicial.setInventario
                Macros ("VolverCor")
                Sleep (195)
                UsarItem ("Slot02")
                UsandoPocion = "Azules"
                Sleep (195)
                Macros ("BuscarCor")
                Call Inicial.setHechizos
                Call Inicial.setHechizos
                Macros ("VolverCor")
            End If
        If UsandoPocion = "Rojas" Then
            Macros ("BuscarCor")
            Call Inicial.setInventario
            Call Inicial.setInventario
            Macros ("VolverCor")
            Sleep (195)
            UsarItem ("Slot02")
            UsandoPocion = "Azules"
            Sleep (195)
            Macros ("BuscarCor")
            Call Inicial.setHechizos
            Call Inicial.setHechizos
            Macros ("VolverCor")
            Else
            UsandoPocion = "Azules"
            UsarItem ("Azules")
            End If

End Sub
Sub PotearRojas()

            If UsandoPocion = "NO" Then
                    Macros ("BuscarCor")
                    Call Inicial.setInventario
                    Call Inicial.setInventario
                    Macros ("VolverCor")
                    Sleep (195)
                    UsarItem ("Slot01")
                    UsandoPocion = "Rojas"
                    Sleep (195)
                    Macros ("BuscarCor")
                    Call Inicial.setHechizos
                    Call Inicial.setHechizos
                    Macros ("VolverCor")
                End If
                
            If UsandoPocion = "Azules" Then
                Macros ("BuscarCor")
                Call Inicial.setInventario
                Call Inicial.setInventario
                Macros ("VolverCor")
                Sleep (195)
                UsarItem ("Slot01")
                UsandoPocion = "Rojas"
                Sleep (195)
                Macros ("BuscarCor")
                Call Inicial.setHechizos
                Call Inicial.setHechizos
                Macros ("VolverCor")
                Else
                UsandoPocion = "Rojas"
                UsarItem ("Rojas")
            End If

End Sub
Private Sub AutoPotAzules_Timer()
    If MinHp < (MaxHp * PorcentajeRojas) / 100 Then
        Exit Sub
    End If
    
If MinHp < MaxHp Then Exit Sub

    If MinMan < MaxMan Then
        Call PotearAzules
End If
    
End Sub

Private Sub AutoPotRojas_Timer()
If MinHp = "0" Then Exit Sub
If MinHp < "1" Then Exit Sub
If MinMan < (MaxMan * PorcentajeAzules) / 100 Then
Call PotearAzules
End If
    If MinHp < MaxHp Then
        Call PotearRojas
    End If
    
End Sub

Private Sub check_Timer()
On Error GoTo JAJA
Inet1.URL = UserURL
asd = Inet1.OpenURL
If InStr(1, asd, "DOCTYPE") Then
Call UnloadAll
MsgBox "Cheat bloqueado", vbExclamation, "ERROR: Posiblemente"
End
End If
Dim separados() As String
separados = Split(asd, "<>")

Dim i As Integer

For i = LBound(separados) To UBound(separados)

Next i
If separados(1) = "1" Then

Call UnloadAll
MsgBox "Se te ha prohibido el uso del cheat"
End
End If

JAJA:
Exit Sub
End Sub

Private Sub Cmb1_Click()
If cmdDat.Caption = "Salir" Then
cmdDat.Caption = "Config Teclas"
Exit Sub
End If
If cmdDat.Caption = "Config Teclas" Then
cmdDat.Caption = "Intervalos"
Exit Sub
End If
If cmdDat.Caption = "Intervalos" Then
cmdDat.Caption = "Ocultar cheat"
Exit Sub
End If
If cmdDat.Caption = "Ocultar cheat" Then
cmdDat.Caption = "Como usar"
Exit Sub
End If
If cmdDat.Caption = "Como usar" Then
cmdDat.Caption = "Como usar"
Exit Sub
End If
End Sub

Private Sub Cmb1_DblClick()
If cmdDat.Caption = "Salir" Then
cmdDat.Caption = "Config Teclas"
Exit Sub
End If
If cmdDat.Caption = "Config Teclas" Then
cmdDat.Caption = "Intervalos"
Exit Sub
End If
If cmdDat.Caption = "Intervalos" Then
cmdDat.Caption = "Ocultar cheat"
Exit Sub
End If
If cmdDat.Caption = "Ocultar cheat" Then
cmdDat.Caption = "Como usar"
Exit Sub
End If
If cmdDat.Caption = "Como usar" Then
cmdDat.Caption = "Como usar"
Exit Sub
End If
End Sub

Private Sub Cmb1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Cmb2_Click()

If cmdDat.Caption = "Como usar" Then
cmdDat.Caption = "Ocultar cheat"
Exit Sub
End If
If cmdDat.Caption = "Ocultar cheat" Then
cmdDat.Caption = "Intervalos"
Exit Sub
End If
If cmdDat.Caption = "Intervalos" Then
cmdDat.Caption = "Config Teclas"
Exit Sub
End If
If cmdDat.Caption = "Config Teclas" Then
cmdDat.Caption = "Salir"
Exit Sub
End If
If cmdDat.Caption = "Salir" Then
cmdDat.Caption = "Salir"
Exit Sub
End If
End Sub

Private Sub Cmb2_DblClick()
If cmdDat.Caption = "Como usar" Then
cmdDat.Caption = "Ocultar cheat"
Exit Sub
End If
If cmdDat.Caption = "Ocultar cheat" Then
cmdDat.Caption = "Intervalos"
Exit Sub
End If
If cmdDat.Caption = "Intervalos" Then
cmdDat.Caption = "Config Teclas"
Exit Sub
End If
If cmdDat.Caption = "Config Teclas" Then
cmdDat.Caption = "Salir"
Exit Sub
End If
If cmdDat.Caption = "Salir" Then
cmdDat.Caption = "Salir"
Exit Sub
End If
End Sub

Private Sub Cmb2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub cmdDat_Click()

If cmdDat.Caption = "Como usar" Then

Me.Visible = False
vMain = "A"
frmInfo.Show
frmInfo.txtData.Text = "Apretar F1 encima del item (POCION ROJA) o (POCION AZUL)" & vbNewLine & vbNewLine & "En cualquier lugar del inventario tambien sirve."
End If

If cmdDat.Caption = "Ocultar cheat" Then

Me.Visible = False
Inicial.Visible = False
frmInfo.Show

frmInfo.txtData.Text = "Información:" & vbNewLine & "-Cheat oculto Sin embargo el cheat seguira en funcionamiento" & vbNewLine & "Para volver a mostrar apretar: " & CustomKeys.ReadableName(CustomKeys.BindedKey(eKeyType.mKey_OcultarCheat))
End If

If cmdDat.Caption = "Config Teclas" Then

Me.Visible = False
FrmTeclas.Show
End If

If cmdDat.Caption = "Intervalos" Then
Me.Visible = False
Inicial.Visible = False
FrmIntervalos.Visible = True

End If

If cmdDat.Caption = "Salir" Then

End
End If


End Sub

Private Sub cmdDat_DblClick()

If cmdDat.Caption = "Como usar" Then
Call FormNoPrimerPlano(Me)
Call FormNoPrimerPlano(Inicial)
Me.Visible = False
vMain = "A"
frmInfo.Show
frmInfo.txtData.Text = "Apretar F1 encima del item (POCION ROJA) o (POCION AZUL)" & vbNewLine & vbNewLine & "En cualquier lugar del inventario tambien sirve."
End If

If cmdDat.Caption = "Ocultar cheat" Then
Call FormNoPrimerPlano(Me)
Call FormNoPrimerPlano(Inicial)
FrmMain.Visible = False
Inicial.Visible = False
frmInfo.Show

frmInfo.txtData.Text = "Información:" & vbNewLine & "-Cheat oculto Sin embargo el cheat seguira en funcionamiento" & vbNewLine & "Para volver a mostrar apretar: " & CustomKeys.ReadableName(CustomKeys.BindedKey(eKeyType.mKey_OcultarCheat))
End If

If cmdDat.Caption = "Config Teclas" Then
Call FormNoPrimerPlano(Me)
Call FormNoPrimerPlano(Inicial)
Me.Visible = False
FrmTeclas.Show
End If

If cmdDat.Caption = "Intervalos" Then
Me.Visible = False
Inicial.Visible = False
FrmIntervalos.Visible = True

End If

If cmdDat.Caption = "Salir" Then
Call FormNoPrimerPlano(Me)
Call FormNoPrimerPlano(Inicial)
End
End If


End Sub

Private Sub cmdDat_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Command2_Click()
Call UsarItem("Slot01")
End Sub

Private Sub Command3_Click()
Call UsarItem("Slot02")
End Sub

Private Sub Form_Activate()
tmrTOP.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Form_Load()
Dim asd As String
'Load FrmTeclas
Activado = "Desactivado"
actual = "[Info: Cheat Cargado]"

FrmMacros.Show
'FrmIntervalos.Show
SeActiva = "ON"

'' // Seteamos en el complemento INET la URL que anteriormente declaramos como variable
Inet1.URL = ("https://prologyx.do.am/MEMORY.txt")
'' //

'' // Declaramos que ASD va a contener la info obtenida de la url
asd = Inet1.OpenURL


separa = Split(asd, "<>")

Dim i As Integer

For i = LBound(separa) To UBound(separa)

Next i

mVidaMin = separa(0)
mVidaMax = separa(1)
mManaMin = separa(2)
mManaMax = separa(3)
mParalizado = separa(4)
UsandoPocion = "NO"
Call FrmIntervalos.LeerYNYS
End Sub

Private Sub HpBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label4_Click()
FrmMain.Visible = False

frmCoord.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label5_Click()
UsarItem ("Slot02")
End Sub

Private Sub ManaBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub TabStrip1_Click()

End Sub



Private Sub ShapeInfo_Timer()
On Error Resume Next
If FrmMain.Visible = True Then
FrmMain.Hpshp.Width = (((MinHp / 100) / (MaxHp / 100)) * 92)
            FrmMain.HpBar.Caption = MinHp & "/" & MaxHp
            FrmMain.ManaBar.Caption = MinMan & "/" & MaxMan
                    If MaxMan > 0 Then
                FrmMain.MANShp.Width = (((MinMan + 1 / 100) / (MaxMan + 1 / 100)) * 92)
            Else
                FrmMain.MANShp.Width = 0
            End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Text1.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub tmrAutoRemo_Timer()

    If vAutoRemo = "1" Then
        If Paralizado = -1 Then
            Call FrmMacros.Remover
        End If
    End If

End Sub

Private Sub tmrEstado_Timer()
On Error Resume Next
Static nVal As Boolean
If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoPotas)) = 0 Then
If nVal Then

If AutoPotRojas.Enabled = False And AutoPotAzules.Enabled = False Then
AutoPotRojas.Enabled = True
AutoPotAzules.Enabled = True
If FrmMain.Visible = False Then
frmInfo.Show
frmInfo.txtData.Text = "Información:" & vbNewLine & "Autopotas Activado"
End If
lblEstadoPotas.Caption = "Activado"
lblEstadoPotas.ForeColor = &HFFFFFF
Else
If FrmMain.Visible = False Then
frmInfo.Show
frmInfo.txtData.Text = "Información:" & vbNewLine & "Autopotas Desactivado"
End If
AutoPotRojas.Enabled = False
AutoPotAzules.Enabled = False
lblEstadoPotas.Caption = "Desactivado"
lblEstadoPotas.ForeColor = &H808080
End If
End If
End If
nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoPotas)))
End Sub

Private Sub tmrReadMemory_Timer()
On Error Resume Next

    MinHp = LeerMemoria(mVidaMin, AO)
    MaxHp = LeerMemoria(mVidaMax, AO)
    MinMan = LeerMemoria(mManaMin, AO)
    MaxMan = LeerMemoria(mManaMax, AO)
    'Paralizado = LeerMemoria(mParalizado, AO)
    AutoPotRojas.Interval = IntervaloAutoRojas
    AutoPotAzules.Interval = IntervaloAutoAzules
    FrmMacros.AutoLanzar.Interval = IntervaloAutoLanzar


End Sub

Private Sub tmrTOP_Timer()

tmrTOP.Enabled = False
End Sub

