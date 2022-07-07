VERSION 5.00
Begin VB.Form FrmMacros 
   BorderStyle     =   0  'None
   ClientHeight    =   1770
   ClientLeft      =   11700
   ClientTop       =   9135
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer AutoLanzaAimbot 
      Enabled         =   0   'False
      Interval        =   410
      Left            =   4200
      Top             =   240
   End
   Begin VB.Timer tmrParalizar 
      Interval        =   1
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer tmrSpecial 
      Interval        =   1
      Left            =   3360
      Top             =   360
   End
   Begin VB.Timer Timer7 
      Left            =   1560
      Top             =   2520
   End
   Begin VB.Timer Timer6 
      Left            =   1080
      Top             =   2520
   End
   Begin VB.Timer ActivaLanzar2 
      Interval        =   400
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   2880
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   280
      Left            =   2400
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1920
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   360
   End
   Begin VB.Timer Combo 
      Interval        =   1
      Left            =   960
      Top             =   360
   End
   Begin VB.Timer AutFlechaz 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   480
      Top             =   360
   End
   Begin VB.Timer ActFlechaz 
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer Restar 
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Sumar 
      Interval        =   1
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer ActivaLanzar 
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer AutoLanzar 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer AutoRemo 
      Interval        =   1
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer LanzarUno 
      Interval        =   1
      Left            =   1920
      Top             =   0
   End
   Begin VB.Label y 
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label x 
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "FrmMacros"
Attribute VB_Exposed = False
'Begin Code


Private Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Const KEYEVENTF_KEYUP = &H2
Const KEYEVENTF_EXTENDEDKEY = &H1
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, _
ByVal bScan As Byte, _
ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)
Const HWND_TOPMOST = -1
Private Type POINTAPI
        x As Long
        y As Long
End Type
Sub EnviarTecla(Tecla As Long)
Call keybd_event(Tecla, 0, 0, 0)
Call keybd_event(Tecla, 0, KEYEVENTF_KEYUP, 0)
End Sub


Private Sub Form_Load()
FrmMacros.Visible = False
FrmMacros.Enabled = False
FrmMacros.Height = "0"
FrmMacros.Width = "0"
Lanzar = "ON"
Remo = "ON"
Para = "off"
Espe = "off"
End Sub
Private Sub LanzarUno_Timer()
Static nVal As Boolean
If IsWindowVisible(texthwnd) = 1 Or (GetForegroundWindow <> hwndexe And GetForegroundWindow <> FrmMain.hwnd) Then Exit Sub


If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_LanzarUno)) = 0 Then
If nVal Then

If Lanzar = "ON" Then
If MaxMan > "0" Then

Macros ("BuscarCor")

If vAutoRemo = "1" And Paralizado = -1 Then Exit Sub
If IsWindowVisible(hwndInventario) = 1 Then
Call Inicial.setHechizos
Sleep (20)
Call Inicial.sndLanzar
Else
Call Inicial.sndLanzar
End If

Macros ("VolverCor")
End If
End If
End If
End If
nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_LanzarUno)))

End Sub
Public Sub Remover()
If Not MaxMan > 0 Then Exit Sub
Call EsperaMs(IntervaloAutoRemo)
                Inicial.MainNewMago.Enabled = False
                
                
                    If IsWindowVisible(hwndInventario) = 1 Then
                    Macros ("BuscarCor")
                        Call Inicial.setHechizos
                        Call EsperaMs(30)
                    Macros ("VolverCor")
                    End If
                    
                Memoria.UsarItem ("setRemover")
                Call ClearRichTextBox(RichTextBoxHwnd)
                Call Inicial.sndLanzar
                Call EsperaMs(300)
                Call Inicial.ClickOnMe
                Call Inicial.ClickOnMe
                Call ClearRichTextBox(RichTextBoxHwnd)
                Inicial.MainNewMago.Enabled = True
                
End Sub
Private Sub AutoRemo_Timer()
Static nVal As Boolean
If IsWindowVisible(texthwnd) = 1 Or (GetForegroundWindow <> hwndexe And GetForegroundWindow <> FrmMain.hwnd) Then Exit Sub

If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoRemo)) = 0 Then
If nVal Then

If Remo = "ON" Then
                Inicial.MainNewMago.Enabled = False
                
                
                    If IsWindowVisible(hwndInventario) = 1 Then
                    Macros ("BuscarCor")
                        Call Inicial.setHechizos
                        Call EsperaMs(30)
                    Macros ("VolverCor")
                    End If
                    
                Memoria.UsarItem ("setRemover")
                Call ClearRichTextBox(RichTextBoxHwnd)
                Call Inicial.sndLanzar
                Call EsperaMs(300)
                Call Inicial.ClickOnMe
                Call Inicial.ClickOnMe
                Call ClearRichTextBox(RichTextBoxHwnd)
                Inicial.MainNewMago.Enabled = True
                
End If
End If
End If
nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoRemo)))


End Sub

Private Sub AutoLanzar_Timer()
Macros ("BuscarCor")
If IsWindowVisible(hwndInventario) = 1 Then
Call Inicial.setHechizos
Sleep (20)
Call Inicial.sndLanzar
Else
Call Inicial.sndLanzar
End If
Macros ("ClickCor")
End Sub
Private Sub ActivaLanzar_Timer()
Static nVal As Boolean
If IsWindowVisible(texthwnd) = 1 Or (GetForegroundWindow <> hwndexe And GetForegroundWindow <> FrmMain.hwnd) Then Exit Sub

If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoLanzar)) = 0 Then
If nVal Then

If MaxMan = "0" Then Exit Sub

If AutoLanzar.Enabled = True Then
AutoLanzar.Enabled = False
Else
AutoLanzar.Enabled = True
End If
End If
End If
nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoLanzar)))
End Sub
Private Sub Timer4_Timer()
Static nVal As Boolean
If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_OcultarCheat)) = 0 Then
If nVal Then
'If IsWindowVisible(texthwnd) = 1 Or (GetForegroundWindow <> hwndexe And GetForegroundWindow <> Me.hwnd) Then Exit Sub

If Not FrmMain.Visible = True Then
FrmMain.Visible = True
Inicial.Visible = True
Call FormPrimerPlano(FrmMain)
Call FormPrimerPlano(Inicial)

Else
FrmMain.Visible = False
Inicial.Visible = False
Call FormNoPrimerPlano(FrmMain)
Call FormNoPrimerPlano(Inicial)

 frmInfo.Visible = True
frmInfo.txtData.Text = "Información:" & vbNewLine & "-Cheat oculto Sin embargo el cheat seguira en funcionamiento" & vbNewLine & "Para volver a mostrar apretar: " & CustomKeys.ReadableName(CustomKeys.BindedKey(eKeyType.mKey_OcultarCheat))
End If
End If
End If
nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_OcultarCheat)))
End Sub

Private Sub Timer5_Timer()

End Sub

Private Sub tmrParalizar_Timer()
Static nVal As Boolean
If IsWindowVisible(texthwnd) = 1 Or (GetForegroundWindow <> hwndexe And GetForegroundWindow <> FrmMain.hwnd) Then Exit Sub

If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_HechizoUno)) = 0 Then
    If nVal Then

        If Lanzar = "ON" Then
                If MaxMan > "0" Then
                    Inicial.MainNewMago.Enabled = False
                    
                        If vAutoRemo = "1" And Paralizado = -1 Then Exit Sub
                    
                        If IsWindowVisible(hwndInventario) = 1 Then
                            Macros ("BuscarCor")
                                Call Inicial.setHechizos
                                Sleep (30)
                            Macros ("VolverCor")
                        End If
                        
                    Memoria.UsarItem ("setParalizar")
                    
                    Call Inicial.sndLanzar
                    
                        If vModoMacros = 0 Then
                        Call Inicial.LanzarAim
                        End If
                    
                    
                    Inicial.MainNewMago.Enabled = True
                Else
                UsarItem ("ARCO")
                Call Inicial.LanzarAimCaza
                End If
        End If
    End If
End If
nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_HechizoUno)))



End Sub

Private Sub tmrSpecial_Timer()
Static nVal As Boolean
If IsWindowVisible(texthwnd) = 1 Or (GetForegroundWindow <> hwndexe And GetForegroundWindow <> FrmMain.hwnd) Then Exit Sub

If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_HechizoDos)) = 0 Then
If nVal Then

        If Lanzar = "ON" Then
                If MaxMan > "0" Then
                Inicial.MainNewMago.Enabled = False
                    If vAutoRemo = "1" And Paralizado = -1 Then Exit Sub


   
                    If IsWindowVisible(hwndInventario) = 1 Then
                    Macros ("BuscarCor")
                        Call Inicial.setHechizos
                        Sleep (30)
                    Macros ("VolverCor")
                    End If
                    
                Memoria.UsarItem ("setEspecial")
                
                Call Inicial.sndLanzar
                If vModoMacros = 0 Then
                Call Inicial.LanzarAim
                End If
                
                Inicial.MainNewMago.Enabled = True
                
                End If
        End If
End If
End If
nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_HechizoDos)))





End Sub

