VERSION 5.00
Begin VB.Form Inicial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "asd"
   ClientHeight    =   4290
   ClientLeft      =   2385
   ClientTop       =   3195
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   7680
      Width           =   4335
   End
   Begin VB.Timer tmrTarget 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   6960
   End
   Begin VB.Timer GetRenderPos 
      Interval        =   1
      Left            =   5520
      Top             =   6960
   End
   Begin VB.ListBox lstTMP 
      Height          =   2205
      Left            =   5280
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstInRange 
      BackColor       =   &H00000000&
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
      Height          =   4350
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
   Begin VB.Timer posbyTarget 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6840
      Top             =   6960
   End
   Begin VB.Timer CheckChars 
      Interval        =   500
      Left            =   3600
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   7320
      Width           =   4335
   End
   Begin VB.Timer tmrConsola 
      Interval        =   50
      Left            =   7920
      Top             =   5280
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   8040
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   5280
      Width           =   615
   End
   Begin VB.Timer Mouse2 
      Interval        =   1
      Left            =   8280
      Top             =   5280
   End
   Begin VB.Timer Lanzar 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   6960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   6600
      TabIndex        =   0
      Top             =   3360
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   240
   End
   Begin VB.Timer MainNewMago 
      Interval        =   1
      Left            =   7560
      Top             =   5280
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label lblTest 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblTarget 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Target:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label lblPosX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPosY 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pos X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pos Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "Inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Begin Code

Option Explicit

Dim addressEntityPlayers As Long
Private Declare Function ScreenToClient Lib "User32" ( _
    ByVal hwnd As Long, _
    lpPoint As POINTAPI) As Long
  
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Dim pt As POINTAPI
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Pntt As POINTAPI
Private PtAim As POINTAPI
Dim x As Integer, y As Integer

Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function ClientToScreen Lib "User32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long
Private KA As Integer
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDOWN As Long = &H201

Private Const KEY_DOWN    As Integer = &H8000
Private Const KEY_PRESSED As Integer = &H1
Private FinalPosX As Integer
Private FinalPosY As Integer
'-------------------------- GWR
Private Declare Function GetWindowRect Lib "User32" ( _
    ByVal hwnd As Long, _
    lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Const ERROR_INVALID_WINDOW_HANDLE As Long = 1400

'-------------------------- GWR

Sub PosRenderInScreen(el_Hwnd As Long)
  
  
End Sub

'
'Sub updateEntity(idPlayer As Long, addressByRef As Long, name As String, posX As Long, posY As Long)
'    CharList(idPlayer).name = name
'    CharList(idPlayer).posX = posX
'    CharList(idPlayer).posY = posY
'    CharList(idPlayer).address = address
'End Sub
'Sub addEntity(idPlayer As Long, address As Long, name As String, posX As Long, posY As Long)
'    CharList(idPlayer).name = name
'    CharList(idPlayer).posX = posX
'    CharList(idPlayer).posY = posY
'    CharList(idPlayer).address = address
'End Sub
'Sub checkEntity(address As Long, name As String, posX As Long, posY As Long)
'    If CharList.HasKey(name) Then
'      Call updateEntity(name, address, name, posX, posY)
'    Else
'      Call addEntity(name, address, name, posX, posY)
'      End If
'End Sub
'-Mouse
Sub ClickOnMe()
Call SendClick(PictureRenderhWnd, 367, 208)
Call EsperaMs(200)
Call ClearRichTextBox(RichTextBoxHwnd)
Call msgREMO(RichTextBoxHwnd)
                
End Sub
Sub MoveMouse(x As Single, y As Single)
Dim pt As POINTAPI

    pt.x = MOUSE_POSITION_ME_X
    pt.y = MOUSE_POSITION_ME_Y

    ClientToScreen PictureRenderhWnd, pt
    If pt.x < 0 Or pt.y < 0 Then Exit Sub
    
    SetCursorPos pt.x, pt.y
End Sub
Sub setHechizos()
Dim pt As POINTAPI

    pt.x = 1200
    pt.y = 240

    ClientToScreen hwndexe, pt
    If pt.x < 0 Or pt.y < 0 Then Exit Sub
    
    SetCursorPos pt.x, pt.y
    
    Call SendClicks.LeftClick
End Sub
Sub setInventario()
Dim pt As POINTAPI

    pt.x = 1085
    pt.y = 240

    ClientToScreen hwndexe, pt
    If pt.x < 0 Or pt.y < 0 Then Exit Sub
    
    SetCursorPos pt.x, pt.y
    
   Call SendClicks.LeftClick
End Sub

Private Sub Aimming_Timer()
Call Aimbot
'ETime2 = GetTickCount
'Aimming.Enabled = False
MainNewMago.Enabled = True
End Sub

Private Sub CheckChars_Timer()
Dim tmpChar As tCharlist

If LastCharTargeted <> -1 Then
    lblTarget.Caption = GetUserNameByIndex(LastCharTargeted)
End If

lstTMP.Clear

MyPosX = mm.readInteger(ADDRESS_MY_POSX)
MyPosY = mm.readInteger(ADDRESS_MY_POSY)
MyPosX2 = MyPosX + 1

For x = 1 To MAX_PLAYERS
    With tmpChar
        'Acá leo todos los address que voy a necesitar
        'address = addressEntityPlayers + (x - 1) * SIZE_PLAYERS
        'PosX = mm.readInteger(address + OFFSET_POSX)
        'PosY = mm.readInteger(address + OFFSET_POSY)
        'addressName = mm.readLong(address + OFFSET_NAME)
        
        .address = addressEntityPlayers + (x - 1) * SIZE_PLAYERS '(SIZE_PLAYERS + (4 * 1))
        .Active = mm.readLong(.address + eOffsets.oActive)
        .Heading = mm.readLong(.address + eOffsets.oHeading)
        .PosX = mm.readInteger(.address + eOffsets.oPosX)
        .PosY = mm.readInteger(.address + eOffsets.oPosY)
        .addressName = mm.readLong(.address + eOffsets.oName)
        .name = ""
        .ScrollDirectionX = mm.readInteger(.address + eOffsets.oScrollDirectionX)
        .ScrollDirectionY = mm.readInteger(.address + eOffsets.oScrollDirectionY)
        .Moving = mm.readByte(.address + eOffsets.oMoving)
        
        'es user activo?
        If .Active = 1 Then
            'limites del mapa¿? sí
            If (.PosX > 5 And .PosX < 95) And (.PosY > 5 And .PosY < 95) Then
                'Está dentro de mi rango de visión?
                If Abs(MyPosX - .PosX) <= RANGE_X And Abs(MyPosY - .PosY) <= RANGE_Y Then
                    'Tiene nick?
                    If (.addressName > 0) Then
                        .name = mm.readString(255, .addressName)
                        
                        'Si tiene nick lo agrego a la lista temporal
                        If Len(.name) > 0 Then
                            lstTMP.AddItem .name
                        End If
                    End If
                End If
            End If
        End If
    End With
    
    Call SetUserByChar((x - 1), tmpChar)
        
    'Los agrego a todos, no importa si están asignados como usuarios/npcs o no
    'Call SetUser((x - 1), address, addressName, name, PosX, PosY)
Next x

'Si el que tengo en la lista, no aparece en el listboxtemporal, lo seteamos delete
For x = LBound(InRangeList) To UBound(InRangeList)
    For y = 0 To lstTMP.ListCount - 1
        If InRangeList(x).name = lstTMP.List(y) Then GoTo Pass
    Next y
    
    InRangeList(x).Delete = True
    
Pass:
Next x

'Acá limpia todos los que están seteandos delete
Call CleanList

add:
'Acá agregamos a los nuevos que aparecieron
For x = 0 To lstTMP.ListCount - 1
    Call AddToList(lstTMP.List(x))
Next x

'Removemos los de la lista listbox que no están en nuestra lista
For x = (lstInRange.ListCount - 1) To 0 Step -1
    For y = LBound(InRangeList) To UBound(InRangeList)
        If lstInRange.List(x) = InRangeList(y).name And Not InRangeList(y).Delete Then GoTo pass2
    Next y
    
    lstInRange.RemoveItem x

pass2:
Next x

'Agregamos a la lista listbox los que están en nuestra lista
For x = LBound(InRangeList) To UBound(InRangeList)
    For y = 0 To lstInRange.ListCount - 1
        If Len(InRangeList(x).name) > 0 Then
            If InRangeList(x).name = lstInRange.List(y) And Not InRangeList(x).Delete Then GoTo pass3
        End If
    Next y
    
    If Len(InRangeList(x).name) > 0 Then lstInRange.AddItem InRangeList(x).name
    
pass3:
Next x

End Sub





Private Sub Form_Activate()
Call FormPrimerPlano(Inicial)
End Sub

Private Sub Form_Load()
Dim hwnd As Long, pid As Long



GetWindowThreadProcessId hwndexe, pid
mm.pid = pid
addressEntityPlayers = mm.readLong(BASE_PLAYERS)

ReDim InRangeList(0)
LastCharTargeted = -1

'PictureRenderhWnd = 0

LastPosX = 100
LastPosY = 100
BestMoveTime = 1000
BestTick = GetTickCount()
BestTickLast = GetTickCount()
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FormNoPrimerPlano(Me)
End Sub

Private Sub GetRenderPos_Timer()
Dim t_Rect As RECT
  
      
    If GetWindowRect(PictureRenderhWnd, t_Rect) = 0 Then
        'If Err.LastDllError = ERROR_INVALID_WINDOW_HANDLE Then
            'MsgBox " Hwnd no válido  ", vbCritical
        'End If
    Else
        
        'Macros ("BuscarCor")
        RenderTop = t_Rect.Top
        
        RenderLeft = t_Rect.Left
        
        RenderRight = t_Rect.Right
        
        RenderBottom = t_Rect.Bottom
        
        'Ésto está bien
        MOUSE_POSITION_ME_X = (t_Rect.Right - t_Rect.Left) / 2 - 16
        MOUSE_POSITION_ME_Y = (t_Rect.Bottom - t_Rect.Top) / 2 - 16
        
        '394-1130 (736)
        '256-736 (480)
        'MyPos = 760,470 (50, 50)
        
        '607-1343 (736)
        '399-879 (480)
        'MyPos = 970,610 (50,50)
        
        'Left+210 = MyMousePosX
        'Top+140 = MyMouseY
        
        'Osea, si quiero hacer un setcursorpos siempre sobre mi pj sería ése
        'Fijate de poner un botón para testear
        
        Dim pt As POINTAPI
        GetCursorPos pt
        
        Dim DeltaPosXPJ As Integer
        Dim DeltaPosYPJ As Integer

        
        MyPosX = mm.readInteger(ADDRESS_MY_POSX)
        MyPosY = mm.readInteger(ADDRESS_MY_POSY)
            
        DeltaPosXPJ = (pt.x - (RenderLeft + 360)) / 32
        DeltaPosYPJ = (pt.y - (RenderTop + 220)) / 32
        
        FinalPosX = MyPosX + DeltaPosXPJ
        FinalPosY = MyPosY + DeltaPosYPJ
        
        'lblTest.Caption = FinalPosX & " - " & FinalPosY
    End If
    
End Sub
Private Sub Command3_Click()
Dim t_Rect As RECT
  
      
If GetWindowRect(PictureRenderhWnd, t_Rect) = 0 Then Exit Sub

Call SetCursorPos(t_Rect.Left + 360, t_Rect.Top + 220)
End Sub

Private Sub Label2_Click()
MsgBox resX & " - " & resY
End Sub
Sub LanzarAimCaza()

If LastCharTargeted = -1 Then
    Estado = eEstado.None
    Exit Sub
End If

Static nVal As Boolean

Static Activated As Boolean
Dim ret
Dim pcin As PCURSORINFO
pcin.cbSize = Len(pcin)
ret = GetCursorInfo(pcin)




            Dim tmpChar As tCharlist
            Dim Rdy As Boolean
            
                    Call EsperaMs(100)
            Rdy = False
            
            With tmpChar
                .address = GetAddressByIndex(LastCharTargeted)
                
                .Active = mm.readLong(.address + eOffsets.oActive)
                .Heading = mm.readLong(.address + eOffsets.oHeading)
                .PosX = mm.readInteger(.address + eOffsets.oPosX)
                .PosY = mm.readInteger(.address + eOffsets.oPosY)
                .addressName = mm.readLong(.address + eOffsets.oName)
                .name = ""
                .ScrollDirectionX = mm.readInteger(.address + eOffsets.oScrollDirectionX)
                .ScrollDirectionY = mm.readInteger(.address + eOffsets.oScrollDirectionY)
                .Moving = mm.readByte(.address + eOffsets.oMoving)

                '---------------------------------------------------
                WaitPosX = .PosX
                WaitPosY = .PosY
                
                WaitTick = GetTickCount()
                
                While GetTickCount() - WaitTick < BestMoveTime
                    .PosX = mm.readInteger(.address + eOffsets.oPosX)
                    .PosY = mm.readInteger(.address + eOffsets.oPosY)
                    .Moving = mm.readByte(.address + eOffsets.oMoving)
                    
                    If WaitPosX <> .PosX Or WaitPosY <> .PosY Then Rdy = True
                Wend
                '---------------------------------------------------
                
                resX = ((.PosX - mm.readInteger(ADDRESS_MY_POSX)) * 32) + MOUSE_POSITION_ME_X + 16
                resY = (((.PosY - 1) - mm.readInteger(ADDRESS_MY_POSY)) * 32) + MOUSE_POSITION_ME_Y + 16
                
                'GoTo label_normal
                
                'If .Moving Then
                If Not Rdy Then
                    'Se mueve hacia la derecha
                    If .ScrollDirectionX = 1 Then
                        resX = resX + 16
                    Else
                        resX = resX - 16
                    End If
                    
                    'Se mueve hacia abajo
                    If .ScrollDirectionY = 1 Then
                        resY = resY + 16 '16
                    Else
                        resY = resY - 8 '8
                    End If
                End If
label_normal:
            End With
            'Call MoveMouse(resX, resY)
            If ret And pcin.hCursor <> 65547 Then
            Call SendClick(PictureRenderhWnd, resX, resY)
            End If

            

End Sub
Sub LanzarAim()

If LastCharTargeted = -1 Then
    Estado = eEstado.None
    Exit Sub
End If

Static nVal As Boolean

Static Activated As Boolean


            Dim tmpChar As tCharlist
            Dim Rdy As Boolean
            
                     Call EsperaMs(100)
            Rdy = False
            
            With tmpChar
                .address = GetAddressByIndex(LastCharTargeted)
                
                .Active = mm.readLong(.address + eOffsets.oActive)
                .Heading = mm.readLong(.address + eOffsets.oHeading)
                .PosX = mm.readInteger(.address + eOffsets.oPosX)
                .PosY = mm.readInteger(.address + eOffsets.oPosY)
                .addressName = mm.readLong(.address + eOffsets.oName)
                .name = ""
                .ScrollDirectionX = mm.readInteger(.address + eOffsets.oScrollDirectionX)
                .ScrollDirectionY = mm.readInteger(.address + eOffsets.oScrollDirectionY)
                .Moving = mm.readByte(.address + eOffsets.oMoving)

                '---------------------------------------------------
                WaitPosX = .PosX
                WaitPosY = .PosY
                
                WaitTick = GetTickCount()
                
                While GetTickCount() - WaitTick < BestMoveTime
                    .PosX = mm.readInteger(.address + eOffsets.oPosX)
                    .PosY = mm.readInteger(.address + eOffsets.oPosY)
                    .Moving = mm.readByte(.address + eOffsets.oMoving)
                    
                    If WaitPosX <> .PosX Or WaitPosY <> .PosY Then Rdy = True
                Wend
                '---------------------------------------------------
                
                resX = ((.PosX - mm.readInteger(ADDRESS_MY_POSX)) * 32) + MOUSE_POSITION_ME_X + 16
                resY = (((.PosY - 1) - mm.readInteger(ADDRESS_MY_POSY)) * 32) + MOUSE_POSITION_ME_Y + 16
                
                'GoTo label_normal
                
                'If .Moving Then
                If Not Rdy Then
                    'Se mueve hacia la derecha
                    If .ScrollDirectionX = 1 Then
                        resX = resX + 16
                    Else
                        resX = resX - 16
                    End If
                    
                    'Se mueve hacia abajo
                    If .ScrollDirectionY = 1 Then
                        resY = resY + 16 '16
                    Else
                        resY = resY - 8 '8
                    End If
                End If
label_normal:
            End With
            'Call MoveMouse(resX, resY)
            
            Call SendClick(PictureRenderhWnd, resX, resY)
            

End Sub
Sub sndLanzar()
                    Dim DaWord As Long
                    
                    '1200, 230 hechizos
                    '1080, 230 inventario
                    DaWord = MakeDWord(1100, 481)
                    'PictureRenderhWnd = 987672
                    'Call SendClick(MainAOhWnd, 1100, 481)
                    'este es el boton lanzar
                    SendMessage hwndexe, WM_LBUTTONDOWN, 1&, ByVal DaWord
                    SendMessage hwndexe, WM_LBUTTONUP, 1&, ByVal DaWord
End Sub
Private Sub MainNewMago_Timer()

If LastCharTargeted = -1 Then
    Estado = eEstado.None
    Exit Sub
End If
If vModoMacros = 1 Then Exit Sub
If IsWindowVisible(texthwnd) = 1 Or (GetForegroundWindow <> hwndexe And GetForegroundWindow <> FrmMain.hwnd) Then Exit Sub

Static nVal As Boolean
Static Activated As Boolean


Dim ret
Dim pcin As PCURSORINFO
pcin.cbSize = Len(pcin)
ret = GetCursorInfo(pcin)


'If ret Then
'    Me.Caption = pcin.hCursor
'End If

'hacerlo con gettickcount o otra cosa para que no pare la app esos 100ms? sí
'o la manera que dijiste de tomar la consola
'para setear flag CastedSpell = true
'el tema con eso que si justo en ese momento aparece otro mensaje en consola se va a la puta xd

'If GetAsyncKeyState(vbKey1) = 0 Then
If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoAim)) = 0 Then



If nVal Then Activated = Not Activated

If Activated Then
    Select Case Estado
        Case eEstado.None
            ret = GetCursorInfo(pcin)
            
            'If ret And pcin.hCursor <> 65547 Then
                'If nVal Then
                    'acá va la funcion cuando apretas 1
                    'Me.Caption = Me.Caption + 1
                    If MaxMan > 0 Then
                    Call sndLanzar
                    Else
                    UsarItem ("ARCO")
                    End If
                    Estado = eEstado.CastedSpell
                    
                    'hago una pausa para que aparezca el cursor
                    Call EsperaMs(300)
                    'Tick_B4Throw = GetTickCount()
                    
                    'Call SendClick(PictureRenderhWnd, resX, resY)
                    'Me.Caption = Rnd * 3
                'End If
            'End If
        Case eEstado.CastedSpell
            Dim tmpChar As tCharlist
            Dim Rdy As Boolean
            
            Rdy = False
            
            With tmpChar
                .address = GetAddressByIndex(LastCharTargeted)
                
                .Active = mm.readLong(.address + eOffsets.oActive)
                .Heading = mm.readLong(.address + eOffsets.oHeading)
                .PosX = mm.readInteger(.address + eOffsets.oPosX)
                .PosY = mm.readInteger(.address + eOffsets.oPosY)
                .addressName = mm.readLong(.address + eOffsets.oName)
                .name = ""
                .ScrollDirectionX = mm.readInteger(.address + eOffsets.oScrollDirectionX)
                .ScrollDirectionY = mm.readInteger(.address + eOffsets.oScrollDirectionY)
                .Moving = mm.readByte(.address + eOffsets.oMoving)

                '---------------------------------------------------
                WaitPosX = .PosX
                WaitPosY = .PosY
                
                WaitTick = GetTickCount()
                
                While GetTickCount() - WaitTick < BestMoveTime
                    .PosX = mm.readInteger(.address + eOffsets.oPosX)
                    .PosY = mm.readInteger(.address + eOffsets.oPosY)
                    .Moving = mm.readByte(.address + eOffsets.oMoving)
                    
                    If WaitPosX <> .PosX Or WaitPosY <> .PosY Then Rdy = True
                Wend
                '---------------------------------------------------
                
                resX = ((.PosX - mm.readInteger(ADDRESS_MY_POSX)) * 32) + MOUSE_POSITION_ME_X + 16
                resY = (((.PosY - 1) - mm.readInteger(ADDRESS_MY_POSY)) * 32) + MOUSE_POSITION_ME_Y + 16
                
                'GoTo label_normal
                
                'If .Moving Then
                If Not Rdy Then
                    'Se mueve hacia la derecha
                    If .ScrollDirectionX = 1 Then
                        resX = resX + 16
                    Else
                        resX = resX - 16
                    End If
                    
                    'Se mueve hacia abajo
                    If .ScrollDirectionY = 1 Then
                        resY = resY + 16 '16
                    Else
                        resY = resY - 8 '8
                    End If
                End If
label_normal:
            End With
            'Call MoveMouse(resX, resY)
            If ret And pcin.hCursor <> 65547 Then
            Call SendClick(PictureRenderhWnd, resX, resY)
            End If
            Estado = eEstado.None
    End Select
End If

'Me.Caption = Activated & " - " & BestMoveTime
End If

nVal = CBool(GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKey_AutoAim)))
End Sub

Public Sub SendClick(ByVal vhWnd As Long, ByVal PosX As Long, ByVal PosY As Long)
Dim DaWord As Long

DaWord = SendClicks.MakeDWord(PosX, PosY)

Call PostMessageA(vhWnd, WM_SETFOCUS, 0&, 0&)
Call PostMessageA(vhWnd, WM_MOUSEMOVE, 0&, ByVal DaWord)
Call SendClicks.LeftClick
End Sub

Private Sub Mouse2_Timer()
'On Error Resume Next
If LastCharTargeted = -1 Then Exit Sub
'mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

    'Call SetCursorPos(resX, resY)
    'resX = ((GetPosXByIndex(LastCharTargeted) - MyPosX) * 32) + MOUSE_POSITION_ME_X
    'resY = ((GetPosYByIndex(LastCharTargeted) - MyPosY) * 32) + MOUSE_POSITION_ME_Y
    'Call SetCursorPos(resX, resY)
    'Call LeftClick
    'Call SetCursorPos(resX, resY)

    'Call SetCursorPos(resX, resY)

Dim tmpChar As tCharlist

With tmpChar
    .address = GetAddressByIndex(LastCharTargeted)
    
    .PosX = mm.readInteger(.address + eOffsets.oPosX)
    .PosY = mm.readInteger(.address + eOffsets.oPosY)
    
    If .PosX <> LastPosX Or .PosY <> LastPosY Then
        LastPosX = .PosX
        LastPosY = .PosY
        
        BestTick = GetTickCount()
        
        If BestTick - BestTickLast < BestMoveTime Then
            BestMoveTime = BestTick - BestTickLast
        End If
        
        'Correct lag
        If BestMoveTime < 50 Then BestMoveTime = 300
        
        BestTickLast = GetTickCount()
    End If
End With
End Sub

Public Sub lstInRange_DblClick()
Dim name As String
Dim index As Integer

'Obtengo el texto según el index del listbox
name = lstInRange.List(lstInRange.ListIndex)

'Busco en todos los chars por nombre
index = GetUserByName(name)

'si es válido
If index <> -1 Then
    'Reset Target
    If LastCharTargeted <> -1 Then Call SetUserTarget(LastCharTargeted, False)
    
    'Set New Target
    Call SetUserTarget(index, True)
    LastCharTargeted = index
    posbyTarget.Enabled = True
End If
End Sub

Private Sub posbyTarget_Timer()
'If LastCharTargeted <> -1 Then
'    resX = ((GetPosXByIndex(LastCharTargeted) - MyPosX) * 32) + MOUSE_POSITION_ME_X
'    resY = ((GetPosYByIndex(LastCharTargeted) - MyPosY) * 32) + MOUSE_POSITION_ME_Y
'End If
End Sub

Private Sub Timer1_Timer()
'Call MoveMouse(resX, resY)
End Sub

Private Sub tmrConsola_Timer()
Text1.Text = WindowTextGetLastLine(RichTextBoxHwnd)
Text3.Text = WindowTextGetPreviousLine(RichTextBoxHwnd)
Dim asd As String
Dim asd2 As String
asd = Text1.Text
asd2 = Text3.Text
Dim separados() As String
Dim i As Integer
'Leo la ultima linea
If InStr(asd, "Ves a Condesa") Then
'MsgBox "paso"
separados = Split(asd, " ")

For i = LBound(separados) To 3
    Text2.Text = separados(3)
    Text2.Text = Replace(Text2.Text, ",", "")
    Call Buscar_ListBox(Text2.Text, CBool(Check1.Value))
    Exit Sub
Next i
'MsgBox "paso"

End If
If InStr(asd, "Ves a Conde") Then
'MsgBox "paso"
separados = Split(asd, " ")

For i = LBound(separados) To 3
    Text2.Text = separados(3)
    Text2.Text = Replace(Text2.Text, ",", "")
    Call Buscar_ListBox(Text2.Text, CBool(Check1.Value))
    Exit Sub
Next i
'MsgBox "paso"

End If
If InStr(asd, "Ves a") Then
'MsgBox "paso"
separados = Split(asd, " ")

For i = LBound(separados) To 2
    Text2.Text = separados(2)
    Text2.Text = Replace(Text2.Text, ",", "")
    Call Buscar_ListBox(Text2.Text, CBool(Check1.Value))
    Exit Sub
    Next i
End If
'Si no encuentro en la ultima, leo en la anteultima
If InStr(asd2, "Ves a Condesa") Then
'MsgBox "paso"
separados = Split(asd2, " ")

For i = LBound(separados) To 3
    Text2.Text = separados(3)
    Text2.Text = Replace(Text2.Text, ",", "")
    Call Buscar_ListBox(Text2.Text, CBool(Check1.Value))
    Exit Sub
Next i
'MsgBox "paso"

End If
If InStr(asd2, "Ves a Conde") Then
'MsgBox "paso"
separados = Split(asd2, " ")

For i = LBound(separados) To 3
    Text2.Text = separados(3)
    Text2.Text = Replace(Text2.Text, ",", "")
    Call Buscar_ListBox(Text2.Text, CBool(Check1.Value))
    Exit Sub
Next i
'MsgBox "paso"

End If
If InStr(asd2, "Ves a") Then
'MsgBox "paso"
separados = Split(asd2, " ")

For i = LBound(separados) To 2
    Text2.Text = separados(2)
    Text2.Text = Replace(Text2.Text, ",", "")
    Call Buscar_ListBox(Text2.Text, CBool(Check1.Value))
    Exit Sub
Next i
'MsgBox "paso"

End If




End Sub


Private Sub tmrSearchTarget_Timer()

End Sub

Private Sub tmrTarget_Timer()
'If LastCharTargeted = -1 Then
Dim name As String

Static nVal As Boolean

If GetAsyncKeyState(vbKey3) = 0 Then
If nVal Then

Dim i As Integer

'Loopeo todos los usuarios hasta encontrar el nombre por parámetro
For i = 1 To MAX_PLAYERS
    If CharList(i - 1).Active Then
        If CharList(i - 1).PosX - 1 >= FinalPosX And CharList(i - 1).PosX + 1 <= FinalPosX And CharList(i - 1).PosY - 1 >= FinalPosY And CharList(i - 1).PosY + 1 <= FinalPosY Then
            Call SetUserTarget((i - 1), True)
            LastCharTargeted = i - 1
            posbyTarget.Enabled = True
            lblTest.Caption = LastCharTargeted
            Exit For
        End If
    End If
Next i

End If
End If

nVal = CBool(GetAsyncKeyState(vbKey3))
'End If
End Sub

