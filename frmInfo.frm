VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   720
   ClientLeft      =   5820
   ClientTop       =   2910
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtData 
      BackColor       =   &H00151515&
      BorderStyle     =   0  'None
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
      Height          =   735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   8250
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   4320
      Top             =   1920
   End
   Begin VB.CommandButton cmdDat 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   5040
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Begin Code

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '
Private Declare Function SetWindowPos _
    Lib "User32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    

Private Sub cmdDat_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Form_Activate()
frmInfo.Top = RenderTop * 15
frmInfo.Left = RenderLeft * 15

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Timer2_Timer()
If frmInfo.Visible = True Then
If vTeclas = "A" Then
Me.Visible = False
Unload Me
FrmTeclas.Visible = True
vTeclas = "B"
End If

If vIntervalos = "A" Then
Me.Visible = False
Unload Me
FrmIntervalos.Visible = True
vIntervalos = "B"
End If
If vMain = "A" Then
Me.Visible = False
Unload Me
FrmMain.Visible = True
vMain = "B"
End If
If vCoord = "A" Then
Me.Visible = False
Unload Me
frmCoord.Visible = True
vCoord = "B"
End If
Me.Visible = False
Unload Me
End If
End Sub

Private Sub txtData_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub txtData_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If vTeclas = "A" Then
Me.Visible = False
Unload Me
FrmTeclas.Visible = True
vTeclas = "B"
End If

If vIntervalos = "A" Then
Me.Visible = False
Unload Me
FrmIntervalos.Visible = True
vIntervalos = "B"
End If
If vMain = "A" Then
Me.Visible = False
Unload Me
FrmMain.Visible = True
vMain = "B"
End If
If vCoord = "A" Then
Me.Visible = False
Unload Me
frmCoord.Visible = True
vCoord = "B"
End If
End Sub

