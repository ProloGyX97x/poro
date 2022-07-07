VERSION 5.00
Begin VB.Form FrmTeclas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3585
   ClientLeft      =   6045
   ClientTop       =   7455
   ClientWidth     =   3840
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Teclas por DEFAULT"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Teclas y Seguir"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Top             =   405
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   690
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   975
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   4
      Top             =   2115
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   2
      Top             =   1830
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   1
      Top             =   1260
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   0
      Top             =   1545
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activate Cheat:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Lanzar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   18
      Top             =   435
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Aim:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Remo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ocultar Cheat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   2130
      Width           =   1410
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Act/Des AutoPota:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   2415
      Width           =   1905
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lanzar un solo H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   1830
      Width           =   1740
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Hechizo 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   1275
      Width           =   1440
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Hechizo 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1440
   End
End
Attribute VB_Name = "FrmTeclas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Begin Code

Option Explicit

Dim Ready As Boolean
Dim LastIndex As Byte
Dim ii As Byte
Dim i As Byte



Private Sub Command1_Click()
Call CustomKeys.LoadDefaults

For ii = 1 To CustomKeys.Count
    Text1(ii).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(ii))
Next ii
Text1(4).Text = "Tab"
End Sub

Private Sub Command2_Click()
For ii = 1 To CustomKeys.Count
    If LenB(Text1(ii).Text) = 0 Then

    End If
Next ii
Call CustomKeys.SaveCustomKeys
FrmMain.Show
Unload Me

End Sub

Private Sub Form_Activate()

For ii = 1 To CustomKeys.Count
    Text1(ii).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(ii))
Next ii
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tLong As Integer


Call CustomKeys.LoadCustomKeys

'Set FrmTeclas = Nothing
End Sub

Private Sub Text1_GotFocus(index As Integer)
Ready = True
End Sub

Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
'If key is not valid, we exit

Text1(index).Text = CustomKeys.ReadableName(KeyCode)
Text1(index).SelStart = Len(Text1(index).Text)

For i = 1 To CustomKeys.Count
    If i <> index Then
        If CustomKeys.BindedKey(i) = KeyCode Then
            Text1(index).Text = "" 'If the key is already assigned, simply reject it
            Call Beep 'Alert the user
            KeyCode = 0
            Exit Sub
        End If
    End If
Next i

CustomKeys.BindedKey(index) = KeyCode
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(index, KeyCode, Shift)
End Sub

Private Sub Text1_LostFocus(index As Integer)
Ready = False
End Sub

Private Sub Text1_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If LastIndex <> index Then
    LastIndex = index
    Exit Sub
End If

If Ready = False Then
'    Ready = True
    Exit Sub
End If

If LenB(CustomKeys.ReadableName(Button)) = 0 Then Exit Sub
'If key is not valid, we exit

Text1(index).Text = CustomKeys.ReadableName(Button)
Text1(index).SelStart = Len(Text1(index).Text)

For i = 1 To CustomKeys.Count
    If i <> index Then
        If CustomKeys.BindedKey(i) = Button Then
            Text1(index).Text = "" 'If the key is already assigned, simply reject it
            Call Beep 'Alert the user
            Button = 0
            Exit Sub
        End If
    End If
Next i

CustomKeys.BindedKey(index) = Button
End Sub

