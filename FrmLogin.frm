VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form FrmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4515
   ClientLeft      =   14220
   ClientTop       =   2445
   ClientWidth     =   4485
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmLogin.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Serial2 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Serial 
      Height          =   285
      Left            =   960
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.CheckBox chkSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Guardar datos"
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
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2880
      Top             =   5040
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1560
      Top             =   5160
   End
   Begin VB.TextBox txtPrueba 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   7320
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "           ProloGyX Cheats             "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2655
      Left            =   5400
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   3960
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   5280
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4680
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "informacion (web)"
      Height          =   5655
      Left            =   9720
      TabIndex        =   1
      Top             =   2160
      Width           =   8175
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00808080&
         Height          =   5055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Label Entrar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000010&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
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
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
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
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label cmdVarios 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000010&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label cmdVarios 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000010&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   3000
      Width           =   2175
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Begin Code

Option Explicit
Private strDesktop As String
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
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
 
  
Const APPLICATION As String = "LOGIN"
  
Dim l_User As String
Dim l_Password As String
Dim l_Guardar As String
  

  
'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
  
'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
  
  
'Lee un dato _
-----------------------------
'Recibe la ruta del archivo, la clave a leer y _
 el valor por defecto en caso de que la Key no exista
Private Function Leer_Ini(Path_INI As String, key As String, Default As Variant) As String
  
Dim bufer As String * 256
Dim Len_Value As Long
  
        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
          
        Leer_Ini = Left$(bufer, Len_Value)
  
End Function
  
'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave
  
Private Function Grabar_Ini(Path_INI As String, key As String, Valor As Variant) As String
  
    WritePrivateProfileString APPLICATION, _
                                         key, _
                                         Valor, _
                                         Path_INI
  
End Function

Private Sub cmdVarios_Click(index As Integer)
Select Case index

 Case 0
Dim r As Long
r = ShellExecute(0, "open", "https://discord.gg/vSwBdmH", 0, 0, 1)
 
 Case 1
 End
 
 End Select
End Sub

Private Sub Entrar_Click()
 'On Error GoTo MyErrorHandler
 Dim Resultadox As String
 Dim Path_Archivo_Ini As String
Dim XA As Integer
Dim TextoSeparado() As String
Dim sarna As String
Dim marco As String

'' // Comprobamos si la cuenta o contraseña esstan vacias
    If txtUser.Text = "" Or txtPass.Text = "" Then
    
    MsgBox "Usuario o contraseña vacías."
    Exit Sub
    
    End If
'' //

'' // Cambiamos el texto a mayúsculas para evitar bugs
    txtUser.Text = UCase(txtUser.Text)
    txtPass.Text = UCase(txtPass.Text)
'' //

'' // Ya que no puedo poner interacciones en la url con el complemento INET lo seteo en una variable
UserURL = ("https://prologyx.do.am/Cuentas/" & txtUser.Text & ".txt")
'' //

'' // Seteamos en el complemento INET la URL que anteriormente declaramos como variable
Inet1.URL = UserURL
'' //

'' // Declaramos que Resultadox va a contener la info obtenida de la url

Resultadox = Inet1.OpenURL
'' //




Path_Archivo_Ini = App.Path & "\cfg.ini"

'' // Comprobamos que la cuenta exista
    If InStr(1, Resultadox, "DOCTYPE") Then
        MsgBox "La cuenta no existe", vbExclamation, "ERROR: Posiblemente:"
        Exit Sub
    End If
'' //

'' //  Separamos la informacion que obtuvimos de la url mediante split con "<>" com oseparador - GRACIAS MISERY


TextoSeparado = Split(Resultadox, "<>")



    For XA = LBound(TextoSeparado) To UBound(TextoSeparado)
    
    Next XA
'' //

'' // Multiples comprobaciónes (se entiende)
Serial.Text = TextoSeparado(2)
    If TextoSeparado(1) = "1" Then
        Call UnloadAll
        MsgBox "Se te ha prohibido el uso del cheat"
        End
    End If

    If txtPass.Text = TextoSeparado(0) Then
    
        For XA = 1 To Len(Environ("COMPUTERNAME"))
          marco = Hex(Asc(Mid(Environ("COMPUTERNAME"), XA, 1)) Xor 23)
          sarna = sarna & marco
        Next
        
        Serial2.Text = sarna
        If Serial.Text = Serial2.Text Then
            Else
            'Call UnloadAll
            'MsgBox "NO ERES EL DUEÑO DE LA CUENTA"
            'End
        End If
        
    End If

    If chkSave.Value = 1 Then
        Call Grabar_Ini(Path_Archivo_Ini, "User", txtUser.Text)
        Call Grabar_Ini(Path_Archivo_Ini, "Pass", txtPass.Text)
        Call Grabar_Ini(Path_Archivo_Ini, "Guardar", "SI")
    Else
        Call Grabar_Ini(Path_Archivo_Ini, "Guardar", "NO")
    End If

'' //
AO = "Subastas"
EXEAO = "Tierras Perdidas.exe"
Me.Visible = False
Unload Me
frmHwnd.Show


'' // Si hay un error seguimos acá
   'On Error GoTo 0
   'Exit Sub
'MyErrorHandler:
   'MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl
'' //
    

End Sub

Private Sub Form_Load()
    Randomize
Me.Caption = Int((Rnd * 9) + 1)

Dim pene As String
Dim PATHA As String
App.Title = ""
App.TaskVisible = False


                          PATHA = App.Path & "\cfg.ini"
                            
                                'Path del fichero Ini

      
    ' Lee las Key y  Les envia el valor por defecto por si no existe
    l_User = Leer_Ini(PATHA, "User", "")
    l_Password = Leer_Ini(PATHA, "Pass", "")
    l_Guardar = Leer_Ini(PATHA, "Guardar", "")

If l_Guardar = "SI" Then
chkSave.Value = vbChecked
txtUser.Text = l_User
txtPass.Text = l_Password
End If
End Sub

Private Sub Intervalos_Click()

End Sub


