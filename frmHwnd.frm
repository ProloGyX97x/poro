VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmHwnd 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView12 
      Height          =   3735
      Left            =   3960
      TabIndex        =   1
      Top             =   4080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6588
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6376
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "asd"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "asdasd"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "asdasdas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "asdasdasd"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmHwnd"
Attribute VB_Exposed = False
'Begin Code


Option Explicit

Public Sub EnumSelectedWindow(sItem As String, hwnd As Long)
   ListView1.ListItems.Clear
   ListView1.ListItems.add , , sItem
   Call EnumChildWindows(hwnd, AddressOf EnumChildProc, &H0)
   
End Sub
Private Sub Command1_Click()
MsgBox ListView1.ListItems(31).SubItems(1)
MsgBox ListView1.ListItems(26).SubItems(1)
MsgBox ListView1.ListItems(43).SubItems(1)

End Sub

Private Sub Form_Activate()
   On Error GoTo MyErrorHandler
Dim hwnd As Long, pid As Long


    PictureRenderhWnd = ListView1.ListItems(31).SubItems(1)
    hwndInventario = ListView1.ListItems(26).SubItems(1)
    hwndHechizos = ListView1.ListItems(43).SubItems(1)
    texthwnd = ListView1.ListItems(46).SubItems(1)

    If hwndexe = 0 Or PictureRenderhWnd = 0 Then
        MsgBox "Fallo"
        End
    End If
Call GetRichTextBoxHwnd(hwndexe)
EsperaMs (100)
Inicial.Caption = ""



If PictureRenderhWnd = 0 Or hwndInventario = 0 Or hwndexe = 0 Then
MsgBox "Fallo en inicio del ejecutalbe, Por favor probar otra vez", vbCritical, Rnd(1000 * 3)
Else
Unload Me
Inicial.Show
FrmMain.Show

End If
Unload Me

   On Error GoTo 0
   Exit Sub
MyErrorHandler:
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl
End Sub

Private Sub Form_Load()
   Dim hwndSelected As Long
 Me.Width = "0"
 Me.Height = "0"
'No pisarse con AO

frmHwnd.Caption = Rnd(100) * 3


Sleep (100)

hwndexe = FindWindow("ThunderRT6FormDC", "")

Sleep (100)

frmHwnd.Caption = ""

'No pisarse con AO
      With ListView1
      .ColumnHeaders.add , , "Window Class or Title"
      .ColumnHeaders.add , , "Handle"
      .ColumnHeaders.add , , "Type"
      .ColumnHeaders.add , , "Class Name"
      .View = lvwReport
   End With
   Call frmHwnd.EnumSelectedWindow("", hwndexe)
   'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub



