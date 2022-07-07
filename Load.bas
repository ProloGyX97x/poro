Attribute VB_Name = "Load"
'Begin Code

Private Declare Function SetWindowPos Lib "user32" (ByVal hnwd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Sub Main()
Dim GuardarSkin As String
Dim archivoSkin() As Byte
Dim GuardarOcx As String
Dim archivoOcx() As Byte
GuardarOcx = App.Path & "/asd.ocx"
archivoOcx = LoadResData(101, "CUSTOM")

GuardarSkin = App.Path & "/Design.prg"
archivoSkin = LoadResData("SKN1", "SKIN")

Open GuardarOcx For Binary As 1#
Put #1, , archivoOcx
Close #1

Open GuardarSkin For Binary As 1#
Put #1, , archivoSkin
Close #1

FrmLogin.Show
End Sub


Public Sub FormPrimerPlano(frm As Form)
Call SetWindowPos(frm.hwnd, -1, 0, 0, 0, 0, &H2 Or &H1)
End Sub

Public Sub FormNoPrimerPlano(frm As Form)
Call SetWindowPos(frm.hwnd, -2, 0, 0, 0, 0, &H2 Or &H1)
End Sub

