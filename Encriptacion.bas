Attribute VB_Name = "Encriptacion"
'Begin Code


Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String
#If Not CASE_SENSITIVE_PASSWORD Then
    strPwd = UCase$(strPwd)
#End If
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function
Public Sub UnloadAll()
'------------------------------------------------------------
' Procedure: UnloadAll
' Purpose: To Unload all forms once the program finishes
' Input parameters: None
'
' Output parameters:None
'
' Return Value:None
' Author: PETER PILUK
' Date: 06/11/00
'------------------------------------------------------------


   Dim f As Integer
   f = Forms.Count
   Do While f > 0
       Unload Forms(f - 1)
       If f = Forms.Count Then Exit Do
       f = f - 1
   Loop
End Sub
Private Function ConvToHex(x As Integer) As String
    If x > 9 Then
        ConvToHex = Chr(x + 55)
    Else
        ConvToHex = CStr(x)
    End If
End Function
Public Function Encriptarr(FrmDataValue As Variant) As Variant
    Dim x As Long
    Dim temp As String
    Dim TempNum As Integer
    Dim TempChar As String
    Dim TempChar2 As String
    For x = 1 To Len(FrmDataValue)
        TempChar2 = Mid(FrmDataValue, x, 1)
        TempNum = Int(Asc(TempChar2) / 16)
        If ((TempNum * 16) < Asc(TempChar2)) Then
            TempChar = ConvToHex(Asc(TempChar2) - (TempNum * 16))
            temp = temp & ConvToHex(TempNum) & TempChar
        Else
            temp = temp & ConvToHex(TempNum) & EnDecryptMUR(EnDecryptMUR(";", "Lxft"), EnDecryptMUR("1C7)0j", "Mt0mA"))
        End If
    Next x
    Encriptarr = temp
End Function
Public Function EnDecryptMUR(sString As String, sPass As String) As String
Dim iLng As Long
Dim i As Integer
    For i = 1 To Len(sString)
        iLng = Asc(Mid$(sPass, ((i Mod Len(sPass)) + 1), 1))
        EnDecryptMUR = EnDecryptMUR & Chr$(Asc(Mid$(sString, i, 1)) Xor iLng)
    Next i
End Function
Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String
#If Not CASE_SENSITIVE_PASSWORD Then
strPwd = UCase$(strPwd)
#End If
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function
Private Function ConvToInt(x As String) As Integer
    Dim X1 As String
    Dim X2 As String
    Dim temp As Integer
    X1 = Mid(x, 1, 1)
    X2 = Mid(x, 2, 1)
    If IsNumeric(X1) Then
        temp = 16 * Int(X1)
    Else
        temp = (Asc(X1) - 55) * 16
    End If
    If IsNumeric(X2) Then
        temp = temp + Int(X2)
    Else
        temp = temp + (Asc(X2) - 55)
    End If
    ConvToInt = temp
End Function
Public Function Desencriptar(FrmDataValue As Variant) As Variant
    Dim x As Integer
    Dim temp As String
    Dim HexByte As String
    For x = 1 To Len(FrmDataValue) Step 2
        HexByte = Mid(FrmDataValue, x, 2)
        temp = temp & Chr(ConvToInt(HexByte))
    Next x
    Desencriptar = temp
End Function
Public Function Cr2(txt As String) As String
Randomize
Dim temp As String
Dim Distorcion As Integer
Dim i As Integer
Distorcion = Int(Rnd * 5)
Distorcion = Distorcion + 100
temp = Distorcion + Asc(Right$(txt, 1)) + Distorcion
For i = 1 To Len(txt)
    temp = temp & (Asc(Mid$(txt, i, 1)) + Distorcion)
Next i
Cr2 = temp
End Function
Public Function Dr1(txt As String) As String
On Error Resume Next
Dim i As Integer
Dim temp As String
Dim Distorcion As Integer
Distorcion = Left$(txt, 3) - Right$(txt, 3)
txt = Right$(txt, Len(txt) - 3)
For i = 1 To (Len(txt) / 3)
    temp = temp & Chr(Mid$(txt, (i * 3) - 2, 3) - Distorcion)
Next i
Dr1 = temp
End Function
Public Function Cr1(ByVal strPassword As String) As String
Dim i As Integer
Dim Char
Dim Palabra As Collection
Cr1 = ""
Set Palabra = New Collection
For i = 1 To Len(strPassword)
Char = Mid(strPassword, i, 1)
Palabra.add Asc(Char) + Asc(Char)
Next i
For Each Char In Palabra
Cr1 = Cr1 & Chr(Char)
Next Char
End Function
Public Function Dr2(ByVal pwdArchi As String) As String
Dim i As Integer
Dim Char
Dim char2
Dim Palabra As Collection
Set Palabra = New Collection
Dr2 = ""
For i = 1 To Len(pwdArchi)
char2 = Mid(pwdArchi, i, 1)
Palabra.add Asc(char2) / 2
Next i
For Each Char In Palabra
Dr2 = Dr2 & Chr(Char)
Next Char
End Function



